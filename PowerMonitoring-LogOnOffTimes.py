"""
Power Monitoring- LogOnOffTimes.py

Description- Continuously monitor AC Line (CH1) and Amp Output (CH2+), logging ON/OFF times (based on AC line voltage).

Equipment- Scope in autorun trigger mode using measurements from Rigol, Tektronix, LeCroy, or Keysight

User must input IP address, number of output load channels to monitor (1-7), if drop-out monitoring required and
if so, the drop-out delay, thresholds for ON/OFF for both AC Line and Amp Output.

Minimum channels to monitor is one, i.e., the line voltage; and if load checks are required, at least one amp 
channel (CH2).  User is given the option for this program to set up the scope on contiguous channels.

Data is saved to CSV file. At the end of test, an Excel file is created from the CSV file.

Author: C. Wong
Last Modified: 20260406
"""

# Standard library
import time
import datetime
import os
import threading
import csv
from collections import deque  #only needed for running average on AC Line
from glob import glob

# Third-party libraries
import pyvisa
import keyboard
from openpyxl import Workbook

# Print the header at runtime.
print(__doc__)

# Global constants and defaults
DEFAULT_IP_ADDRESS = '10.100.52.231' # 10.101.100.151, 169.254.131.118, 192.168.1.90
MAX_LINE_VOLTAGE_VRMS = 350.0       # volts rms limit for AC Line (CH1)
MAX_VRMS = 50.0                     # volts (at 8 ohms that's ~312 W foraudio CHs (CH2+)
LINE_VOLTAGE_WINDOW_SIZE = 4        # Window size for the running average
SETTLING_TIME = 3.0                 # seconds
LOAD_CHECK_INTERVAL = 2             # seconds between checking amplifier channels
LAST_LOAD_CHECK_TIME = 0            # timestamp of the last amplifier load check

# Global control
stop_program_event = threading.Event()
last_quit_attempt = 0               # time of last quit attempt to prevent accidental key presses
line_voltage_readings_queue = deque(maxlen=LINE_VOLTAGE_WINDOW_SIZE)  #running average on line voltage readings

# Find user desktop one level down from home [~/* /Desktop] as optional path to account for OneDrive
DESKTOP_PATH = glob(os.path.expanduser("~\\*\\Desktop"))
default_path = DESKTOP_PATH[0] if DESKTOP_PATH else os.path.expanduser("~\\Desktop")

class Scope:
    """Base class for Oscilloscope communication."""
    def __init__(self, instr, brand):
        self.instr = instr
        self.brand = brand
        self.timeout = 10000

    def setup(self, num_channels, max_channels):
        pass

    def set_trigger(self):
        pass

    def set_trigger_level(self, channel, level):
        pass

    def _query_vrms(self, channel):
        return "0.0"
    
    def stop(self): 
        pass

    def close(self):
        """Cleans up the instrument state and closes the VISA connection."""
        try:
            self.instr.write("CLEAR") 
        except Exception:
            pass 
        self.instr.close()

    def get_measurements(self, num_channels, current_state, check_interval, force_load=False):
        global LAST_LOAD_CHECK_TIME
        measurement_time = datetime.datetime.now()
        readings_all = []
        try:
            raw = self._query_vrms(1)
            v_line = apply_line_voltage_bounds(parse_visa_numeric(raw))
            readings_all.append(v_line)
            line_voltage_readings_queue.append(v_line)
            v_line_avg = sum(line_voltage_readings_queue)/len(line_voltage_readings_queue)
            if current_state == "UNKNOWN" or force_load or (time.time() - LAST_LOAD_CHECK_TIME > check_interval):
                LAST_LOAD_CHECK_TIME = time.time()
                for i in range(2, num_channels + 1):
                    r = self._query_vrms(i)
                    readings_all.append(max(min(parse_visa_numeric(r), MAX_VRMS), 0)) 
            else:
                readings_all.extend([None] * (num_channels - 1)) 
            return measurement_time, readings_all, v_line_avg 
        except pyvisa.errors.VisaIOError:
            return None, None, None
        
class TekScope(Scope):
    def setup(self, num_channels, max_channels):
        print("Setting up Tektronix scope...")
        reset_request = input("Reset scope? (Y/N): ").strip().lower()
        if reset_request == 'y':
            self.instr.write("*RST"); time.sleep(5); self.instr.write("*CLS"); time.sleep(2)
        for i in range(1, num_channels + 1):
            self.instr.write(f"SELect:CH{i} ON")
            self.instr.write(f"CH{i}:POSition 0")
            scale = 100 if i == 1 else 10
            self.instr.write(f"CH{i}:SCALe {scale}")
            self.instr.write(f"MEASUrement:MEAS{i}:SOUrce CH{i}; STATE 1")
            self.instr.write(f"MEASUrement:MEAS{i}:SOUrce CH{i}; TYPE RMS")
        self.instr.write("HORizontal:SCAle 200E-6")
        self.instr.write("HORizontal:POSition 50")
        self.instr.write("TRIGger:A:EDGE:SOUrce CH1")
        self.instr.write("TRIGger:A:EDGE:COUPling DC")
        self.instr.write("TRIGger:A:EDGE:SLOpe RISE")
        self.instr.write("TRIGger:A:LEVel:CH1 50")
        time.sleep(1)

    def set_trigger(self):
        self.instr.write("ACQuire:STATE ON")
        return True

    def set_trigger_level(self, channel, level):
        self.instr.write(f"TRIGger:A:LEVel:CH{channel} {level}")
        self.instr.query("*OPC?")

    def _query_vrms(self, channel):
        return self.instr.query(f"MEASUrement:MEAS{channel}:VALue?")

    def stop(self):
        self.instr.write("ACQuire:STATE OFF")

class RigolScope(Scope):
    def setup(self, num_channels, max_channels):
        reset_request = input("Reset scope? (Y/N): ").strip().lower()
        if reset_request == 'y':
            self.instr.write("*RST"); time.sleep(5); self.instr.write("*CLS"); time.sleep(2)
        for i in range(1, max_channels + 1):
            self.instr.write(f":CHANnel{i}:DISPlay OFF")
        for i in range(1, num_channels + 1):
            scale = 100 if i == 1 else 10
            self.instr.write(f":CHANnel{i}:SCALe {scale}")
            self.instr.write(f":CHANnel{i}:DISPlay ON")
            self.instr.write(f":CHANnel{i}:PROBe 10")
            self.instr.write(f":CHANnel{i}:OFFSet 0")
            self.instr.write(f":CHANnel{i}:BWLimit 20M")
            self.instr.write(f":CHANnel{i}:COUPling DC")
            self.instr.write(f":CHANnel{i}:INVert OFF")
            self.instr.write(f":CHANnel{i}:UNITs VOLT")
        self.instr.write(":TIMebase:SCALe 200e-6")
        self.instr.write(":TIMebase:DELay 50")
        self.instr.write(":TRIGger:MODE EDGE")
        self.instr.write(":TRIGger:EDGe:SOUrce CHAN1")
        self.instr.write(":TRIGger:EDGe:COUPling DC")
        self.instr.write(":TRIGger:EDGe:SLOpe POSitive")
        self.instr.write(":TRIGger:EDGe:LEVel 50")
        time.sleep(1)

    def set_trigger(self):
        self.instr.write(":RUN")
        return True

    def set_trigger_level(self, channel, level):
        self.instr.write(f":TRIGger:EDGe:SOUrce CHAN{channel}")
        self.instr.write(f":TRIGger:EDGe:LEVel {level}")
        self.instr.query("*OPC?")

    def _query_vrms(self, channel):
        return self.instr.query(f":MEASure:VRMS? CHAN{channel}")

    def stop(self):
        self.instr.write(":STOP")

class LeCroyScope(Scope):
    def setup(self, num_channels, max_channels):
        print("Be sure to set scope, utilities to TCPIP (VXI-11). Setting up LeCroy scope...")
        reset_request = input("Reset scope? (Y/N): ").strip().lower()
        if reset_request == 'y':
            self.instr.write("*RST"); time.sleep(5); self.instr.write("*CLS"); time.sleep(2)
        self.instr.write("VBS 'app.Measure.ClearAll'")
        self.instr.write("VBS 'app.Measure.ShowMeasure = True'")
        for i in range(1, max_channels + 1):
            self.instr.write(f"VBS 'app.Acquisition.C{i}.View = False'")
        for i in range(1, num_channels + 1):
            label = "AC Power Line" if i == 1 else f"Amp Out {i-1}"
            scale = 100 if i == 1 else 10
            vertical_settings = [
                f"app.Acquisition.C{i}.VerOffset = 0",
                f"app.Acquisition.C{i}.Coupling = \"DC1M\"",
                f"app.Acquisition.C{i}.View = True",
                f"app.Acquisition.C{i}.ViewLabels = True",
                f"app.Acquisition.C{i}.BandwidthLimit = \"20MHz\"",
                f"app.Acquisition.C{i}.VerScale = {scale}",
                f"app.Acquisition.C{i}.LabelsText = \"{label}\"",
                f"app.Measure.P{i}.ParamEngine = \"RootMeanSquare\"",
                f"app.Measure.P{i}.Operator.Cyclic = \"True\"",
                f"app.Measure.P{i}.Source1 = \"C{i}\"",
                f"app.Measure.P{i}.View = True"
            ]
            for cmd in vertical_settings: self.instr.write(f"VBS '{cmd}'")
        horizontal_settings = [
            "app.Acquisition.Horizontal.HorScale = 5e-3",
            "app.Acquisition.Horizontal.HorOffset = 0",
            "app.Acquisition.Trigger.Type = \"Edge\"",
            "app.Acquisition.Trigger.Edge.Source = \"C1\"",
            "app.Acquisition.Trigger.Edge.Level = 60",
            "app.Acquisition.TriggerMode = \"Auto\""
        ]
        for cmd in horizontal_settings: self.instr.write(f"VBS '{cmd}'")
        time.sleep(1)

    def set_trigger(self):
        self.instr.write("VBS 'app.Acquisition.TriggerMode = \"Auto\"'")
        return True

    def set_trigger_level(self, channel, level):
        self.instr.write(f"VBS 'app.Acquisition.Trigger.Edge.Source = \"C{channel}\"'")
        self.instr.write(f"VBS 'app.Acquisition.Trigger.Edge.Level = {level}'")
        self.instr.query("*OPC?")

    def _query_vrms(self, channel):
        return self.instr.query(f"VBS? 'return=app.Measure.P{channel}.Out.Result.Value'")

    def stop(self):
        self.instr.write(":STOP")

class KeysightScope(Scope):
    def setup(self, num_channels, max_channels):
        reset_request = input("Reset scope? (Y/N): ").strip().lower()
        if reset_request == 'y':
            self.instr.write("*RST"); time.sleep(5); self.instr.write("*CLS"); time.sleep(2)
        self.instr.write(":MEASure:CLEar")
        for i in range(1, max_channels + 1):
            self.instr.write(f":CHANnel{i}:DISPlay OFF")
        for i in range(1, num_channels + 1):
            scale = 100 if i == 1 else 10
            self.instr.write(f":CHANnel{i}:SCALe {scale}")
            self.instr.write(f":CHANnel{i}:DISPlay ON")
            self.instr.write(f":CHANnel{i}:PROBe 10")
            self.instr.write(f":CHANnel{i}:OFFSet 0")
            self.instr.write(f":CHANnel{i}:BWLimit ON")
            self.instr.write(f":CHANnel{i}:COUPling DC")
            self.instr.write(f":CHANnel{i}:INVert OFF")
            self.instr.write(f":CHANnel{i}:UNITs VOLT")
        self.instr.write(":TIMebase:RANGe 200e-6")
        self.instr.write(":TIMebase:POSition 0")
        self.instr.write(":TRIGger:MODE EDGE")
        self.instr.write(":TRIGger:EDGE:SOURce CHANnel1")
        self.instr.write(":TRIGger:EDGE:COUPling DC")
        self.instr.write(":TRIGger:EDGE:SLOPe POSitive")
        self.instr.write(":TRIG:EDGE:SOUR CHAN1;LEVel 50")
        time.sleep(1)

    def set_trigger(self):
        self.instr.write(":RUN")
        return True

    def set_trigger_level(self, channel, level):
        self.instr.write(f":TRIGger:EDGE:SOUR CHAN1;LEVel {level}")
        self.instr.query("*OPC?")

    def _query_vrms(self, channel):
        return self.instr.query(f":MEASure:VRMS? CHANnel{channel}")

    def stop(self):
        self.instr.write(":STOP")

def on_q_press():
    """
    Deliberate quit logic: Requires two presses of 'q' within 2 seconds.
    """
    global last_quit_attempt
    current_time = time.time()
    
    if current_time - last_quit_attempt < 2.0:
        print("\nConfirmation received. Signaling program to stop...")
        stop_program_event.set()
    else:
        print("\n[QUIT ATTEMPT] Press 'q' again within 2 seconds to confirm exit.")
        last_quit_attempt = current_time

def on_esc_press():
    """
    Callback function when 'Esc' is pressed.
    """
    print("You pressed Esc!")
    stop_program_event.set()

def connect_to_instrument(resource_manager: pyvisa.ResourceManager, default_ip: str = DEFAULT_IP_ADDRESS):
    """
    Prompts the user for an IP address and attempts to establish a connection
    to a PyVISA instrument and returns the object plus its identity.

    Args:
        resource_manager: The PyVISA ResourceManager instance.
        default_ip: The default IP address to suggest to the user.

    Returns:
        The connected PyVISA instrument object and brand label (string).
    """
    while True:
        user_input = input(f"Enter IP address or 'd' for default (Default {default_ip}): ").strip()
        
        if user_input.lower() == 'q':
            return None, None
        
        # Use default if input is 'd' or empty
        if user_input.lower() == 'd' or not user_input:
            visa_address = default_ip
        else:
            visa_address = user_input

        # Smart string construction
        if "::" in visa_address:
            resource_string = visa_address
        else:
            resource_string = f'TCPIP::{visa_address}::INSTR'

        print(f"Attempting connection to: {resource_string}...")

        try:
            instr = resource_manager.open_resource(resource_string)
            instr.timeout = 10000 
            idn = instr.query('*IDN?').strip().upper()
            
            # Instantiate the appropriate Scope subclass
            if "RIGOL" in idn:
                scope_obj = RigolScope(instr, "rigol")
            elif "TEKTRONIX" in idn:
                scope_obj = TekScope(instr, "tek")
            elif "LECROY" in idn:
                scope_obj = LeCroyScope(instr, "lecroy")
            elif "KEYSIGHT" in idn:
                scope_obj = KeysightScope(instr, "keysight")
            else:
                scope_obj = RigolScope(instr, "other")

            print(f"Connected! Identity: {idn}")
            return scope_obj, scope_obj.brand

        except (pyvisa.errors.VisaIOError, Exception) as e:
            print(f"Error: {e}")
            # Ensure we don't leave a half-open connection
            if 'instr' in locals():
                try:
                    instr.close()
                except:
                    pass
            print("Retrying... (Press Ctrl+C to stop)")
            time.sleep(1)

def get_max_channels():
    """
    Asks the user to input the physical maximum number of channels (2-8) on scope.
    'd' or Enter returns the default.
    """
    user_input = input("Enter the TOTAL physical maximum number of channels (2-8) or 'd' for default = 4): ").strip().lower()

    # 1. Handle the explicit default cases
    if user_input == 'd' or user_input == '':
        return 4

    # 2. Try to process as an integer
    try:
        max_channels = int(user_input)
        
        # Check if it's within your specific 2-8 range
        if 2 <= max_channels <= 8:
            return max_channels
        else:
            print(f"Value {max_channels} is out of range. Using default.")
            return 4
            
    except ValueError:
        print("Invalid input format. Using default.")
        return 4

def get_num_channels(max_channels):
    """
    Asks the user to input the number of load channels to monitor (2-8).
    Note- First channel (CH1) monitors AC Line. Additional channels (CH2+) monitor
    amplifier outputs (CH2+).
    """
    print("First channel monitors AC Line. Additional channels monitor amplifier outputs.")
    while True:
        try:
            num_channels = int(input(f"Including AC Line (CH1), enter total number of CHs (Enter 2-{max_channels}): "))
            if 2 <= num_channels <= max_channels:
                return num_channels
            else:
                print(f"Invalid input. Please enter a number between 2 and {max_channels}.")
        except ValueError:
            print("Invalid input. Please enter a number.")

def get_thresholds(default_on_value, default_off_value, max_limit=300, min_limit=0):
    """
    Prompts user for ON/OFF Vrms thresholds with dynamic limits and defaults.
    """
    # DEVELOPER SANITY CHECK:
    # Ensures the programmer didn't pass defaults that violate the limits.
    if not (min_limit <= default_off_value < default_on_value <= max_limit):
        print(f"DEBUG WARNING: Defaults ({default_on_value}, {default_off_value}) "
              f"are outside limits ({min_limit}, {max_limit})!")

    while True:
        try:
            # ON Threshold Prompt
            on_prompt = (f"  Enter ON rms level (Default: {default_on_value:.2f}V, "
                         f"Min: {min_limit}V, Max: {max_limit}V, 'd' for default): ")
            user_on_input = input(on_prompt).strip().lower()
            
            on_rms_level = default_on_value if user_on_input == 'd' else float(user_on_input)

            # OFF Threshold Prompt (Added min_limit here for UI consistency)
            off_prompt = (f"  Enter OFF rms level (Default: {default_off_value:.2f}V, "
                          f"Min: {min_limit}V, 'd' for default): ")
            user_off_input = input(off_prompt).strip().lower()
            
            off_rms_level = default_off_value if user_off_input == 'd' else float(user_off_input)

            # --- VALIDATION LOGIC ---
            if on_rms_level < min_limit:
                print(f"  Error: ON threshold cannot be below {min_limit}V.")
            elif on_rms_level > max_limit:
                print(f"  Error: ON threshold cannot exceed {max_limit}V.")
            elif off_rms_level < min_limit:
                print(f"  Error: OFF threshold cannot be below {min_limit}V.")
            elif off_rms_level >= on_rms_level:
                print(f"  Error: OFF threshold must be less than the ON threshold (OFF < ON).")
            else:
                return on_rms_level, off_rms_level

        except ValueError:
            print("Invalid input. Please enter a numerical value or 'd'.")

def get_power_on_delay():
    """
    Prompts user for the power-on delay in seconds before checking load channels.
    Default is 10 seconds. Minimum is 5 second.
    """
    while True:
        try:
            user_input = input("Enter power-on delay in seconds before checking load channels ('d' for default: 10s, Min: 5s): ").strip().lower()
            if user_input == 'd' or user_input == '':
                return 10
            delay = float(user_input)
            if delay < 5:
                print("Power-on delay must be at least 5 second.")
            else:
                return delay
        except ValueError:
            print("Invalid input. Please enter a numerical value or 'd'.")

def make_datafile(timestamp, dropout_enabled,dropout_interval, path=default_path):
    """
    Generates a csv data file for the data based on start date and time.
    Includes headers for each monitored channel. Asks user for directory
    or defaults to desktop.

    Returns tuple for user_path and data_log_file_name
    """
    header_list = [
        "Event_Count", 
        "Start_Time_Absolute", 
        "End_Time_Absolute", 
        "Line Voltage", 
        "State", 
        "Duration_Seconds"
    ]
    if dropout_enabled:
        header_list.append(f"Drop-out Check ({dropout_interval} sec)")
    header_string = ",".join(header_list)

    user_path_input = input(f"Enter data path or 'd' for desktop: ").strip()
    user_path = path if user_path_input.lower() == 'd' or not user_path_input else user_path_input
    filename = timestamp.strftime("%Y%m%d_%H%M%S.csv")
    full_path = os.path.join(user_path, filename)
    with open(full_path, "w") as f:
        f.write(header_string + "\n")

    return user_path, filename

def apply_line_voltage_bounds(number: float) -> float:
    """
    Applies upper and lower (0) bounds to the Vrms readings for the AC Line (CH1).
    """
    return max(min(number, MAX_LINE_VOLTAGE_VRMS), 0)

def apply_amp_output_bounds(number: float) -> float:
    """
    Applies upper and lower (0)bounds to the Vrms readings 2nd CH+   Note- NOT AC line (CH 1).
    """
    return max(min(number, MAX_VRMS), 0)

def parse_visa_numeric(response):
    """
    Extracts the numeric value from a VISA response string.
    Handles 'VBS 123.4', '123.4\n', or scientific notation '1.234E+02'.
    """
    try:
        if not response:
            return 0.0
        # Split by whitespace and take the last element
        clean_val = response.strip().split()[-1]
        return float(clean_val)
    except (ValueError, IndexError):
        return 0.0

def get_dropout_settings():
    """
    Prompts user for drop-out monitoring configuration.
    """
    check_dropout = input("Enable Signal Drop-out checking? (Y/N, 'd' = Yes): ").strip().lower() != 'n'
    
    interval = 2
    delay = 10
    
    if check_dropout:
        while True:
            val = input("Enter interval for checking load signals (2-300 sec, 'd' for default: 2): ").strip()
            
            if val == 'd' or val == '':
                interval = 2
                break

            try:
                interval = int(val) if val else 2
                if 2 <= interval <= 300: 
                    break
                print("Out of range (2-300).")
            except ValueError:
                print("Invalid entry. Please enter a number or 'd'.")
            
        while True:
            val = input("Enter wait or delay time (after AC Line is ON) before checking load signals (5-300 sec or 'd' for default = 10): ").strip()
            
            if val == 'd' or val == '':
                delay = 10
                break

            try:
                delay = int(val) if val else 10
                if 5 <= delay <= 300: 
                    break
                print("Out of range (5-300).")
            except ValueError: 
                print("Invalid entry. Please enter a number or 'd'.")
            
    return check_dropout, interval, delay

def log_event(path, filename, count, start_time, end_time, line_v, state, duration, label=""):  
    # duration is in seconds formatted to 3 decimal places. If unable log as string.
    try:
        full_path = os.path.join(path, filename)
        try:
            # Try formatting duration as a float with 3 decimal places
            duration_str = f"{float(duration):9.3f}"
        except (ValueError, TypeError):
            # If formatting fails, log as string
            duration_str = f"{str(duration):>9}"

        with open(full_path, "a") as f:
            line = (f"{count},{start_time.strftime('%Y-%m-%d %H:%M:%S.%f')}, "
                   f"{end_time.strftime('%Y-%m-%d %H:%M:%S.%f')},{line_v:.3f}, "
                   f"{state},{duration_str},{label}")
            f.write(line + "\n")


    except IOError as e:
        print(f"Error appending data to file '{full_path}': {e}")

def write_to_excel(filename, path): 
    full_path_csv = os.path.join(path, filename)
    full_path_xlsx = os.path.join(path, os.path.splitext(filename)[0] + ".xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "Power Monitoring"
    try:
        with open(full_path_csv, 'r') as f:
            reader = csv.reader(f)
            for r_idx, row in enumerate(reader, 1):
                processed = []
                for c_idx, val in enumerate(row):
                    if r_idx > 1 and c_idx in [0, 3, 5]:   
                        try: processed.append(float(val))
                        except: processed.append(val)
                    else: processed.append(val)
                ws.append(processed)
        wb.save(full_path_xlsx)
    except Exception as e: print(f"Excel Error: {e}")


# ************** MAIN
rm = None
scope = None
start_time = datetime.datetime.now()
last_state_time = datetime.datetime.now()
event_counter = 0
current_state = "UNKNOWN"
user_path = None
datafile_name = None
first_transition_logged = False
steady_state_line_voltage = 0.0

try:
    # Register the 'q' hotkey
    keyboard.add_hotkey('q', on_q_press)
    keyboard.add_hotkey('esc', on_esc_press)

    # Initialize Resource Manager and check connection to instrument
    rm = pyvisa.ResourceManager()
    scope, brand = connect_to_instrument(rm, default_ip=DEFAULT_IP_ADDRESS)
    if scope is None:
        print("Failed to connect to the instrument. Exiting.")
        exit() # Exit if connection failed
    if brand == "other":
        print("Connected to an unsupported instrument. Proceed with caution.")

    # Get maximum number of channels on this scope model
    max_ch_on_scope = get_max_channels()      # Some scopes have 2, 4, or 8 CH.

    # Get the number of channels from user being hooked up and monitoring. 
    num_channels_to_monitor = get_num_channels(max_ch_on_scope)

    # Get AC Line threshold levels
    print("AC Line ON/OFF thresholds:")
    ac_line_high_limit, ac_line_low_limit = get_thresholds(
        default_on_value=80, 
        default_off_value=70, 
        max_limit=275, 
        min_limit=1
    )
    print("AC Line Vrms ON = ", ac_line_high_limit, ", AC Line Vrms OFF = ", ac_line_low_limit)

    # Get Amp Output threshold levels
    print("Amp Output ON/OFF thresholds:")
    amp_high_limit, amp_low_limit = get_thresholds(
        default_on_value=7.0, 
        default_off_value=2.0, 
        max_limit=20.0,
        min_limit=1
    )
    print("Amp Output Vrms ON = ", amp_high_limit, ", Amp Output Vrms OFF = ", amp_low_limit)

    # Get Dropout configuration
    do_enabled, do_interval, do_delay = get_dropout_settings()

    # Set up scope?
    setup_needed = input("(L)eave scope alone or (S)etup SCOPE CHANNELS (contiguous): ").strip()
    if setup_needed.lower() == 'l':
        print("Skipping scope setup.")
    else:
        print("Attempting to set up scope.. .")
        scope.setup(num_channels_to_monitor, max_ch_on_scope)
        print("Setup complete.")

    # Trigger mode auto or normal?
    run_mode = input("(L)eave scope alone or (S)et up TRIGGER in auto or normal mode? : ").strip().lower()
    if run_mode == 's':
        print("Setting trigger mode.")
        scope.set_trigger()
        print("Trigger set up.")
    else: 
        print("Skipping trigger setup.")

    # Create a data file for logging based on the current timestamp. This time/data log be used for duration of test.
    start_time = datetime.datetime.now()
    user_path, datafile_name = make_datafile(start_time, do_enabled,do_interval)
    full_data_path = os.path.join(user_path, datafile_name)
    print("Created file ", datafile_name, " in path: ", user_path)

    # Notify user ready to start. Provide instructions for stopping the program and ensuring proper initial conditions.
    print(f"Monitoring AC Line voltage > {ac_line_high_limit:.2f} Vrms ON and < {ac_line_low_limit:.2f} Vrms OFF.")
    print("Verify scope settings are acceptable.\nPress 'Crtl-C' to stop the program.")  # q twice if using keyboard hotkey method and exe is run with admin privileges.
    input("Hit Enter to start monitoring...")

    # ************** MAIN LOOP  *****************     
    while not stop_program_event.is_set():
        time.sleep(0.05) # Small delay for keyboard input before checking if scope triggered

        # Get measurements from scope.
        meas_time, all_readings, avg_line = scope.get_measurements(
            num_channels_to_monitor, current_state, do_interval
        )
        # Check if measurement time failed and skip this cycle.
        if meas_time is None: continue

        ac_line_voltage = all_readings[0]
        new_amp_data = all_readings[1:] 

        # Assign state Booleans from ac line (instantaneous)
        ac_line_on, ac_line_off = ac_line_voltage >= ac_line_high_limit, ac_line_voltage <= ac_line_low_limit     

        # Update steady state voltage logic
        if current_state == "ON":
            elapsed = (meas_time - last_state_time).total_seconds()
            if elapsed < SETTLING_TIME:
                steady_state_line_voltage = max(steady_state_line_voltage, ac_line_voltage)
            elif ac_line_on: steady_state_line_voltage = avg_line

            # Periodic Dropout Check and Log
            if do_enabled and None not in new_amp_data:
                if elapsed > do_delay:
                    if not all(v >= amp_high_limit for v in new_amp_data):
                        event_counter += 1
                        min_amp_out = min(new_amp_data)
                        # Drop-out records meas_time twice (both the start and end time) to note 'instance' of event
                        log_event(user_path, datafile_name, event_counter, meas_time, meas_time, avg_line, "ON", "'-", f"* Amp Out MIN: {min_amp_out:.3f}Vrms")
                        print(f"[{meas_time.strftime('%H:%M:%S')}] DROP-OUT DETECTED!  "
                              f"Line Voltage: {avg_line:.3f}Vrms, "
                              f"Amp Out MIN: {min_amp_out:.3f}Vrms)")
                                    
        # State Establishment/Transition Logic
        if current_state == "UNKNOWN":
            if ac_line_on:
                current_state = "ON"
                print(f"Initial state detected as ON. Waiting for next transition.")
                start_time = meas_time
                last_state_time = meas_time
                scope.set_trigger_level(1, ac_line_low_limit)
            elif ac_line_off:
                current_state = "OFF"
                print(f"Initial state detected as OFF. Waiting for next transition.")
                start_time = meas_time
                scope.set_trigger_level(1, ac_line_high_limit)
            else:
                # Still in an indeterminate state or no clear ON/OFF. Keep current_state as UNKNOWN.
                pass # No change, continue will be called below
            continue # Always continue if still in UNKNOWN or just established initial state

        # Current_state is either 'ON' or 'OFF'. Proceed to check for transitions.
        new_actual_state = current_state # Assume no change unless clear transition

        # Determine if there's a *real* transition from the established state
        if current_state == "OFF" and ac_line_on:
            new_actual_state = "ON"
        elif current_state == "ON" and ac_line_off:
            new_actual_state = "OFF"

        # This handles the intermediate or stable state where no clear transition occurs.
        if new_actual_state != current_state:
            if not first_transition_logged:
                # This is the very first transition detected. Don't log the initial state.
                print(f"[{meas_time.strftime('%H:%M:%S')}] OFF. Waiting for first transition to ON.")
                current_state = new_actual_state
                start_time = meas_time
                first_transition_logged = True
                # Set the trigger level for the next transition depending on state
                if current_state == "ON":       # State is ON, set low limit
                    scope.set_trigger_level(1, ac_line_low_limit)
                else:                           # State is OFF, set high limit
                    scope.set_trigger_level(1, ac_line_high_limit)
            else: 
                # This is a subsequent transition; start logging from this pointon
                duration = (meas_time - start_time).total_seconds()
                event_counter += 1
                if current_state == "OFF" and new_actual_state == "ON":
                    # Transition from OFF to ON  (log the previous OFF state, reset timers, line voltage, off trigger)
                    log_event(user_path, datafile_name, event_counter, start_time, meas_time, 0.0, "OFF", duration)
                    print(f"[{meas_time.strftime('%H:%M:%S')}] OFF to ON.  OFF duration was: {duration:.3f} seconds.")
                    current_state = "ON"
                    start_time = meas_time
                    last_state_time = meas_time
                    steady_state_line_voltage = 0.0
                    scope.set_trigger_level(1, ac_line_low_limit)

                elif current_state == "ON" and new_actual_state == "OFF":
                    # Transition from ON to OFF
                    # Determine the line voltage to log
                    voltage_to_log = steady_state_line_voltage

                    log_event(user_path, datafile_name, event_counter, start_time, meas_time, voltage_to_log, "ON", duration)
                    print(f"[{meas_time.strftime('%H:%M:%S')}] ON to OFF.  ON duration was: {duration:.3f} seconds. Line Voltage: {voltage_to_log:.3f}Vrms")
                    current_state = "OFF"
                    start_time = meas_time
                    # Set trigger for next OFF detection
                    scope.set_trigger_level(1, ac_line_high_limit)

except KeyboardInterrupt:
    print("\nProgram terminated by user (Ctrl+C).")
except Exception as e:
    print(f"An error occurred during program execution: {e}")
finally:
    # Close the instrument connection and resource manager
    if scope:
        # Stop acquisition before closing
        try:
            scope.stop()
            scope.close()
            print("Instrument connection closed.")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error closing instrument connection: {e}")
    if rm:
        try:
            rm.close()
            print("Resource Manager closed.")
        except Exception as e:
            print(f"Error closing Resource Manager: {e}")

    # Log last entry to log file and final state if it was not already logged
    event_counter += 1
    final_time = datetime.datetime.now()
    duration = (final_time - start_time).total_seconds()
    final_voltage_to_log = 0.000 # Default to 0.0

    # Only attempt to use current_line_voltage_snapshot if the queue has data
    if len(line_voltage_readings_queue) > 0:
        final_voltage_to_log = sum(line_voltage_readings_queue) / len(line_voltage_readings_queue)

    if current_state == "OFF":
        # Always log 0.0 for line voltage if the final state is OFF
        if user_path and datafile_name:
            log_event(user_path, datafile_name, event_counter, start_time, final_time, final_voltage_to_log, current_state, duration)
        print(f"Program stopped. Final {current_state} duration: {duration:.3f} seconds. Line Voltage hardcoded for OFF state.)")
    else: # If the final state was ON
        # Apply the same hard-code logic for the final ON state if event_counter is 0 or 1
        # This handles cases where the program exits very quickly after starting
        if event_counter == 0 or (event_counter == 1 and duration < 1.0): # event_counter 0 means no transitions logged. 1 means the initial state captured.
             final_voltage_to_log = 0.0 # Hardcode as too  early in program for accurate readings
             print("Note: Line voltage set 0.0V due to early program termination or initial state.")
        if user_path and datafile_name:
            log_event(user_path, datafile_name, event_counter, start_time, final_time, final_voltage_to_log, current_state, duration)
        print(f"Program stopped. Final {current_state} duration: {duration:.3f} seconds. Line Voltage: {final_voltage_to_log:.3f}Vrms")

    # Create Excel file from log file.
    if user_path and datafile_name:
        print("Creating mirror Excel data file...")
        write_to_excel(datafile_name, user_path)

    input("\nExecution complete. Press Enter to exit...")
    