"""
Power Monitoring- LogOnOffTimes.py

Description: 
Continuously monitor AC Line (CH1) and Amplifier output channels (CH2+) 
to log ON/OFF times based on CH1 trigger level.

Equipment- Rigol, Tektronix, LeCroy, or Keysight scope in autorun trigger 
mode using measurements (like a DMM or Power line Logger).
Note- LeCroy implemented with VBS and can be slightly slower than standard
SCPI commands.

User must input IP address and number of output load channels to monitor 
(1-7).  Maximum is total phyiscal channesls (8).  Minimum channels to monitor
is 2, i.e., the line voltage and at least one amplifier channel (CH2). 

If the line voltage is present for longer then 10 seconds, 
all amplifier channels must be above threshold or time is reported.

User is given the option to automatically set up the scope on contiguous
channels or to leave setup alone. User must set up AC line on first 
measurement and amplifier channels on subsequent channels.

Data is saved to csv file at the user directed path or defaults to the 
desktop.

Author: C. Wong
Last Modified: 20260402
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

DEFAULT_IP_ADDRESS = '10.100.52.231' # 10.101.100.151, 169.254.131.118, 192.168.1.90
MAX_LINE_VOLTAGE_VRMS = 350.0       # volts rms limit for AC Line (CH1)
MAX_VRMS = 50.0                     # volts (at 8 ohms that's ~312 W foraudio CHs (CH2+)
LINE_VOLTAGE_WINDOW_SIZE = 4        # Window size for the running average
SETTLING_TIME = 3.0                 # seconds
LOAD_CHECK_INTERVAL = 2             # seconds between checking amplifier channels
LAST_LOAD_CHECK_TIME = 0            # timestamp of the last amplifier load check

user_path = None                    # user directory for data file
datafile_name = None                # data file name for logging durations
last_quit_attempt = 0               # time of last quit attempt to prevent accidental key presses

# Find user desktop one level down from home [~/* /Desktop] as optional path
desktop_path = glob(os.path.expanduser("~\\*\\Desktop"))
# If not found, revert to the standard location (~/Desktop)
if len(desktop_path) == 0:
    desktop_path = os.path.expanduser("~\\Desktop")
else:
    desktop_path = desktop_path[0] # glob returns a list, take the first result

# Global flag to signal the main loop and threads to stop
stop_program_event = threading.Event()

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
        user_input = input(f"Enter IP address, 'q' twice to quit, 'd' for default (Default {default_ip}): ").strip()
        
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
            
            # Create a label/metadata tag
            label = "unknown"
            if "RIGOL" in idn: label = "rigol"
            elif "TEKTRONIX" in idn: label = "tek"
            elif "LECROY" in idn: label = "lecroy"
            elif "KEYSIGHT" in idn: label = "keysight"
            else: label = "other"

            print(f"Connected! Identity: {idn}")
            return instr, label

        except (pyvisa.errors.VisaIOError, Exception) as e:
            print(f"Error: {e}")
            # Ensure we don't leave a half-open connection
            if 'instr' in locals():
                try:
                    instr.close()
                except:
                    pass
            print("Retrying... (Press 'q' twice to stop)")
            time.sleep(1)

def get_max_channels():
    """
    Asks the user to input the physical maximum number of channels (2-8) on scope.
    'd' or Enter returns the default of 4.
    """
    user_input = input("Enter the TOTAL physical maximum number of channels (2-8) or 'd' (Default = 4): ").strip().lower()

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
            print(f"Value {max_channels} is out of range. Using default (4).")
            return 4
            
    except ValueError:
        print("Invalid input format. Using default (4).")
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
            on_prompt = (f"Enter ON rms level (Default: {default_on_value:.2f}V, "
                         f"Min: {min_limit}V, Max: {max_limit}V, 'd' for default): ")
            user_on_input = input(on_prompt).strip().lower()
            
            on_rms_level = default_on_value if user_on_input == 'd' else float(user_on_input)

            # OFF Threshold Prompt (Added min_limit here for UI consistency)
            off_prompt = (f"Enter OFF rms level (Default: {default_off_value:.2f}V, "
                          f"Min: {min_limit}V, 'd' for default): ")
            user_off_input = input(off_prompt).strip().lower()
            
            off_rms_level = default_off_value if user_off_input == 'd' else float(user_off_input)

            # --- VALIDATION LOGIC ---
            if on_rms_level < min_limit:
                print(f"Error: ON threshold cannot be below {min_limit}V.")
            elif on_rms_level > max_limit:
                print(f"Error: ON threshold cannot exceed {max_limit}V.")
            elif off_rms_level < min_limit:
                print(f"Error: OFF threshold cannot be below {min_limit}V.")
            elif off_rms_level >= on_rms_level:
                print(f"Error: OFF threshold must be less than the ON threshold (OFF < ON).")
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

def setup_scope(scope, num_channels, max_channels, brand):
    """
    Configures channels for the specified number of channels, setting
    scale and position, and adapting to brand- rigol, tek, lecroy, keysight, other.
    """
    # DEVELOPER SANITY CHECK:
    # Ensure programmer brand is valid
    if brand.lower() not in ["rigol", "tek", "lecroy", "keysight", "other"]:
        print(f"DEBUG WARNING: Invalid brand specified ({brand}). Please use 'rigol', 'tek', 'lecroy', or 'other'.")
        stop_program_event.set()
        return
    
    reset_request = input("Would you like a scope reset? (Y/N)? ").strip().lower()
    if reset_request == 'y':
        scope.write("*RST")
        print("Scope reset command sent. Please wait for the scope to reset before proceeding.")
        time.sleep(5)  # Wait for the scope to reset. Adjust as needed based on scope response time.
        scope.write("*CLS")
        time.sleep(5)
        
    if brand.lower() == "tek":
        print("Setting up Tektronix scope...")
        # Set up vertical
        for i in range(1, num_channels + 1):
            scope.write(f"SELect:CH{i} ON")
            scope.write(f"CH{i}:POSition 0")
            if i == 1:
                scope.write(f"CH{1}:SCALe 100")
            else:
                scope.write(f"CH{i}:SCALe 10")
            scope.write(f"MEASUrement:MEAS{i}:SOUrce CH{i}; STATE 1")   # Need separate STATE 1 command for DPO
            scope.write(f"MEASUrement:MEAS{i}:SOUrce CH{i}; TYPE RMS")  # Need separate TYPE command for MSO
            # scope.write(f"DISplay:SPECView1:VIEWStyle OVERly")   # Not working on MSO

        # Set up timebase
        scope.write("HORizontal:SCAle 200E-6")
        scope.write("HORizontal:POSition 50")

        # Set up trigger to line
        scope.write("TRIGger:A:EDGE:SOUrce CH1")
        scope.write("TRIGger:A:EDGE:COUPling DC")
        scope.write("TRIGger:A:EDGE:SLOpe RISE")
        scope.write("TRIGger:A:LEVel:CH1 50")

    elif  brand.lower() == "rigol" or brand.lower() == "other":
        # Hide all channels and measurements
        for i in range(1, max_channels + 1):
            scope.write(f":CHANnel{i}:DISPlay OFF") 

         # Vertical configure
        for i in range(1, num_channels + 1):
            if i == 1:
                scope.write(f":CHANnel{1}:SCALe 100")
            else:
                scope.write(f":CHANnel{i}:SCALe 10")
            scope.write(f":CHANnel{i}:DISPlay ON")
            scope.write(f":CHANnel{i}:PROBe 10")
            scope.write(f":CHANnel{i}:OFFSet 0")
            scope.write(f":CHANnel{i}:BWLimit 20M")
            scope.write(f":CHANnel{i}:COUPling DC")
            scope.write(f":CHANnel{i}:INVert OFF")
            scope.write(f":CHANnel{i}:UNITs VOLT")

        # Horizontal configure
        scope.write(":TIMebase:SCALe 200e-6")
        scope.write(":TIMebase:DELay 50")
        
        # Trigger configure to line
        scope.write(":TRIGger:MODE EDGE")
        scope.write(":TRIGger:EDGe:SOUrce CHAN1")
        scope.write(":TRIGger:EDGe:COUPling DC")
        scope.write(":TRIGger:EDGe:SLOpe POSitive")
        scope.write(":TRIGger:EDGe:LEVel 50")

    elif brand.lower() == "lecroy":
        # LeCroy scopes can be quite different. Using Visual Basic Scripting commands instead of SCPI.
        print("Be sure to set scope, utilities to TCPIP (VXI-11).  Setting up LeCroy scope...")
        
        # Hide all channels and measurements
        scope.write("VBS 'app.Measure.ClearAll'")
        scope.write("VBS 'app.Measure.ShowMeasure = True'")
        for i in range(1, max_channels + 1):
            scope.write(f"VBS 'app.Acquisition.C{i}.View = False'") 

        # Vertical- Configure active channels
        for i in range(1, num_channels + 1):
            # label and scale for first channel (AC line) differently
            label = "AC Power Line" if i == 1 else f"Amp Out {i-1}"
            scale = 100 if i == 1 else 10  # scales in V/div
            # list of VBS commands
            vertical_settings = [
                f"app.Acquisition.C{i}.VerOffset = 0",
                f"app.Acquisition.C{i}.Coupling = \"DC1M\"",
                f"app.Acquisition.C{i}.View = True",
                f"app.Acquisition.C{i}.ViewLabels = True",
                f"app.Acquisition.C{i}.BandwidthLimit = \"20MHz\"",
                f"app.Acquisition.C{i}.VerScale = {scale}",
                f"app.Acquisition.C{i}.LabelsText = \"{label}\"",
                # Setup Measurements (P1, P2, etc.)
                f"app.Measure.P{i}.ParamEngine = \"RootMeanSquare\"",
                f"app.Measure.P{i}.Operator.Cyclic = \"True\"",
                f"app.Measure.P{i}.Source1 = \"C{i}\"",
                f"app.Measure.P{i}.View = True"
            ]
            # Send all vertical commands
            for cmd in vertical_settings:
                scope.write(f"VBS '{cmd}'")

        # Horizontal - Configure initial timebase and trigger
        horizontal_settings = [
            "app.Acquisition.Horizontal.HorScale = 5e-3",
            "app.Acquisition.Horizontal.HorOffset = 0",
            "app.Acquisition.Trigger.Type = \"Edge\"",
            "app.Acquisition.Trigger.Edge.Source = \"C1\"",
            f"app.Acquisition.Trigger.Edge.Level = 60",
            "app.Acquisition.TriggerMode = \"Auto\""
        ]
        # Send all horizontal/trigger commands
        for cmd in horizontal_settings:
            scope.write(f"VBS '{cmd}'")

    elif brand.lower() == "keysight":
        print("Keysight oscilloscope detected.")
        # Need to add scope setup for Keysight 
        #
        #
    
    else:
        print("Unsupported oscilloscope model.")
        # enter code for Keysight or other scope
        # # stop_program_event.set()

    # # Wait for scope to finish setting up
    time.sleep(1)

def set_scope_trigger(scope, brand):
    # Ensure scope is running normally
    try:
        if brand.lower() == "lecroy":
            scope.write("VBS 'app.Acquisition.TriggerMode = \"Auto\"'")
        elif brand.lower() == "tek":
            scope.write("ACQuire:STATE ON")
        elif brand.lower() == "rigol" or brand.lower() == "other" or brand.lower() == "keysight":
            scope.write(":RUN") 
        return True
    except pyvisa.errors.VisaIOError as e:
        print(f"Communications Error (Trigger): {e}")
        return None

def set_scope_trigger_level(scope, brand, channel, level):
    """
    Sets the scope trigger level based on the brand, channel, and level.
    """
    brand_low = brand.lower()
    try:
        if brand_low == "lecroy":
            scope.write(f"VBS 'app.Acquisition.Trigger.Edge.Source = \"C{channel}\"'")
            scope.write(f"VBS 'app.Acquisition.Trigger.Edge.Level = {level}'")
        elif brand_low == "tek":
            scope.write(f"TRIGger:A:LEVel:CH{channel} {level}")
        elif brand_low == "rigol" or brand_low == "other":
            scope.write(f":TRIGger:EDGe:SOUrce CHAN{channel}")
            scope.write(f":TRIGger:EDGe:LEVel {level}")
        elif brand_low == "keysight":
            scope.write(f":TRIGger:LEVel CHAN{channel}, {level}")
        scope.query("*OPC?")
    except pyvisa.errors.VisaIOError as e:
        print(f"Error setting trigger level: {e}")

def make_datafile(timestamp, desktop_path: str = desktop_path): # num_channels removed as it's not needed for header anymore
    """Generates a data file for the data based on start date and time.
    Includes headers for each monitored channel. Asks user for directory
    or defaults to desktop.
    Returns tuple for user_path and data_log_file_name
    """
    while True:
        try:
            user_path_input = input(f"Enter path for data or 'd' for default ({desktop_path}): ").strip()
            if user_path_input.lower() == 'd':
                user_path = desktop_path
            else:
                user_path = user_path_input

            data_log_file_name = timestamp.strftime("%Y%m%d_%H%M%S.csv")
            full_data_path = os.path.join(user_path, data_log_file_name)

            # Check if data file already exists
            if not os.path.exists(full_data_path):
                print(f"Creating new data file at {user_path}.. .")
                with open(full_data_path, "w") as datafile:
                    # UPDATED: Header for duration logging
                    header = "Event_Count, Start_Time_Absolute, End_Time_Absolute, Line Voltage, State, Duration_Seconds"
                    datafile.write(header + "\n")
            else:
                print(f"File exist. We will be appending to existing file: {full_data_path}")
            return user_path, data_log_file_name

        except ValueError:
            print("Invalid path. Please enter a valid path.")

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

def get_scope_measurements(scope, brand, num_channels, current_state, force_load_query=False):
    """
    Queries all active channels, applies bounds, and calculates the 
    running average AC line.
    
    Returns:
        tuple: (timestamp, v_rms_readings, line_avg)
               Returns (None, None, None, None) if a reading fails.

        v_rms_readings: array of Vrms readings for all channels (CH1 AC Line and CH2+ Amplifier outputs)
    """
    global LAST_LOAD_CHECK_TIME
    # timestamp and setup
    current_time = datetime.datetime.now()
    v_rms_readings = []
    line_avg = 0.0
    reading = "0.0" # Default value
    brand_low = brand.lower()    
    # 1. Query CH1 ALWAYS (AC Line is heartbeat of the system)
    try:
        if brand_low == "lecroy":
            query_str = "VBS? 'return=app.Measure.P1.Out.Result.Value'"
            raw_reading = scope.query(query_str)
            vrms_reading = parse_visa_numeric(raw_reading)
        elif brand_low == "tek":
            vrms_reading = parse_visa_numeric(scope.query("MEASUrement:MEAS1:VALue?"))
        elif brand_low == "rigol" or brand_low == "other" or brand_low == "keysight":
            vrms_reading = parse_visa_numeric(scope.query(":MEASure:VRMS? CHAN1"))

        v_line_bounded = apply_line_voltage_bounds(vrms_reading)
        v_rms_readings.append(v_line_bounded)  # for all_readings[0] in main loop)
        
        # Update running average for steady-state logging
        line_voltage_readings_queue.append(v_line_bounded)
        line_avg = sum(line_voltage_readings_queue) / len(line_voltage_readings_queue)

    except (ValueError, pyvisa.errors.VisaIOError):
            return None, None, None


    # 2. PERIODICALLY Query Load Channels (CH2+)
    # Only query if UNKNOWN (initialization), if forced, or if ON and interval has passed.
    time_since_load_check = (time.time() - LAST_LOAD_CHECK_TIME)
    should_query_loads = (current_state == "UNKNOWN" or 
                          force_load_query or 
                          (current_state == "ON" and time_since_load_check > LOAD_CHECK_INTERVAL))
    
    if should_query_loads:
         # Update the last check time and get readings for all load channels (CH2+)
        LAST_LOAD_CHECK_TIME = time.time()
        for i in range(2, num_channels + 1):
            try:
                if brand_low == "lecroy":
                    query_str = f"VBS? 'return=app.Measure.P{i}.Out.Result.Value'"
                    reading = parse_visa_numeric(scope.query(query_str))
                elif brand_low == "tek":
                    reading = parse_visa_numeric(scope.query(f"MEASUrement:MEAS{i}:VALue?"))
                elif brand_low == "rigol" or brand_low == "keysight" or brand_low == "other":
                    reading = parse_visa_numeric(scope.query(f":MEASure:VRMS? CHAN{i}"))

                # Convert to float, apply bound
                v_rms_readings.append(apply_amp_output_bounds(float(reading)))
            except:
                v_rms_readings.append(0.0)
    else:
        # Fill list with 'None' to indicate no NEW data was fetched for loads
        v_rms_readings.extend([None] * (num_channels - 1))
    return current_time, v_rms_readings, line_avg

def log_duration_to_file(save_directory, data_file_name, event_count, start_time, end_time, line_voltage_avg, state, duration_seconds):
    """
    Appends duration data to the specified file.
    """
    try:
        datafile_and_path = os.path.join(save_directory, data_file_name)
        with open(datafile_and_path, "a") as f:
            line = f"{event_count:4d}, {start_time.strftime('%Y-%m-%d %H:%M:%S.%f')}, {end_time.strftime('%Y-%m-%d %H:%M:%S.%f')}, {line_voltage_avg:7.3f}, {state}, {duration_seconds:9.3f}"
            f.write(line + "\n")
    except IOError as e:
        print(f"Error appending data to file '{datafile_and_path}': {e}")

def write_to_excel(datafile_name: str, save_directory: str): # num_channels removed
    """
    Reads data from the specified CSV file, writes it to an Excel worksheet

    Args:
        datafile_name: The name of the CSV data file.
        save_directory: The directory where the CSV and Excel files are saved.
    """
    full_csv_path = os.path.join(save_directory, datafile_name)
    excel_file_name = os.path.splitext(datafile_name)[0] + ".xlsx"
    full_excel_path = os.path.join(save_directory, excel_file_name)

    print(f"\nAttempting to create Excel file: {full_excel_path}")

    wb = Workbook()
    ws = wb.active
    ws.title = "Power Monitoring Durations"

    try:
        with open(full_csv_path, 'r') as f:
            reader = csv.reader(f)
            # Write header row
            header = next(reader)
            ws.append(header)

            # Write data rows and convert to numbers, then format cells
            row_count = 1 # To track actual row in Excel, starting from 1 for headers
            for row in reader:
                row_count += 1 # Increment for data rows
                processed_row = []
                # Event_Count, Start_Time_Absolute, End_Time_Absolute, State, Duration_Seconds
                for i, value in enumerate(row):
                    if i == 0:  # Event_Count
                        try:
                            num_value = int(value)
                            processed_row.append(num_value)
                        except ValueError:
                            processed_row.append(value)
                    elif i in [1, 2]: # Start_Time_Absolute, End_Time_Absolute - keep as strings for now or convert to datetime objects if needed for excel, but direct plotting might be tricky
                        processed_row.append(value)
                    elif i == 3: # Line_Voltage
                        try:
                            num_value = float(value)
                            processed_row.append(num_value)
                            ws.cell(row=row_count, column=i+1).number_format = '0.000'
                        except ValueError:
                            processed_row.append(None)
                    elif i == 4: # State
                        processed_row.append(value)
                    elif i == 5: # Duration_Seconds
                        try:
                            num_value = float(value)
                            processed_row.append(num_value)
                            ws.cell(row=row_count, column=i+1).number_format = '0.000'
                        except ValueError:
                            processed_row.append(None) # Append None for invalid duration values
                ws.append(processed_row)
        wb.save(full_excel_path)
        print(f"Excel data saved successfully to: {full_excel_path}")

    except FileNotFoundError:
        print(f"Error: CSV data file not found at {full_csv_path}. Cannot create Excel file.")
    except Exception as e:
        print(f"An error occurred while creating the Excel file: {e}")

# ************** MAIN
rm = None
scope = None
last_state_change_time = datetime.datetime.now()
event_counter = 0
first_transition_logged = False

# Initialize deque for running average on line voltage readings
line_voltage_readings_queue = deque(maxlen=LINE_VOLTAGE_WINDOW_SIZE)
steady_state_line_voltage = 0.0

try:
    num_channels_to_monitor = 0
    scope = None
    datafile_name = None  # Initialize datafile_name to None
    current_state = "UNKNOWN"  # States can be "ON" or "OFF"
   
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

    # Get user settings
    # Get maximum number of channels on this scope model
    max_ch_on_scope = get_max_channels()      # Some scopes have 2, 4, or 8 CH.

    # Get the number of channels from user being hooked up and monitoring. 
    num_channels_to_monitor = get_num_channels(max_ch_on_scope)

    # Get AC Line threshold levels
    print("Enter AC Line Vrms ON/OFF thresholds:")
    ac_line_high_limit, ac_line_low_limit = get_thresholds(
        default_on_value=80, 
        default_off_value=70, 
        max_limit=275, 
        min_limit=1
    )
    print("AC Line Vrms ON = ", ac_line_high_limit, ", AC Line Vrms OFF = ", ac_line_low_limit)

    # Get Amp Output threshold levels
    print("Enter Amp Output Vrms ON/OFF thresholds:")
    amp_high_limit, amp_low_limit = get_thresholds(
        default_on_value=7.0, 
        default_off_value=2.0, 
        max_limit=20.0,
        min_limit=1
    )
    print("Amp Output Vrms ON = ", amp_high_limit, ", Amp Output Vrms OFF = ", amp_low_limit)

    # Get power-on delay to check output signals
    power_on_delay = get_power_on_delay()
    print(f"Power-on delay set to {power_on_delay} seconds.")

    # Set up scope?
    setup_needed = input("(L)eave scope alone or (S)etup contiguous channels?: ").strip()
    if setup_needed.lower() == 'l':
        print("Skipping scope setup. Ensure channels are configured correctly, and auto-triggering is enabled before starting data acquisition.")
    else:
        print("Attempting to set up scope.. .")
        setup_scope(scope, num_channels_to_monitor, max_ch_on_scope, brand)
        print("Setup complete.")

    # Trigger mode normal?
    run_mode = input("Is trigger mode set up? (Y/N): ").strip().lower()
    if run_mode == 'n':
        print("Setting trigger mode.")
        set_scope_trigger(scope, brand)
    else: 
        print("Setup complete. Check scope settings acceptable before starting data acquisition.")

    # Create a data file for logging based on the current timestamp. This time/data log be used for duration of test.
    last_state_change_time = datetime.datetime.now()
    user_path, datafile_name = make_datafile(last_state_change_time, desktop_path)
    full_data_path = os.path.join(user_path, datafile_name)
    print("Created file for data as", datafile_name)

    # Notify user ready to start. Provide instructions for stopping the program and ensuring proper initial conditions.
    print(f"Monitoring for ON (all channels > {ac_line_high_limit:.2f} Vrms) and OFF (all channels < {ac_line_low_limit:.2f} Vrms) states.")
    input("Hit Enter to start monitoring...")
    print("Press 'q','esc', or 'Crtl-C' to stop the program at any time.")
    
    # ************** MAIN LOOP  *****************
    while not stop_program_event.is_set():
        time.sleep(0.05) # Small delay for keyboard input before checking if scope triggered

        # Get measurements from scope.  Also assure auto run mode, averaging, apply bounds.
        meas_time, all_readings, ac_line_avg = get_scope_measurements(
            scope, brand, num_channels_to_monitor, current_state
        )
        # Check if measurement failed and skip this cycle.
        if meas_time is None: continue

        ac_line_voltage = all_readings[0]
        new_amp_data = all_readings[1:] 

        # Assign state Booleans from ac line (instantaneous)
        ac_line_on = ac_line_voltage >= ac_line_high_limit
        ac_line_off = ac_line_voltage <= ac_line_low_limit

        # Update steady state voltage logic
        if current_state == "ON":
            time_since_on = (meas_time - last_state_change_time).total_seconds()
            
            # 1. During the ramp (under 3s), keep capturing the highest value seen
            # This ensures short cycles (<3s) still have the best possible data.
            if time_since_on < SETTLING_TIME:
                if ac_line_voltage > steady_state_line_voltage:
                    steady_state_line_voltage = ac_line_voltage
            
            # 2. Once settled (over 3s), update to the current stable reading.
            # We use ac_line_avg here for a cleaner, noise-filtered steady state.
            elif ac_line_voltage >= ac_line_high_limit:
                steady_state_line_voltage = ac_line_avg

        if None not in new_amp_data:
            amp_load_voltage = new_amp_data  # new data for load channels (CH2+)
            all_load_channels_on = all(v >= amp_high_limit for v in amp_load_voltage)
            all_load_channels_off = all(v <= amp_low_limit for v in amp_load_voltage)

            # Load Drop-out Notification
            if current_state == "ON" and ac_line_on:
                # Check if  ON state for more than the power-on delay 10 seconds
                time_in_on_state = (meas_time - last_state_change_time).total_seconds()
                if time_in_on_state > power_on_delay and not all_load_channels_on:
                    # Calculate average of all amp load channels
                    avg_amp_vrms = sum(amp_load_voltage) / len(amp_load_voltage)
                    print(
                        f"[{meas_time.strftime('%H:%M:%S')}] "
                        f"Signal drop-out with a line voltage of {ac_line_avg:5.2f} Vrms "
                        f"and on-time of {time_in_on_state:5.1f}s. "
                        f"Avg Load: {avg_amp_vrms:.3f} Vrms."
                    )

        # State Establishment/Transition Logic
        if current_state == "UNKNOWN":
            if ac_line_on:
                current_state = "ON"
                print(f"Initial state detected as ON. Will start logging on next transition to OFF.")
                last_state_change_time = meas_time
                set_scope_trigger_level(scope, brand, 1, ac_line_low_limit)
            elif ac_line_off:
                current_state = "OFF"
                print(f"Initial state detected as OFF. Will start logging on next transition to ON.")
                last_state_change_time = meas_time
                set_scope_trigger_level(scope, brand, 1, ac_line_high_limit)
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
                print(f"State captured. Waiting for first transition to ON.")
                current_state = new_actual_state
                last_state_change_time = meas_time
                first_transition_logged = True
                # Set the trigger level for the next transition depending on state
                if current_state == "ON":       # State is ON, set low limit
                    set_scope_trigger_level(scope, brand, 1, ac_line_low_limit)
                else:                           # State is OFF, set high limit
                    set_scope_trigger_level(scope, brand, 1, ac_line_high_limit)
            else: 
                # This is a subsequent transition; start logging from this pointon
                duration = (meas_time - last_state_change_time).total_seconds()
                event_counter += 1
                if current_state == "OFF" and new_actual_state == "ON":
                    # Transition from OFF to ON  (log the previous OFF state, reset timers, line voltage, off trigger)
                    log_duration_to_file(user_path, datafile_name, event_counter, last_state_change_time, meas_time, 0.0, "OFF", duration)
                    print(f"State Change: OFF to ON. Previous OFF duration: {duration:.3f} seconds.")
                    current_state = "ON"
                    last_state_change_time = meas_time
                    steady_state_line_voltage = 0.0
                    set_scope_trigger_level(scope, brand, 1, ac_line_low_limit)

                elif current_state == "ON" and new_actual_state == "OFF":
                    # Transition from ON to OFF
                    # Determine the line voltage to log
                    voltage_to_log = steady_state_line_voltage

                    # # HARD-CODE: Set line voltage to 0.0 if this is the very first ON-to-OFF transition (event_counter == 1)
                    # # This explicitly handles the initial corrupted reading you observed.
                    # if event_counter == 1:
                    #     voltage_to_log = 0.0

                    log_duration_to_file(user_path, datafile_name, event_counter, last_state_change_time, meas_time, voltage_to_log, "ON", duration)
                    print(f"State Change: ON to OFF. Previous ON duration: {duration:.3f} seconds. Line Voltage: {voltage_to_log:.3f}Vrms")
                    current_state = "OFF"
                    last_state_change_time = meas_time
                    # Set trigger for next OFF detection
                    set_scope_trigger_level(scope, brand, 1, ac_line_high_limit)

            # *********  Only needed for debug  ***********
            # print(f"CH2+ Readings for State Check: {[f'{v:.3f}' for v in v_rms_readings_for_state_check]}")
            # print(f"All channels ON condition: {all_channels_on}")
            # print(f"All channels OFF condition: {all_channels_off}")
            # print(f"Current State: {current_state}")

        # At this point, both all_channels_on and all_channels_off are false
        # No change in state, just loop around and continue to monitor
        # new_actual_state = current_state

except KeyboardInterrupt:
    print("\nProgram terminated by user (Ctrl+C).")
except Exception as e:
    print(f"An error occurred during program execution: {e}")
finally:
    # Close the instrument connection and resource manager
    if scope:
        # Stop acquisition before closing
        try:
            scope_brand = brand.lower() if 'brand' in locals() else "unknown"
            if brand.lower() in ["lecroy", "rigol", "keysight", "other"]:
                scope.write(":STOP")
            elif brand.lower() == "tek":
                scope.write("ACQuire:STATE OFF") 
            else:
                print("Unsupported oscilloscope model.")
                stop_program_event.set()
            scope.write("CLEAR") # Ensure scope acquisition is stopped
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
    duration = (final_time - last_state_change_time).total_seconds()
    final_voltage_to_log = 0.000 # Default to 0.0

    # Only attempt to use current_line_voltage_snapshot if the queue has data
    if len(line_voltage_readings_queue) > 0:
        final_voltage_to_log = sum(line_voltage_readings_queue) / len(line_voltage_readings_queue)

    if current_state == "OFF":
        # Always log 0.0 for line voltage if the final state is OFF
        if user_path and datafile_name:
            log_duration_to_file(user_path, datafile_name, event_counter, last_state_change_time, final_time, final_voltage_to_log, current_state, duration)
        print(f"Program stopped. Final {current_state} duration: {duration:.3f} seconds. Line Voltage hardcoded for OFF state.)")
    else: # If the final state was ON
        # Apply the same hard-code logic for the final ON state if event_counter is 0 or 1
        # This handles cases where the program exits very quickly after starting
        if event_counter == 0 or (event_counter == 1 and duration < 1.0): # event_counter 0 means no transitions logged. 1 means the initial state captured.
             final_voltage_to_log = 0.0 # Hardcode as too  early in program for accurate readings
             print("Note: Line voltage set 0.0V due to early program termination or initial state.")
        if user_path and datafile_name:
            log_duration_to_file(user_path, datafile_name, event_counter, last_state_change_time, final_time, final_voltage_to_log, current_state, duration)
        print(f"Program stopped. Final {current_state} duration: {duration:.3f} seconds. Line Voltage: {final_voltage_to_log:.3f}Vrms")

    # Create Excel file from log file.
    if user_path and datafile_name:
        print("Creating mirror Excel data file...")
        write_to_excel(datafile_name, user_path)

    input("\nExecution complete. Press Enter to exit...")
    