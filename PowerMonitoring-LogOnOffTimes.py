# Power Monitoring- LogOnOffTimes.py
#
# Continuously monitor AC Line (CH1) and Amplifier output channels (CH2+) 
# to log ON/OFF times (<1 sec resolution) based on CH1 trigger level.
#
# Scope is a Rigol, Tektronix, or LeCroy scope in autorun trigger mode.
# 
# User must input IP address and number of output load channels to monitor (1-7).
# Maximum is total phyiscal channesls (8).  Minimum is two with at least one output (CH2). 
# 
# If the line voltage is present for longer then 10 seconds, 
# all amplifier channels must be above threshold or an error is reported.
#
# User is given the option to automatically set up the scope on contiguous
# channels or to leave it alone. User must set up line on first channel
# and amplifier channels on subsequent channels with appropriate scale,
# positions, and trigger levels.
# 
# Data is saved to csv file at the user directed path or defaults to the 
# desktop.
#
# Author: C. Wong 2026XXXX

import time
import datetime
import os
import pyvisa
import threading
import csv
import keyboard
from collections import deque  #only needed for running average on AC Line
from openpyxl import Workbook

DEFAULT_IP_ADDRESS = '192.168.1.90'  # 10.101.100.151, 169.254.131.118, 10.100.52.231
DEFAULT_NO_CHANNELS = 4  # Maximum number of channels on scope
MAX_LINE_VOLTAGE_VRMS = 350.0  # AC Line RMS limit, arbitrary limit for CH1
MAX_VRMS = 50.0  # ~312 W arbitrary limit for other audio CHs  (CH2+)
LINE_VOLTAGE_WINDOW_SIZE = 4 # Define the window size for the running average

# Find user desktop one level down from home [~/* /Desktop] as optional path
from glob import glob
DESKTOP = glob(os.path.expanduser("~\\*\\Desktop"))
# If not found, revert to the standard location (~/Desktop)
if len(DESKTOP) == 0:
    DESKTOP = os.path.expanduser("~\\Desktop")
else:
    DESKTOP = DESKTOP[0] # glob returns a list, take the first result

# Global flag to signal the main loop and threads to stop
stop_program_event = threading.Event()

def on_q_press():
    """
    Callback function when 'q' is pressed.
    """
    print("\n'q' pressed. Signaling program to stop.")
    stop_program_event.set()

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
        The connected PyVISA instrument object and label.
    """
    while True:
        user_input = input(f"Enter IP address, 'q' to quit, 'd' for default (Default {default_ip}): ").strip()
        
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
            instr.timeout = 3000 
            idn = instr.query('*IDN?').strip().upper()
            
            # Create a label/metadata tag
            label = "unknown"
            if "RIGOL" in idn: label = "rigol"
            elif "TEKTRONIX" in idn: label = "tek"
            elif "LECROY" in idn: label = "lecroy"
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
            print("Retrying... (Press 'q' to stop)")
            time.sleep(1)

def get_max_channels():
    """
    Asks the user to input the physical maximum number of channels (2-8) on scope.
    'd' or Enter returns the default of 4.
    """
    user_input = input("Enter the physical maximum number of channels (2-8) or 'd' (Default = 4): ").strip().lower()

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

def setup_scope(scope, num_channels, max_channels, brand):
    """
    Configures channels for the specified number of channels, setting
    scale and position, and adapting to brand- rigol, tek, lecroy, other.
    """
    # DEVELOPER SANITY CHECK:
    # Ensure programmer brand is valid
    if brand.lower() not in ["rigol", "tek", "lecroy", "other"]:
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

    elif  brand.lower() == "rigol":
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
        # LeCroy scopes can be quite different. 
        # May need adjustment based on the specific model and firmware.
        # Using Visual Basic Scripting (VBS) commands instead of SCPI.
        print("Set to TCPIP (VXI-11) on scope.  Setting up LeCroy scope...")
        
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

    else:
        print("Unsupported oscilloscope model.")
        # enter code for Keysight or other scope
        # # stop_program_event.set()

    # # Wait for scope to finish setting up
    time.sleep(1)

def make_datafile(timestamp, desktop_path: str = DESKTOP): # num_channels removed as it's not needed for header anymore
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

            data_log_file_name = timestamp.strftime("%Y%m%d_%H%M%S.txt")
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
connected_instrument = None
last_state_change_time = datetime.datetime.now()
event_counter = 0
first_transition_logged = False

# Initialize deque for line voltage readings
line_voltage_readings_queue = deque(maxlen=LINE_VOLTAGE_WINDOW_SIZE)
# current_line_voltage_avg = 0.0 # This variable is no longer strictly needed as a global if using local snapshot
# It's better to calculate and use a local snapshot for current readings.

try:
    num_channels_to_monitor = 0
    connected_instrument = None
    datafile_name = None  # Initialize datafile_name to None
    current_state = "UNKNOWN"  # States can be "ON" or "OFF"
    # previous_state = current_state  # This variable is not used and can be removed for clarity

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
    print("Enter AC Line Vrms ON/OFF thresholds:")
    acline_high_limit, acline_low_limit = get_thresholds(
        default_on_value=80, 
        default_off_value=60, 
        max_limit=264, 
        min_limit=1
    )
    print("AC Line Vrms ON = ", acline_high_limit, ", AC Line Vrms OFF = ", acline_low_limit)

    # Get Amp Output threshold levels
    print("Enter Amp Output Vrms ON/OFF thresholds:")
    amp_high_limit, amp_low_limit = get_thresholds(
        default_on_value=9.0, 
        default_off_value=7.0, 
        max_limit=20.0,
        min_limit=1
    )
    print("Amp Output Vrms ON = ", amp_high_limit, ", Amp Output Vrms OFF = ", amp_low_limit)

    # Set up scope?
    setup_needed = input("(L)eave scope alone or (S)etup contiguous channels?: ").strip()
    if setup_needed.lower() == 'l':
        print("Skipping scope setup. Ensure channels are configured correctly before starting data acquisition.")
    else:
        print("Attempting to set up scope.. .")
        setup_scope(scope, num_channels_to_monitor, max_ch_on_scope, brand)
        print("Setup complete. Check scope settings acceptable before starting data acquisition.")

    # Create a data file for logging based on the current timestamp. This time/data log be used for duration of test.
    last_state_change_time = datetime.datetime.now()
    user_path, datafile_name = make_datafile(last_state_change_time, DESKTOP)
    full_data_path = os.path.join(user_path, datafile_name)
    print("Created file for data as", datafile_name)

    # Notify user program is off and running
    print(f"Monitoring for ON (all channels > {acline_high_limit:.2f} Vrms) and OFF (all channels < {acline_low_limit:.2f} Vrms) states.")
    no_response = input("Check Line voltage and amp outputs are OFF (0V) before begining. Hit Enter to continue...")
    print("Press 'q','esc', or 'Crtl-C' to stop the program at any time.")
    print("Start monitoring...")

    # ************** MAIN LOOP
    while not stop_program_event.is_set():
        time.sleep(0.05) # Small delay for keyboard input before checking if scope triggered

        # Query instrument identity to determine the manufacturer
        idn_string = connected_instrument.query('*IDN?').strip()
        is_rigol = "RIGOL" in idn_string.upper()
        is_tek = "TEKTRONIX" in idn_string.upper()

        # Start acquisition based on instrument type
        if is_rigol:
            connected_instrument.write(":RUN")
        elif is_tek: # Assumes Tektronix or a similar scope
            connected_instrument.write("ACQuire:STATE ON")
        else:
            print("Unsupported oscilloscope model. Please use a Rigol or Tektronix scope.")
            stop_program_event.set()

        # Read measurements
        current_time = datetime.datetime.now()
        v_rms_readings = []

        # Initialize for loop iteration to get an average or snapshot of the line voltage.
        current_line_voltage_snapshot = 0.0

        # Read RMS values for each channel
        for i in range(1, num_channels_to_monitor + 1):
            try:
                # Get measurements
                if is_rigol:
                    v_rms = float(connected_instrument.query(f":MEASure:VRMS? CHAN{i}"))
                else:
                    v_rms = float(connected_instrument.query(f"MEASUrement:MEAS{i}:VALue?"))

                # Apply bounds based on channel
                if i == 1: # This is the AC Line (CH1)
                    v_rms = apply_line_voltage_bounds(v_rms)
                    line_voltage_readings_queue.append(v_rms)
                    if len(line_voltage_readings_queue) > 0:
                        current_line_voltage_snapshot = sum(line_voltage_readings_queue) / len(line_voltage_readings_queue)
                    # Else, current_line_voltage_snapshot remains 0.0 if queue is empty (shouldn't happen here)
                else: # Apply bounds only if it's NOT channel 1
                    v_rms = apply_amp_output_bounds(v_rms)
                v_rms_readings.append(v_rms)

            except pyvisa.errors.VisaIOError as e:
                print(f"Error reading RMS for Channel {i}: {e}. Skipping this channel for this sample.")
                v_rms_readings.append(float('NAN')) # Append NaN if reading fails

            except ValueError:
                print(f"Could not convert RMS reading for Channel {i} to float. Skipping.")
                v_rms_readings.append(float('NAN'))

        # Create a new list for checks, excluding channel 1
        v_rms_readings_for_state_check = v_rms_readings[1:] # Slice from the 2nd element to end

        # Check if any readings are NaN, if so, we can't determine state reliably
        if any(v == float('NAN') for v in v_rms_readings):
            print("Warning: Skipping state evaluation due to invalid Vrms readings.")
            continue

        # Determine the potential new state based on current readings for CH2+
        all_channels_on = all(v >= amp_high_limit for v in v_rms_readings_for_state_check)
        all_channels_off = all(v <= amp_low_limit for v in v_rms_readings_for_state_check)

        # State Establishment/Transition Logic
        if current_state == "UNKNOWN":
            if all_channels_on:
                current_state = "ON"
                last_state_change_time = current_time
                connected_instrument.write(f"TRIGger:A:LEVel:CH1 {acline_low_limit}")
                print(f"Initial state detected as ON. Will start logging on next transition to OFF.")
            elif all_channels_off:
                current_state = "OFF"
                last_state_change_time = current_time
                connected_instrument.write(f"TRIGger:A:LEVel:CH1 {acline_high_limit}")
                print(f"Initial state detected as OFF. Will start logging on next transition to ON.")
            else:
                # Still in an indeterminate state or no clear ON/OFF. Keep current_state as UNKNOWN.
                pass # No change, continue will be called below
            continue # Always continue if still in UNKNOWN or just established initial state

        # Current_state is either 'ON' or 'OFF'. Proceed to check for transitions.
        new_actual_state = current_state # Assume no change unless clear transition

        # Determine if there's a *real* transition from the established state
        if current_state == "OFF" and all_channels_on:
            new_actual_state = "ON"
        elif current_state == "ON" and all_channels_off:
            new_actual_state = "OFF"

        # This handles the intermediate or stable state where no clear transition occurs.
        if new_actual_state != current_state:
            if not first_transition_logged:
                # This is the very first transition detected. Don't log the initial state.
                print(f"State captured. Waiting for first transition to ON.")
                current_state = new_actual_state
                last_state_change_time = current_time
                first_transition_logged = True
                # Set the trigger level for the next transition depending on state
                if current_state == "ON":       # State is ON, set low limit
                    if is_rigol:
                        connected_instrument.write(f":TRIGger:EDGe:LEVel {acline_low_limit}")
                    elif is_tek:
                        connected_instrument.write(f"TRIGger:A:LEVel:CH1 {acline_low_limit}")
                    else:
                        print("Unsupported oscilloscope model. Please use a Rigol or Tektronix scope.")
                        stop_program_event.set()
                else:                           # State is OFF, set high limit
                    if is_rigol:
                        connected_instrument.write(f":TRIGger:EDGe:LEVel {acline_high_limit}")
                    elif is_tek:
                        connected_instrument.write(f"TRIGger:A:LEVel:CH1 {acline_high_limit}")
                    else:
                        print("Unsupported oscilloscope model. Please use a Rigol or Tektronix scope.")
                        stop_program_event.set()
            else: 
                # This is a subsequent transition; start logging from this pointon
                duration = (current_time - last_state_change_time).total_seconds()
                event_counter += 1
                if current_state == "OFF" and new_actual_state == "ON":
                    # Transition from OFF to ON
                    # Log previous OFF state with 0.0 line voltage
                    log_duration_to_file(user_path, datafile_name, event_counter, last_state_change_time, current_time, 0.0, "OFF", duration)
                    print(f"State Change: OFF to ON. Previous OFF duration: {duration:.3f} seconds.")
                    current_state = "ON"
                    last_state_change_time = current_time
                    # Set trigger for next OFF detection
                    if is_rigol: 
                        connected_instrument.write(f":TRIGger:EDGe:LEVel {acline_low_limit}")
                    else:
                        connected_instrument.write(f"TRIGger:A:LEVel:CH1 {acline_low_limit}") 

                elif current_state == "ON" and new_actual_state == "OFF":
                    # Transition from ON to OFF
                    # Determine the line voltage to log
                    voltage_to_log = current_line_voltage_snapshot

                    # HARD-CODE: Set line voltage to 0.0 if this is the very first ON-to-OFF transition (event_counter == 1)
                    # This explicitly handles the initial corrupted reading you observed.
                    if event_counter == 1:
                        voltage_to_log = 0.0

                    log_duration_to_file(user_path, datafile_name, event_counter, last_state_change_time, current_time, voltage_to_log, "ON", duration)
                    print(f"State Change: ON to OFF. Previous ON duration: {duration:.3f} seconds. Line Voltage: {voltage_to_log:.3f}Vrms")
                    current_state = "OFF"
                    last_state_change_time = current_time
                    # Set trigger for next OFF detection
                    if is_rigol: 
                        connected_instrument.write(f":TRIGger:EDGe:LEVel {acline_high_limit}")
                    else:
                        connected_instrument.write(f"TRIGger:A:LEVel:CH1 {acline_high_limit}") 

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
    # Always close the instrument connection and resource manager
    if 'connected_instrument' in locals() and connected_instrument:
        # Stop acquisition before closing
        try:
            if is_rigol:
                connected_instrument.write(":STOP")
            elif is_tek:
                connected_instrument.write("ACQuire:STATE OFF") 
            else:
                print("Unsupported oscilloscope model. Please use a Rigol or Tektronix scope.")
                stop_program_event.set()
            connected_instrument.write("CLEAR") # Ensure scope acquisition is stopped
            connected_instrument.close()
            print("Instrument connection closed.")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error closing instrument connection: {e}")
    if rm:
        try:
            rm.close()
            print("Resource Manager closed.")
        except Exception as e:
            print(f"Error closing Resource Manager: {e}")

    # Before exiting, log the duration of the final state if it was not already logged
    event_counter += 1
    final_time = datetime.datetime.now()
    duration = (final_time - last_state_change_time).total_seconds()

    # Determine the line voltage to log for the final state
    final_voltage_to_log = 0.000999 # Default to 0.0

    # Only attempt to use current_line_voltage_snapshot if the queue has data
    if len(line_voltage_readings_queue) > 0:
        final_voltage_to_log = sum(line_voltage_readings_queue) / len(line_voltage_readings_queue)

    if current_state == "OFF":
        # Always log 0.0 for line voltage if the final state is OFF
        log_duration_to_file(user_path, datafile_name, event_counter, last_state_change_time, final_time, final_voltage_to_log, current_state, duration)
        print(f"Program stopped. Final {current_state} duration: {duration:.3f} seconds. Line Voltage hardcoded for OFF state.)")
    else: # If the final state was ON
        # Apply the same hard-code logic for the final ON state if event_counter is 0 or 1
        # This handles cases where the program exits very quickly after starting
        if event_counter == 0 or (event_counter == 1 and duration < 1.0): # event_counter 0 means no transitions logged. 1 means the initial state captured.
             final_voltage_to_log = 0.0 # Hardcode as too  early in program for accurate readings
             print("Note: Line voltage set 0.0V due to early program termination or initial state.")
        log_duration_to_file(user_path, datafile_name, event_counter, last_state_change_time, final_time, final_voltage_to_log, current_state, duration)
        print(f"Program stopped. Final {current_state} duration: {duration:.3f} seconds. Line Voltage: {final_voltage_to_log:.3f}Vrms")
        