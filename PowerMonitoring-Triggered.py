# Power Monitoring- Triggered
#
# Continuous monitor Amplifier Output channels and/or Line Inputs with
# TRIGGERED time logging.
#
# User inputs IP address and number of channels to monitor (1-8), and
# the option to set up scope or allow contiguous channels to be configured
# for RMS measurements.  User sets desired ON or OFF threshold.
#
# Saves data to csv file to user path or defaults to the desktop.
#
# Author: C. Wong 20250712

import time
import datetime
import os
import pyvisa
import threading
import csv
import keyboard

from openpyxl import Workbook

DEFAULT_IP_ADDRESS = '192.168.1.53'  #default IP, 192.168.1.53, 10.101.100.151
MAX_VRMS = 50
ON_THRESHOLD = 3.0  #default trigger levels for 'ON'
OFF_THRESHOLD = 1.0 #default trigger levels for 'OFF'

# Find user desktop one level down from home (~/* /Desktop) and set up as optional save path
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
    to a PyVISA instrument, retrying until successful.

    Args:
        resource_manager: The PyVISA ResourceManager instance.
        default_ip: The default IP address to suggest to the user.

    Returns:
        The connected PyVISA instrument object.
    """
    my_instrument = None
    while my_instrument is None:
        ip_address_input = input(
            f"Enter the instrument's IP address or 'd' for default ({default_ip}): "
        ).strip()
        if ip_address_input.lower() == 'd':
            visa_address = default_ip
        else:
            visa_address = ip_address_input

        # Construct the full VISA resource string
        resource_string = f'TCPIP::{visa_address}::INSTR'

        print(f"Attempting to connect to: {resource_string}")

        try:
            # Attempt to open the resource
            my_instrument = resource_manager.open_resource(resource_string)

            # Try to query the instrument to verify connection
            print(f"Successfully connected! Instrument ID: {my_instrument.query('*IDN?').strip()}")

        except pyvisa.errors.VisaIOError as e:
            print(f"Connection failed: {e}")
            print("Please ensure the IP address is correct and the instrument is on and connected to the network.")
            print("Retrying in 2 seconds...")
            my_instrument = None   # Ensure my_instrument is None to continue the loop
            time.sleep(2)   # Wait before retrying

        except Exception as e:
            print(f"An unexpected error occurred: {e}")
            print("Retrying in 2 seconds...")
            my_instrument = None
            time.sleep(2)

    return my_instrument

def get_num_channels():
    """
    Asks the user to input the number of channels (1-8).
    """
    while True:
        try:
            num_channels = int(input("Enter the number of channels to monitor (1-8): "))
            if 1 <= num_channels <= 8:
                return num_channels
            else:
                print("Invalid input. Please enter a number between 1 and 8.")
        except ValueError:
            print("Invalid input. Please enter a number.")

def get_thresholds(on_trig_level: float = ON_THRESHOLD, off_trig_level: float = OFF_THRESHOLD):
    """
    Prompts the user for ON and OFF threshold levels for Vrms.
    """
    while (on_trig_level <= 0 or off_trig_level <= 0 or on_trig_level <= off_trig_level or on_trig_level >= MAX_VRMS or off_trig_level >= MAX_VRMS):
        on_trig_level_input = input(
            f"Enter trigger level for ON cycle ({on_trig_level}): "
        ).strip()
        if on_trig_level_input.lower() == 'd':
            on_trig_level = on_trig_level
        else:
            on_trig_level = on_trig_level_input

        off_trig_level_input = input(
            f"Enter trigger level for ON cycle ({off_trig_level}): "
        ).strip()
        if off_trig_level_input.lower() == 'd':
            off_trig_level = off_trig_level
        else:
            off_trig_level = off_trig_level_input
    return on_trig_level, off_trig_level

def setup_scope(scope_device, num_channels):
    """
    Configures channels for the specified number of channels.
    Minimally tries to set scale and position.
    """
    print("Ok. Setting up oscilloscope...", end='')

    scope_device.write("*RST")  # Only needed for stubborn scopes

    for i in range(1, num_channels + 1):
        scope_device.write(f"SELect:CH{i} ON")
        scope_device.write(f"CH{i}:SCALe 10")
        scope_device.write(f"CH{i}:POSition 0")
        scope_device.write(f"MEASUrement:MEAS{i}:SOUrce CH{i}; STATE 1")   # Need separate STATE 1 command for DPO
        scope_device.write(f"MEASUrement:MEAS{i}:SOUrce CH{i}; TYPE RMS")  # Need separate TYPE command for MSO
    # scope_device.write("TRIGger:A:EDGE:SOUrce CH1")
    # scope_device.write("TRIGger:A:EDGE:COUPling DC")
    # scope_device.write("TRIGger:A:EDGE:SLOpe RISE")
    # scope_device.write("TRIGger:A:LEVel:CH1 50")  #Set initially to high level to delay trigger
    # scope_device.write("TRIGger:A:MODe NORMal")
    # scope_device.write("TRIGger:A:TYPe EDGE")
    # scope_device.write("ACQuire:STOPAfter SEQUENCE")
    # scope_device.write("ACQuire:STATE ON")

    # Wait for scope to finish setting up
    scope_device.query("*OPC?")
    print("Scope setup complete.")

def make_datafile(timestamp, desktoppath: str = DESKTOP): # num_channels removed as it's not needed for header anymore
    """Generates a data file for the data based on start date and time.
    Includes headers for each monitored channel. Asks user for directory
    or defaults to desktop.
    Returns tuple for user_path and data_log_file_name
    """
    while True:
        try:
            user_path_input = input(f"Enter path for data or 'd' for default ({desktoppath}): ").strip()
            if user_path_input.lower() == 'd':
                user_path = desktoppath
            else:
                user_path = user_path_input

            data_log_file_name = timestamp.strftime("%Y%m%d_%H%M%S.txt")
            full_data_path = os.path.join(user_path, data_log_file_name)

            # Check if data file already exists
            if not os.path.exists(full_data_path):
                print(f"Creating new data file: {full_data_path}")
                with open(full_data_path, "w") as datafile:
                    # UPDATED: Header for duration logging
                    header = "Event_Count, Start_Time_Absolute, End_Time_Absolute, State, Duration_Seconds"
                    datafile.write(header + "\n")
            else:
                print(f"File exist. We will be appending to existing file: {full_data_path}")
            return user_path, data_log_file_name

        except ValueError:
            print("Invalid path. Please enter a valid path.")

def log_duration_to_file(save_directory, data_file_name, event_count, start_time, end_time, state, duration_seconds):
    """
    Appends duration data to the specified file.
    """
    try:
        datafile_and_path = os.path.join(save_directory, data_file_name)
        with open(datafile_and_path, "a") as f:
            line = f"{event_count:4d}, {start_time.strftime('%Y-%m-%d %H:%M:%S.%f')}, {end_time.strftime('%Y-%m-%d %H:%M:%S.%f')}, {state}, {duration_seconds:9.3f}"
            f.write(line + "\n")
    except IOError as e:
        print(f"Error appending data to file '{datafile_and_path}': {e}")

def apply_vrms_bounds(number: float) -> float:
    """
    Applies upper and lower bounds to the Vrms reading.
    """
    return max(min(number, MAX_VRMS), 0)

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
                    elif i == 3: # State
                        processed_row.append(value)
                    elif i == 4: # Duration_Seconds
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

try:
    num_channels_to_monitor = 0
    connected_instrument = None
    datafile_name = None  # Initialize datafile_name to None
    current_state = "UNKNOWN"  # Can be "ON" or "OFF"
    last_state_change_time = datetime.datetime.now()
    event_counter = 0

    # Register the 'q' hotkey
    keyboard.add_hotkey('q', on_q_press)
    keyboard.add_hotkey('esc', on_esc_press)

    # Initialize the Resource Manager
    rm = pyvisa.ResourceManager()
    connected_instrument = connect_to_instrument(rm, DEFAULT_IP_ADDRESS)
    if connected_instrument is None:
        print("Failed to connect to the instrument. Exiting.")
        exit() # Exit if connection failed

    # Get the number of channels from the user
    num_channels_to_monitor = get_num_channels()

    # Get the number of channels from the user
    Limits = get_thresholds()
    print("Limit[0] = ", Limits[0], ", Limit[1] = ", Limits[1])

    # Set up channels based on the user input
    setup_needed = input("(L)eave scope alone or (S)etup contiguous channels?: ").strip()
    if setup_needed.lower() == 'l':
        print("Skipping scope setup. Ensure channels are configured correctly before starting data acquisition.")
    else:
        setup_scope(connected_instrument, num_channels_to_monitor)

    # Create a data file for logging
    last_state_change_time = datetime.datetime.now()
    paths = make_datafile(last_state_change_time, DESKTOP)
    user_path = paths[0]
    datafile_name = paths[1]
    full_data_path = os.path.join(user_path, datafile_name)
    print("Created file for data as ", datafile_name)

    # Setting voltage thresholds for ON and OFF states
    print(f"Monitoring for ON (all channels > {ON_THRESHOLD:.2f}Vrms) and OFF (all channels < {OFF_THRESHOLD:.2f}Vrms) states.")
    print("Press 'q' or 'Crtl-C' to stop the program at any time.")
    print("Starting monitoring...")

    # Main loop
    while not stop_program_event.is_set():
        time.sleep(0.1) # Small delay for keyboard input before checking if scope triggered
        Status = connected_instrument.query('ACQuire:STATE?').strip()

        if Status == '0' :  
            # Scope triggered; turn off by setting super high level
            connected_instrument.write("ACQuire:STATE OFF")
            connected_instrument.write("TRIGger:A:LEVel:CH1 5")  # reset trigger level
            print("Status = " , Status, ". Scope triggered.")
            event_counter += 1
            print("Trigger count- ", event_counter)
            current_time = datetime.datetime.now()

            # Read measurements
            v_rms_readings = []
            for i in range(1, num_channels_to_monitor + 1):
                try:
                    v_rms = float(connected_instrument.query(f"MEASUrement:MEAS{i}:VALue?"))
                    # Apply bounds using the new routine call
                    v_rms = apply_vrms_bounds(v_rms)
                    v_rms_readings.append(v_rms)
                except pyvisa.errors.VisaIOError as e:
                    print(f"Error reading RMS for Channel {i}: {e}. Skipping this channel for this sample.")
                    v_rms_readings.append(float('NAN')) # Append NaN if reading fails
                except ValueError:
                    print(f"Could not convert RMS reading for Channel {i} to float. Skipping.")
                    v_rms_readings.append(float('NAN'))

            # Check if any readings are NaN, if so, we can't determine state reliably
            if any(v == float('NAN') for v in v_rms_readings):
                print("Warning: Skipping state evaluation due to invalid Vrms readings.")
                continue

            # Check state change and log to file
            all_channels_on = all(v >= ON_THRESHOLD for v in v_rms_readings)
            all_channels_off = all(v <= OFF_THRESHOLD for v in v_rms_readings)

            if current_state == "UNKNOWN":
                if all_channels_on:
                    new_state = "ON"
                    current_state = new_state
                elif all_channels_off:
                    new_state = "OFF"
                    current_state = new_state

            if current_state == "OFF":
                if all_channels_on:
                    # Transition from OFF to ON
                    new_state = "ON"
                    duration = (current_time - last_state_change_time).total_seconds()
                    event_counter += 1
                    log_duration_to_file(user_path, datafile_name, event_counter, last_state_change_time, current_time, "OFF", duration)
                    print(f"State Change: OFF to ON. Previous OFF duration: {duration:.3f} seconds.")
                    current_state = new_state
                    last_state_change_time = current_time

            elif current_state == "ON":
                if all_channels_off:
                    # Transition from ON to OFF
                    new_state = "OFF"
                    duration = (current_time - last_state_change_time).total_seconds()
                    event_counter += 1
                    log_duration_to_file(user_path, datafile_name, event_counter, last_state_change_time, current_time, "ON", duration)
                    print(f"State Change: ON to OFF. Previous ON duration: {duration:.3f} seconds.")
                    current_state = new_state
                    last_state_change_time = current_time

            # Print readings for user 
            print_output = f"Current Readings ({current_time.strftime('%H:%M:%S.%f')}): "
            for i, v_rms in enumerate(v_rms_readings):
                print_output += f"CH{i+1}: {v_rms:6.3f}Vrms "
            print(print_output + f" -> State: {current_state}")

            # ready next trigger
            connected_instrument.write("ACQuire:STATE ON")
            time.sleep(1)  # wait before checking again

        elif Status == '1': # Still awaiting trigger
            print ("not triggered")
            connected_instrument.write("ACQuire:MODe SAMPLE")
            connected_instrument.write("ACQuire:STOPAfter SEQuence")
            connected_instrument.write("TRIGger:A:LEVel:CH1 2.0") # Revise trigger level and verify triggered
            connected_instrument.write("ACQuire:STATE ON")
            time.sleep(2)

except KeyboardInterrupt:
    print("\nProgram terminated by user (Ctrl+C).")
except Exception as e:
    print(f"An error occurred during program execution: {e}")
finally:
    # Always close the instrument connection and resource manager
    if 'connected_instrument' in locals() and connected_instrument:
        try:
            connected_instrument.write("ACQuire:STATE OFF") # Stop acquisition before closing
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
    if current_state != "UNKNOWN":
        final_time = datetime.datetime.now()
        duration = (final_time - last_state_change_time).total_seconds()
        event_counter += 1 # Increment for the final state duration
        log_duration_to_file(user_path, datafile_name, event_counter, last_state_change_time, final_time, current_state, duration)
        print(f"Program stopped. Final {current_state} duration: {duration:.3f} seconds.")
