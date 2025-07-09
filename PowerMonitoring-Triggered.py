# Power Monitoring- Synchronous
#
# Continuous monitor Amplifier Output channels and/or Line Inputs with
# synchronous logging. Trigger mode should be in autorun.
#
# User inputs IP address, sample time in seconds with default of 5 seconds,
# and number of channels to monitor (1-8).
#
# User has option to set up scope or allow continous channels to be configured
# for RMS measurements.
#
# Currently, the maximum voltage is set to 50Vp or about 300W/ch.
# Uses pyvisa for generic scope SCPI communications for both DPO4k and MSO58
# series scopes.
#
# Saves data to csv file, closes file, and imports csv into MS Excel file
# and plots a chart.
#
# Author: C. Wong 20250703

import time
import datetime
import os
import keyboard
import pyvisa
import threading
import csv

from openpyxl import Workbook
from openpyxl.drawing.text import Paragraph, CharacterProperties, Font
from openpyxl.styles import Font as ExcelFont
from openpyxl.chart import ScatterChart, Reference, Series
from openpyxl.drawing.line import LineProperties
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.colors import ColorChoice

DEFAULT_IP_ADDRESS = '192.168.1.53'  #default IP, 192.168.1.53, 10.101.100.151
MAX_VRMS = 50

# --- NEW: Voltage Thresholds ---
ON_THRESHOLD = 5.0  # Vrms - all channels must be above this to be considered 'ON'
OFF_THRESHOLD = 1.0 # Vrms - all channels must be below this to be considered 'OFF'
# -----------------------------

# Find user esktop one level down from home (~/* /Desktop) and set up as optional save path
from glob import glob
DESKTOP = glob(os.path.expanduser("~\\*\\Desktop"))
# If not found, revert to the standard location (~/Desktop)
if len(DESKTOP) == 0:
    DESKTOP = os.path.expanduser("~\\Desktop")
else:
    DESKTOP = DESKTOP[0] # glob returns a list, take the first result

# Global flag to signal the main loop and threads to stop
stop_program_event = threading.Event()

# --- FUNCTION DEFINITION ---
# REMOVED: timer_thread_func - No longer needed for threshold-based sampling

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

# REMOVED: sample_period - No longer applies

def setup_scope(scope_device, num_channels):
    """
    Configures channels for the specified number of channels.
    Minimally tries to set scale and position.
    """
    print("Ok. Setting up oscilloscope...", end='')

    for i in range(1, num_channels + 1):
        scope_device.write(f"SELect:CH{i} ON")
        scope_device.write(f"CH{i}:SCALe 10")
        scope_device.write(f"CH{i}:POSition 0")
        scope_device.write(f"MEASUrement:MEAS{i}:SOUrce CH{i}; STATE 1")   # Need separate STATE 1 command for DPO
        scope_device.write(f"MEASUrement:MEAS{i}:SOUrce CH{i}; TYPE RMS")  # Need separate TYPE for MSO

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

def apply_vrms_bounds(v_rms):
    """
    Applies upper and lower bounds to the Vrms reading.
    """
    return max(min(v_rms, MAX_VRMS), 0)

def write_to_excel_with_chart(datafile_name: str, save_directory: str): # num_channels removed
    """
    Reads data from the specified CSV file, writes it to an Excel worksheet,
    and creates a scatter chart for durations.

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
        print("Data successfully written to Excel worksheet.")


        # --- Charting Section ---
        chart = ScatterChart()
        chart.title = "State Durations Over Time"
        chart.style = 10
        chart.x_axis.title = "Event Number" # Using event number for X-axis
        chart.y_axis.title = "Duration (seconds)"
        max_row = ws.max_row

        # Add ON durations
        x_values_on = Reference(ws, min_col=1, min_row=2, max_row=max_row, max_col=1) # Event_Count
        y_values_on = Reference(ws, min_col=5, min_row=2, max_row=max_row, max_col=5) # Duration_Seconds
        # Filter for "ON" state
        series_on = Series(y_values_on, x_values_on, title="ON Duration")
        # Customizing series to only show 'ON' data is tricky with direct openpyxl references
        # A more advanced approach would involve creating separate filtered lists in Python
        # For simplicity, this chart will plot all durations, and the user can filter in Excel.
        # Or, we can iterate and add series conditionally based on the 'State' column.

        # Let's create two series, one for ON and one for OFF, based on the state column.
        # This requires reading the data into lists first.
        on_durations_for_chart = []
        off_durations_for_chart = []
        event_counts_on = []
        event_counts_off = []

        # Read data again to filter for chart series
        with open(full_csv_path, 'r') as f:
            reader = csv.reader(f)
            next(reader) # Skip header
            for row in reader:
                try:
                    event_count = int(row[0])
                    state = row[3]
                    duration = float(row[4])
                    if state == "ON":
                        event_counts_on.append(event_count)
                        on_durations_for_chart.append(duration)
                    elif state == "OFF":
                        event_counts_off.append(event_count)
                        off_durations_for_chart.append(duration)
                except (ValueError, IndexError):
                    continue # Skip malformed rows

        # Create new sheets or ranges for chart data if filtering heavily
        # For minimal change, let's just make one series and rely on Excel for filtering or manual separation.
        # A simple scatter of all durations is the easiest given the 'minimal change' constraint.

        # Simpler approach: Just plot all durations, with different colors for ON/OFF
        # This still requires iterating over the data
        chart_data_rows = []
        # Re-read data from the Excel sheet, skipping header (row 1)
        for r_idx in range(2, max_row + 1):
            row_data = []
            for c_idx in range(1, ws.max_column + 1):
                row_data.append(ws.cell(row=r_idx, column=c_idx).value)
            chart_data_rows.append(row_data)

        on_events_idx = []
        on_durations = []
        off_events_idx = []
        off_durations = []

        for r_idx, row in enumerate(chart_data_rows):
            event_count = row[0]
            state = row[3]
            duration = row[4]
            if state == "ON" and duration is not None:
                on_events_idx.append(event_count)
                on_durations.append(duration)
            elif state == "OFF" and duration is not None:
                off_events_idx.append(event_count)
                off_durations.append(duration)

        # Write filtered data to temporary columns or directly use in chart with inline data if possible
        # This is beyond "minimal changes" to the Excel charting.
        # The easiest is to just plot a single series of all durations, and the 'State' column will be available in the table.

        # Reverting to the simplest chart: plot all durations against event number.
        # This plots every row's duration. The state info is in the table.
        x_values = Reference(ws, min_col=1, min_row=2, max_row=max_row) # Event_Count
        y_values = Reference(ws, min_col=5, min_row=2, max_row=max_row) # Duration_Seconds
        series = Series(y_values, x_values, title="All Durations")
        chart.series.append(series)

        # Add the chart to the worksheet
        ws.add_chart(chart, "F2") # Adjust cell to place the chart as needed
        # Ensure axes are not deleted
        chart.x_axis.delete = False
        chart.y_axis.delete = False

        # Access legend's graphical properties to add fill and outline
        if chart.legend: # Ensure legend exists before trying to style it
            legend_spPr = GraphicalProperties()
            # Using preset colors directly within ColorChoice for solidFill
            legend_spPr.solidFill = ColorChoice(prstClr='white')  # White fill
            line_props = LineProperties()
            line_props.solidFill = ColorChoice(prstClr='black') # Black outline
            line_props.width = 12700 # 1 pt in EMU (English Metric Units), 12700 EMU = 1 pt
            legend_spPr.line = line_props
            chart.legend.spPr = legend_spPr
        # --- End Charting Section ---

        wb.save(full_excel_path)
        print(f"Excel data and chart saved successfully to: {full_excel_path}")

    except FileNotFoundError:
        print(f"Error: CSV data file not found at {full_csv_path}. Cannot create Excel file.")
    except Exception as e:
        print(f"An error occurred while creating the Excel file: {e}")

# --- MAIN ---
if __name__ == "__main__":
    # Initialize the Resource Manager
    rm = pyvisa.ResourceManager('@py')
    print("Resources found " , rm.list_resources())

    # REMOVED: sample_time = sample_period(MIN_ACQUISITION_INTERVAL)
    # REMOVED: acquisition_allowed_event and timer_thread

    # Register the 'q' hotkey
    keyboard.add_hotkey('q', on_q_press)
    keyboard.add_hotkey('esc', on_esc_press)

    num_channels_to_monitor = 0
    connected_instrument = None
    datafile_name = None # Initialize datafile_name to None

    # --- NEW State Management Variables ---
    current_state = "UNKNOWN" # Can be "ON" or "OFF"
    last_state_change_time = datetime.datetime.now()
    event_counter = 0
    # --------------------------------------

    try:
        # Call the new function to connect to the instrument
        connected_instrument = connect_to_instrument(rm, DEFAULT_IP_ADDRESS)

        # Get the number of channels from the user
        num_channels_to_monitor = get_num_channels()

        # Set up channels based on the user input
        setup_needed = input("If answering no, I will attempt to setup contiguous channels (CH1 throug X) and measurements.  Leave scope alone (y/n)?:").strip()
        if setup_needed.lower() == 'y':
            print("Skipping scope setup. Ensure channels are configured correctly before starting data acquisition.")
        else:
            setup_scope(connected_instrument, num_channels_to_monitor)

        # Create a data file for logging
        starting_date_and_time = datetime.datetime.now()
        paths = make_datafile(starting_date_and_time, DESKTOP) # num_channels removed
        user_path = paths[0]
        datafile_name = paths[1]
        full_data_path = os.path.join(user_path, datafile_name)
        print("Created file for data as ", datafile_name)

        print(f"Monitoring for ON (all channels > {ON_THRESHOLD:.2f}Vrms) and OFF (all channels < {OFF_THRESHOLD:.2f}Vrms) states.")
        print("Press 'q' or 'Crtl-C' to stop the program at any time.")
        print("Starting monitoring...")

        # Main loop
        while not stop_program_event.is_set():
            time.sleep(0.1) # Small delay to prevent busy-waiting and allow for keyboard input

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

            all_channels_on = all(v >= ON_THRESHOLD for v in v_rms_readings)
            all_channels_off = all(v <= OFF_THRESHOLD for v in v_rms_readings)
            current_time = datetime.datetime.now()

            new_state = current_state # Assume state doesn't change unless conditions met

            if current_state == "UNKNOWN":
                if all_channels_on:
                    new_state = "ON"
                elif all_channels_off:
                    new_state = "OFF"
                # If still UNKNOWN (e.g., in transition), just keep polling
                if new_state != "UNKNOWN":
                    current_state = new_state
                    last_state_change_time = current_time
                    print(f"Initial state determined: {current_state} at {current_time.strftime('%Y-%m-%d %H:%M:%S.%f')}")

            elif current_state == "OFF":
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

            # Print current readings for monitoring purposes
            print_output = f"Current Readings ({current_time.strftime('%H:%M:%S.%f')}): "
            for i, v_rms in enumerate(v_rms_readings):
                print_output += f"CH{i+1}: {v_rms:6.3f}Vrms "
            print(print_output + f" -> State: {current_state}")

    except Exception as e:
        print(f"An error occurred during program execution: {e}")
    finally:
        # Always close the instrument connection and resource manager
        if 'connected_instrument' in locals() and connected_instrument:
            print("Closing instrument connection.")
            connected_instrument.write("CLEAR") # Ensure scope acquisition is stopped
            connected_instrument.close()
        if rm:
            print("Closing Resource Manager.")
            rm.close()

        # Before exiting, log the duration of the final state if it was not already logged
        if current_state != "UNKNOWN":
            final_time = datetime.datetime.now()
            duration = (final_time - last_state_change_time).total_seconds()
            event_counter += 1 # Increment for the final state duration
            log_duration_to_file(user_path, datafile_name, event_counter, last_state_change_time, final_time, current_state, duration)
            print(f"Program stopped. Final {current_state} duration: {duration:.3f} seconds.")


        # After data acquisition stops, write to Excel if a datafile was created
        if datafile_name: # No longer dependent on num_channels_to_monitor > 0 for this logic
            write_to_excel_with_chart(datafile_name, user_path)