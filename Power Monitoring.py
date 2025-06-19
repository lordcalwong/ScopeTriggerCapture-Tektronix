# Power Monitoring- Synchronous
#
# Continuous monitor Amplifier Output channels and/or Line Inputs with
# asynchronous ON and OFF time logging. Configure all measurements on scope
# to be # RMS readings (~10V/div and ~trigger level of 10V).
# Trigger level is 18W/ch (9.8V) ON and 1V/ch OFF.
# autorun.
# Keeps track of number of valid cycles. 
# Uses pyvisa for generic scope SCPI communications.
#
# Author: C. Wong 2025XXXX

import time
import datetime
import os
import keyboard
import pyvisa
import threading

# Configure IP '192.168.1.53', '10.101.100.151', '10.101.100.236', '10.101.100.254', '10.101.100.176'
DEFAULT_IP_ADDRESS = '192.168.1.53'   # CHANGE FOR YOUR PARTICULAR SCOPE!
SAVE_PATH = r"C:\Users\Calvert.Wong\OneDrive - qsc.com\Desktop"
MIN_ACQUISITION_INTERVAL = 5   # sampling rate
MAX_VRMS = 300

# Global flag to signal the main loop and threads to stop
stop_program_event = threading.Event()

# --- FUNCTION DEFINITION ---
def timer_thread_func(event_to_set: threading.Event, interval: float, stop_event: threading.Event):
    """
    Separate thread that signals an event after the specified interval.
    """
    while not stop_event.is_set():
        # Wait for the interval, but allow to be interrupted by stop_event
        # If stop_event is set during this wait, it will return True immediately.
        if stop_event.wait(interval):
            break # Exit the loop if stop_event was set
        # If the wait completed (meaning interval passed and stop_event wasn't set),
        # then set the event for the main thread.
        if not stop_event.is_set(): 
            event_to_set.set()

def on_q_press():
    """
    Callback function when 'q' is pressed.
    """
    print("\n'q' pressed. Signaling program to stop.")
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
            f"Enter the instrument's IP address (e.g., {default_ip}) or 'd' for default: "
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

def get_max_graph_voltage():
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

def sample_period():
    """
    Asks user to input the time between samples.
    """
    while True:
        try:
            sample_period = input("Enter time between samples in seconds (e.g., 1-300 or {MIN_ACQUISITION_INTERVAL}) or 'd' for default: ").strip()
            if sample_period.lower() == 'd':
                return MIN_ACQUISITION_INTERVAL
            else:
                period = int(sample_period)
                if 1 <= period <= 300:
                    return period
        except ValueError:
            print("Invalid input. Please enter a number between 1 and 300.")

def setup_scope(scope_device, num_channels):
    """
    Configures channels for the specified number of channels.
    Minimally tries to set scale and position.
    """
    print("Setting up oscilloscope...", end='')
    
    # # Clear existing measurements to avoid conflicts
    # scope_device.write("MEASUrement:CLEarALL")

    for i in range(1, num_channels + 1):
        scope_device.write(f"SELect:CH{i} ON")
        scope_device.write(f"CH{i}:SCALe 10")
        scope_device.write(f"CH{i}:POSition -4")
        # scope_device.write(f"MEASUrement:MEAS{i}:ENABle ON")
        # scope_device.write(f"MEASUrement:MEAS{i}:SOUrce1 CH{i}; STATE 1; TYPE RMS")  #works for DPO4K
        scope_device.write(f"MEASUrement:MEAS{i}:SOUrce CH{i}; TYPE RMS")  #works for MSO5

    # Wait for scope to finish setting up
    scope_device.query("*OPC?")
    print("Scope setup complete.")

def make_datafile(num_channels, timestamp):
    """Generates a data file for the data based on start date and time.
    Includes headers for each monitored channel.
    """
    data_log_file_name = timestamp.strftime("%Y%m%d_%H%M%S.txt")
    full_data_path = os.path.join(SAVE_PATH, data_log_file_name)
    if not os.path.exists(full_data_path):
        print(f"Creating new data file: {full_data_path}")
        with open(full_data_path, "w") as datafile:
            header = "Count, Time"
            for i in range(1, num_channels + 1):
                header += f", Vrms_CH{i}"
            datafile.write(header + "\n")
    else:
        print(f"Appending to existing data file: {full_data_path}")
    return data_log_file_name

def add_sample_to_file(save_directory, data_file_name, counter, time_in_seconds, v_rms_values):
    """
    Appends sample data for all channels to the specified file.
    """
    try:
        datafile_and_path = os.path.join(save_directory, data_file_name)
        with open(datafile_and_path, "a") as f:
            # Start with count and time
            line = f"{counter:4d}, {time_in_seconds:9.3f}"
            # Add each Vrms value
            for v_rms in v_rms_values:
                line += f", {v_rms:6.3f}"
            f.write(line + "\n")
    except IOError as e:
        print(f"Error appending data to file '{datafile_and_path}': {e}")

def apply_vrms_bounds(v_rms):
    """
    Applies upper and lower bounds to the Vrms reading.
    """
    return max(min(v_rms, MAX_VRMS), 0)

# --- MAIN ---
if __name__ == "__main__":
    # Initialize the Resource Manager
    rm = pyvisa.ResourceManager('@py')
    print(rm.list_resources())

    sample_time = sample_period()

    # Create Event object for sampling time and start timer
    acquisition_allowed_event = threading.Event()
    timer_thread = threading.Thread(
        target=timer_thread_func,
        args=(acquisition_allowed_event, sample_time, stop_program_event),
        daemon=True # Daemon threads exit automatically when the main program exits
    )
    timer_thread.start()

    # Register the 'q' hotkey
    keyboard.add_hotkey('q', on_q_press)
    
    num_channels_to_monitor = 0
    connected_instrument = None
    try:
        # Call the new function to connect to the instrument
        connected_instrument = connect_to_instrument(rm, DEFAULT_IP_ADDRESS)

        # Get the number of channels from the user
        num_channels_to_monitor = get_num_channels()

        # Set up channels based on the user input
        setup_scope(connected_instrument, num_channels_to_monitor)

        # Create a data file for logging
        starting_date_and_time = datetime.datetime.now()
        datafile = make_datafile(num_channels_to_monitor, starting_date_and_time)
        print("Created file for data as ", datafile)
        
        count = 1
        # Main loop
        while not stop_program_event.is_set():
            # Wait for the minimum acquisition interval to pass before arming
            acquisition_allowed_event.wait(timeout=0.1) # This will block until the timer thread sets the event
            if stop_program_event.is_set(): # Check immediately after waiting
                break
            if not acquisition_allowed_event.is_set():
                # Continue waiting for the interval to pass or q to be pressed.
                continue
            acquisition_allowed_event.clear() # Reset the event for the next cycle

            # OK to sample
            current_date_and_time = datetime.datetime.now()
            dt = current_date_and_time - starting_date_and_time
            dt_in_seconds = dt.total_seconds()

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

            add_sample_to_file(SAVE_PATH, datafile, count, dt_in_seconds, v_rms_readings)
            
            # Print the current sample data
            print_output = f"Sample {count:4d},   Time: {dt_in_seconds:9.3f} sec"
            for i, v_rms in enumerate(v_rms_readings):
                print_output += f", CH{i+1}: {v_rms:6.3f}"
            print(print_output)
            
            # Increment the counter
            count += 1

            # Check if 'q' was pressed taking samples 
            if stop_program_event.is_set():
                # If 'q' was pressed, stop the acquisition on the scope
                connected_instrument.write("CLEAR")
                break

    except Exception as e:
        print(f"An error occurred during program execution: {e}")
    finally:
        # Always close the instrument connection and resource manager
        if 'connected_instrument' in locals() and connected_instrument:
            print("Closing instrument connection.")
            connected_instrument.close()
        if rm:
            print("Closing Resource Manager.")
            rm.close()
