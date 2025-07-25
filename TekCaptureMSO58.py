# For MSO58 Series Scope   (MSO5B and higher, not DPO4KB)
# Connect to scope to set up, trigger and wait, and save measurements and image
# at triggered events.
# Currently speed is limited to 1 second with constant MIN_ACQUISITION_INTERVAL
# Also, scope setup is programmatically setup with routine setup_scope() which
# can be commented out if you rather want to use the scope's front panel.

import time
import datetime
import os
import keyboard
import pyvisa
import threading # Import the threading module
from tm_devices import DeviceManager
from tm_devices.drivers import MSO5B

# Global flag to signal the main loop and threads to stop
stop_program_event = threading.Event()

def setup_scope(scope_device: MSO5B):
    """
    Configures the oscilloscope settings for measurement.
    """
    print("Setting up oscilloscope...")
    scope_device.write("SELect:CH1 ON")
    scope_device.write("SELect:CH2 OFF; CH3 OFF; CH4 OFF; CH5 OFF; CH6 OFF; CH7 OFF; CH8 OFF")
    scope_device.write('CH1:LABel:NAMe \"CH1 Vout\"')
    scope_device.write("CH1:COUPling DC")
    scope_device.write("CH1:PRObe:GAIN 0.1")
    scope_device.write("CH1:TERmination MEG")
    scope_device.write("CH1:SCALe 0.5")
    scope_device.write("CH1:INVert OFF")
    scope_device.write("CH1:POSition -2")
    scope_device.write("CH1:OFFSet 0")
    scope_device.write("CH1:BANdwidth 250E6")
    scope_device.write("HORizontal:SCAle 200E-6")
    scope_device.write("HORizontal:POSition 50")
    scope_device.write("TRIGger:A:EDGE:SOUrce CH1")
    scope_device.write("TRIGger:A:EDGE:COUPling DC")
    scope_device.write("TRIGger:A:EDGE:SLOpe RISE")
    scope_device.write("TRIGger:A:LEVel:CH1 2.0")  # TRIGGER:A SETLEVEL may be easier for a mid-level trigger
    scope_device.write("TRIGger:A:MODe NORMal")
    scope_device.write("TRIGger:A:TYPe EDGE")
    scope_device.write("HORizontal:MODe AUTO")
    scope_device.write("HORizontal:SAMPLERate:ANALYZemode:MINimum:VALue 250e6")   # set MIN sample rate
    scope_device.write("HORizontal:SAMPLERate:ANALYZemode:MINimum:OVERRide ON")   # but allow horizontal scale override
    scope_device.write("HORizontal:MODe SCALE 400E-6")
    scope_device.write("DISplay:WAVEView1:CH1:VERTical:SCAle 0.5")
    scope_device.write("DISplay:WAVEView1:CURSor:CURSOR1:STATE 0")
    scope_device.write("DISplay:WAVEView1:INTENSITy:GRATicule 100")
    scope_device.write('MEASUrement:DELETE "MEAS1"')
    scope_device.write('MEASUrement:DELETE "MEAS2"')
    scope_device.write("MEASUrement:MEAS1:SOUrce CH1; TYPE PK2Pk")
    scope_device.write("MEASUrement:MEAS2:SOUrce CH1; TYPE RMS")
    scope_device.write("ACQuire:STATE 0")
    scope_device.write("ACQuire:MODe SAMPLE")
    scope_device.write("ACQuire:STOPAfter SEQuence")

    # List of commands that don't work for MSO5 Series
    # device.write("CURSor:FUNCtion OFF")    
    # device.write("MEASUrement:DELETEALL") 
    # device.write("MEASUrement:MEAS1:STATE OFF")

    # Wait for scope to finish setting up
    scope_device.commands.opc.query()
    print("Scope setup complete.")
    

def capture_data_and_image(scope_device, save_directory: str, data_file_name: str, counter: int):
    """
    Captures measurement data and an oscilloscope screen image.

    Args:
        scope_device: An instance of the MSO5B oscilloscope device.
        save_directory: The directory path to save data and images.
        data_file_name: The name of the data file (e.g., "YYYYMMDD.txt").
        counter: The current trigger counter.

    Returns:
        A tuple containing the updated counter and the current datetime object.
    """
    current_dt = datetime.datetime.now()

    # Get measured data
    v_peak_to_peak = float(scope_device.query("MEASUREMENT:MEAS1:VALUE?"))
    v_rms = float(scope_device.query("MEASUREMENT:MEAS2:VALUE?"))
    if v_peak_to_peak > 999:
        v_peak_to_peak = 999
    if v_rms > 999:
        v_rms = 999
    print(f"Count: {counter}, Vpk2pk: {v_peak_to_peak:.3f}, Vrms: {v_rms:.3f}")

    # Append measured data to the data file
    try:
        data_full_path = os.path.join(save_directory, data_file_name)
        with open(data_full_path, "a") as f:
            f.write(f"{counter:4.0f}, {current_dt.hour:02d}:{current_dt.minute:02d}:{current_dt.second:02d}, {v_peak_to_peak:.3f}, {v_rms:.3f}\n")
    except IOError as e:
        print(f"Error appending data to file '{data_full_path}': {e}")

    # Save image to a temporary path on the instrument's internal drive
    temp_image_path_on_scope = "C:/Temp.png"
    scope_device.write(f'SAVE:IMAGe "{temp_image_path_on_scope}"')
    time.sleep(0.2)
    scope_device.write('FILESystem:READfile \"C:/Temp.png\"')
    image_data = scope_device.read_raw()

    # Generate a unique filename for the image on the local disk
    image_file_name = os.path.join(save_directory, f"{current_dt.strftime('%Y%m%d_%H%M%S')}.png")
    print(f'Saving image to: {image_file_name}')

    # Save image data to local disk
    try:
        with open(image_file_name, "wb") as f:
            f.write(image_data)
            f.close()
    except IOError as e:
        print(f"Error saving image file '{image_file_name}': {e}")
        # Return current counter and datetime if image saving fails
        return counter, current_dt
    return counter + 1, current_dt


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
    """Callback function when 'q' is pressed."""
    print("\n'q' pressed. Signaling program to stop.")
    stop_program_event.set()


if __name__ == "__main__":
    # Configure visaResourceAddr, e.g., '192.168.1.53', '10.101.100.151', '10.101.100.236', '10.101.100.254', '10.101.100.176'
    VISA_RESOURCE_ADDRESS = '10.101.100.151'   # CHANGE FOR YOUR PARTICULAR SCOPE!
    SAVE_PATH = r"C:\Users\Calvert.Wong\OneDrive - qsc.com\Desktop\ScopeData" # Ensure this direqctory exists or create it
    MIN_ACQUISITION_INTERVAL = 1.0  #Desired minimum delay time in seconds between acquisitions

    # Create save directory if it doesn't exist
    os.makedirs(SAVE_PATH, exist_ok=True)
    print(f"Data and images will be saved to: {SAVE_PATH}")
    print("-" * 30)

    trigger_counter = 1

    # Create an Event object to signal when the minimum interval has passed
    acquisition_allowed_event = threading.Event()

    # Start the timer thread
    timer_thread = threading.Thread(
        target=timer_thread_func,
        args=(acquisition_allowed_event, MIN_ACQUISITION_INTERVAL, stop_program_event),
        daemon=True # Daemon threads exit automatically when the main program exits
    )
    timer_thread.start()

    # Register the 'q' hotkey
    keyboard.add_hotkey('q', on_q_press)
    
    scope = None  # Initialize scope to None
    device_manager = None # Initialize device_manager to None

    try:
        # Use DeviceManager for robust connection management
        # The 'with' statement ensures the connection is closed automatically
        with DeviceManager(verbose=True) as device_manager:
            scope: MSO5B = device_manager.add_scope(VISA_RESOURCE_ADDRESS)
            print(f"\nConnected to: {scope.idn_string}")

            # Configure scope settings for capture
            setup_scope(scope)

            # Generate a filename for the data based on start date and time
            data_log_file_name = datetime.datetime.now().strftime("%Y%m%d_%H%M%S.txt")
            full_data_path = os.path.join(SAVE_PATH, data_log_file_name)

            # Create data file with header if it doesn't exist
            if not os.path.exists(full_data_path):
                print(f"Creating new data file: {full_data_path}")
                with open(full_data_path, "w") as datafile:
                    datafile.write("Count, Time, Vpk2pk, Vrms\n")
            else:
                print(f"Appending to existing data file: {full_data_path}")

            print("Scope acquisition starting. Press 'q' to quit.")

            # Main loop
            while  not stop_program_event.is_set():
                # Wait for the minimum acquisition interval to pass before arming
                print(f"Waiting for MIN_ACQUISITION_INTERVAL ({MIN_ACQUISITION_INTERVAL}s)...")
                acquisition_allowed_event.wait(timeout=0.1) # This will block until the timer thread sets the event

                if stop_program_event.is_set(): # Check immediately after waiting
                    break

                if not acquisition_allowed_event.is_set():
                    # If the event wasn't set, it means the wait timed out, and we should loop again
                    # and continue waiting for the interval to pass or q to be pressed.
                    continue

                acquisition_allowed_event.clear() # Reset the event for the next cycle

                # Re-arm the acquisition for the next trigger
                scope.write("ACQUIRE:STATE 1")
                print("Scope re-armed. Waiting for trigger...")
                
                # Wait for acquisition to complete or stop signal
                while scope.query('ACQUIRE:STATE?').strip() == '1' and not stop_program_event.is_set():
                    time.sleep(0.05) # Shorter sleep for responsiveness

                if stop_program_event.is_set(): # Check if 'q' was pressed while waiting for trigger
                    # If 'q' was pressed, stop the acquisition on the scope
                    scope.write("ACQUIRE:STATE 0")
                    scope.write("CLEAR")
                    break

                # Triggered event occurred, capture data and image
                trigger_counter, _ = capture_data_and_image(
                    scope, SAVE_PATH, data_log_file_name, trigger_counter
                )

    except pyvisa.errors.VisaIOError as e:
        print(f"VISA I/O Error: {e}")
        print("Ensure oscilloscope is connected and IP address is correct. May need to send CLEAR via Windows PowerShell & ncat or recycle power")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
    finally:
        # Unregister the hotkey to prevent issues after the program exits
        keyboard.unhook_all_hotkeys()
        print("Script finished.")
