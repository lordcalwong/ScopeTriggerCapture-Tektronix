# For MSO58 Series Scope
# Connect to scope to set up, trigger and wait, and 
# save measurements and image at triggered event.


import time
import datetime
import os
import pyvisa
from tm_devices import DeviceManager
from tm_devices.drivers import MSO5B   # CHANGE FOR YOUR PARTICULAR SCOPE MODEL
# Ensure TM_OPTIONS is set for standalone operation if not using
# pyvisa.ResourceManager('@py').  This needs to be set before importing 
# tm_devices for the first time in a session.
# if "TM_OPTIONS" not in os.environ:
#     os.environ["TM_OPTIONS"] = "STANDALONE"


def setup_scope(scope_device: MSO5B):
    """
    Configures the oscilloscope settings for measurement.

    Args:
        scope_device: An instance of the MSO5B oscilloscope device.
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
    print(f"Count: {counter}, Vpk2pk: {v_peak_to_peak:.3f}, Vrms: {v_rms:.3f}")

    # Define a temporary path on the instrument's internal drive
    temp_image_path_on_scope = "C:/Temp.png"
    scope_device.write(f'SAVE:IMAGe "{temp_image_path_on_scope}"')

    # Wait for the instrument to finish writing the image to its disk
    scope_device.query('*OPC?') # Operation Complete query

    # Read image file from instrument into memory
    image_data = scope_device.read_raw()

    # Generate a unique filename for the image on the local disk
    image_file_name = os.path.join(save_directory, f"{current_dt.strftime('%Y%m%d_%H%M%S')}.png")
    print(f'Saving image to: {image_file_name}')

    # Save image data to local disk
    try:
        with open(image_file_name, "wb") as f:
            f.write(image_data)
    except IOError as e:
        print(f"Error saving image file '{image_file_name}': {e}")
        # Return current counter and datetime if image saving fails
        return counter, current_dt

    # --- Data Logging ---
    # Append measured data to the data file
    try:
        data_full_path = os.path.join(save_directory, data_file_name)
        with open(data_full_path, "a") as f:
            f.write(f"{counter:4.0f}, {current_dt.hour:02d}.{current_dt.minute:02d}.{current_dt.second:02d}, {v_peak_to_peak:.3f}, {v_rms:.3f}\n")
    except IOError as e:
        print(f"Error appending data to file '{data_full_path}': {e}")

    # Clear output buffers
    scope_device.device_clear()

    return counter + 1, current_dt

if __name__ == "__main__":
    # Initialize PyVISA Resource Manager
    # Use '@py' for the PyVISA-py backend
    resource_manager = pyvisa.ResourceManager('@py')

    # List available resources for debugging/verification
    print("\nAvailable VISA Resources:")
    equipment_list = resource_manager.list_resources()
    for resource in equipment_list:
        print(resource)
    print("-" * 30)

    # --- Configuration ---
    # Modify these lines to configure the script for your needs/instrument
    # Use raw string (r"...") for Windows paths to avoid issues with backslashes
    # CHANGE FOR YOUR PARTICULAR SCOPE IP ADDRESS, e.g., 10.101.100.236, 10.101.100.254, 10.101.100.151 
    VISA_RESOURCE_ADDRESS = '10.101.100.151'  
    SAVE_PATH = r"C:\Users\Calvert.Wong\OneDrive - qsc.com\Desktop\ScopeData" # Ensure this directory exists or create it

    # Create save directory if it doesn't exist
    os.makedirs(SAVE_PATH, exist_ok=True)
    print(f"Data and images will be saved to: {SAVE_PATH}")
    print("-" * 30)

    trigger_counter = 0

    try:
        # Use DeviceManager for robust connection management
        # The 'with' statement ensures the connection is closed automatically
        with DeviceManager(verbose=True) as device_manager:
            # Add your scope device. Ensure the driver matches your specific model (e.g., MSO5B)
            scope: MSO5B = device_manager.add_scope(VISA_RESOURCE_ADDRESS)
            print(f"\nConnected to: {scope.idn_string}")

            # Set up scope capture
            setup_scope(scope)

            # Generate a filename for the data based on the current Date
            data_log_file_name = datetime.datetime.now().strftime("%Y%m%d.txt")
            full_data_path = os.path.join(SAVE_PATH, data_log_file_name)

            # Create data file with header if it doesn't exist
            if not os.path.exists(full_data_path):
                print(f"Creating new data file: {full_data_path}")
                with open(full_data_path, "w") as datafile:
                    datafile.write("Count, Time, Vpk2pk, Vrms\n")
            else:
                print(f"Appending to existing data file: {full_data_path}")

            # Initial trigger to start acquisition
            scope.write("ACQUIRE:STATE 1")
            print("Scope acquisition started. Waiting for triggers...")

            # Trigger Capture Loop
            while True:
                # Poll acquisition state. Use .strip() to remove whitespace/newline characters.
                acquisition_status = scope.query('ACQUIRE:STATE?').strip()

                if acquisition_status == '0':  # Scope triggered and acquisition complete
                    print("Triggered!")
                    trigger_counter, _ = capture_data_and_image(
                        scope, SAVE_PATH, data_log_file_name, trigger_counter
                    )

                    # Allow time for saving before re-triggering for the next sequence
                    time.sleep(2)
                    # Re-arm the acquisition for the next trigger
                    scope.write("ACQUIRE:STATE 1")
                    time.sleep(1) # Small delay to allow scope to re-arm

                else:  # Still waiting for a trigger
                    print("Not triggered yet, waiting...")
                    time.sleep(1) # Wait before polling again

    except pyvisa.errors.VisaIOError as e:
        print(f"VISA I/O Error: {e}")
        print("Please ensure the oscilloscope is connected and the IP address is correct.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
    finally:
        # Resource Manager is closed here even if an error occurs
        if 'resource_manager' in locals() and resource_manager:
            print("\nClosing VISA Resource Manager.")
            resource_manager.close()
        print("Script finished.")
