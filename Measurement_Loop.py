# For MSO58 Series Scope
# Connect to scope to set up, trigger, and save image.

import time
import datetime
import os
import pyvisa

DEFAULT_IP_ADDRESS = '192.168.1.53'  #10.101.100.151, 10.101.100.236, 10.101.100.254, 10.101.100.176
SAVE_PATH = r"C:\Users\Calvert.Wong\OneDrive - qsc.com\Desktop\ScopeData" # Ensure this directory exists or create it

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

def create_data_file_header(SAVE_PATH, fileName):
    # Create data file with header info (create directory if it doesn't exist).
    try:
        os.makedirs(SAVE_PATH, exist_ok=True)
        with open(os.path.join(SAVE_PATH, fileName), "w") as datafile:
            datafile.write("Count, Time, Vpk2pk, Vrms\n")
        # The 'with' statement automatically handles closing the file,
        # even if errors occur within the block.
    except OSError as e:
        # Catch OS-related errors (e.g., permission denied, disk full)
        print(f"Error creating or writing to file: {e}")
    except Exception as e:
        # Catch any other unexpected errors
        print(f"An unexpected error occurred: {e}")

# ************** MAIN    
counter = 0  # trigger counter to track data record
rm = pyvisa.ResourceManager()
connected_instrument = connect_to_instrument(rm, DEFAULT_IP_ADDRESS)
    
# Generate a filename based on the current Date & Time
dt = datetime.datetime.now()
fileName = dt.strftime("%Y%m%d.txt")

# Trigger ready level, mode, and set trigger
connected_instrument.write("TRIGger:A:LEVel:CH1 5")
connected_instrument.write("ACQuire:STOPAfter SEQUENCE")
connected_instrument.write("ACQuire:STATE ON")


# Trigger Capture Loop
while (True):
    # Check if scope is triggered
    status = connected_instrument.query('ACQuire:STATE?')
    # print(f"Scope status: {status}")

    if status == '0' :  
        # pause and write measurement stats
        connected_instrument.write("ACQuire:STATE OFF")
        connected_instrument.write("TRIGger:A:LEVel:CH1 5")  # reset trigger level
        status = 1
        counter += 1
        print()
        print("Trigger count- ", counter)

        dt = datetime.datetime.now()
        Vp2p = float(connected_instrument.query("MEASUrement:MEAS1:VALue?"))
        Vrms = float(connected_instrument.query("MEASUrement:MEAS2:VALue?"))
        if Vp2p > 1000:
            Vp2p =999.999
        if Vrms > 1000:
            Vrms =999.999     
        print(f"counter: {counter} Vpk2pk: {Vp2p:.3f}, Vrms: {Vrms:.3f}")
        with open(os.path.join(SAVE_PATH, fileName), "a") as datafile:  # append mode
            datafile.write(f"{counter:4.0f}, {dt.hour:02d}.{dt.minute:02d}.{dt.second:02d}, {Vp2p:.3f}, {Vrms:.3f}\n")
            datafile.close()

        # ready next trigger
        connected_instrument.write("ACQuire:STATE ON")
        time.sleep(1)  # wait before checking again
    else:
        # print ("not triggered")
        connected_instrument.write("TRIGger:A:LEVel:CH1 2.0")
        connected_instrument.write("ACQuire:STATE ON")
        time.sleep(1)  # wait before checking again


