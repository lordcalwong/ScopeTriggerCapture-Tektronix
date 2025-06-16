# For MSO58 Series Scope
# Connect to scope to set up, trigger, and save image.

import time
import datetime
import os
import pyvisa
from tm_devices import DeviceManager
from tm_devices.drivers import MSO5B   # CHANGE FOR YOUR PARTICULAR SCOPE

# Configure visaResourceAddr, e.g., 'TCPIP::10.101.100.236::INSTR',  '10.101.100.236', '10.101.100.254', '10.101.100.176'
visaResourceAddr = '10.101.100.151'   # CHANGE FOR YOUR PARTICULAR SCOPE!
SAVE_PATH = r"C:\Users\Calvert.Wong\OneDrive - qsc.com\Desktop\ScopeData" # Ensure this directory exists or create it

# ************** MAIN    
counter = 0  # trigger counter to track data record
with DeviceManager(verbose=True) as device_manager:
    
    scope:MSO5B = device_manager.add_scope(visaResourceAddr)  # CHANGE FOR YOUR PARTICULAR SCOPE USING Intellisense!
    print(scope.idn_string)
    
    # Generate a filename based on the current Date & Time
    dt = datetime.datetime.now()
    fileName = dt.strftime("%Y%m%d.txt")

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

    
    # Trigger ready level, mode, and set trigger
    scope.write("TRIGger:A:LEVel:CH1 5")
    scope.write("ACQuire:STOPAfter SEQUENCE")
    scope.write("ACQuire:STATE ON")


    # Trigger Capture Loop
    while (True):
        # Check if scope is triggered
        status = scope.query('ACQuire:STATE?')
        # print(f"Scope status: {status}")

        if status == '0' :  
            # pause and write measurement stats
            scope.write("ACQuire:STATE OFF")
            scope.write("TRIGger:A:LEVel:CH1 5")  # reset trigger level
            status = 1
            counter += 1
            print()
            print("Trigger count- ", counter)

            dt = datetime.datetime.now()
            Vp2p = float(scope.query("MEASUrement:MEAS1:VALue?"))
            Vrms = float(scope.query("MEASUrement:MEAS2:VALue?"))
            if Vp2p > 1000:
                Vp2p =999.999
            if Vrms > 1000:
                Vrms =999.999     
            print(f"counter: {counter} Vpk2pk: {Vp2p:.3f}, Vrms: {Vrms:.3f}")
            with open(os.path.join(SAVE_PATH, fileName), "a") as datafile:  # append mode
                datafile.write(f"{counter:4.0f}, {dt.hour:02d}.{dt.minute:02d}.{dt.second:02d}, {Vp2p:.3f}, {Vrms:.3f}\n")
                datafile.close()

            # ready next trigger
            scope.write("ACQuire:STATE ON")
            time.sleep(1)  # wait before checking again
        else:
            # print ("not triggered")
            scope.write("TRIGger:A:LEVel:CH1 2.0")
            scope.write("ACQuire:STATE ON")
            time.sleep(1)  # wait before checking again

rm.close()

