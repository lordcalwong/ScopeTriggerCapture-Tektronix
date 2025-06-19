# For DPO4034 Series Scope
# Connect to scope via Ethernet to set up, trigger, and save image/screenshot (hardcopy) to PC for 4 Series MSO Oscilloscopes
# Cal Wong, 2024-01-24

# OS and PyVISA-py
import os       # interact with the operating system and file management
import pyvisa   # control of instruments over wide range of interfaces
rm = pyvisa.ResourceManager('@py')
# List available resources
rm.list_resources()
os.environ["TM_OPTIONS"] = "STANDALONE"


# Use Python device management package from Tektronix
# from tm_devices.helpers import PYVISA_PY_BACKEND, SYSTEM_DEFAULT_VISA_BACKEND
from tm_devices import DeviceManager
from tm_devices.drivers import MSO4     # CHANGE FOR YOUR PARTICULAR SCOPE USING Intellisense!


# Packages used
import time
import os
import datetime
import keyboard


# Configure visaResourceAddr, e.g., '10.101.100.151', '10.101.100.236', '10.101.100.254', '10.101.100.176' 
visaResourceAddr = '192.168.1.53'     # CHANGE FOR YOUR PARTICULAR SCOPE!
savePath = "C:\\Users\\Calvert.Wong\\OneDrive - qsc.com\\Desktop\\DATA"       # CHANGE TO YOUR PREFERRED DESTINATION


def set_up_scope(device):

    # Clear settings
    device.write("*RST")
    device.commands.opc.query()

    # Set up dispaly channels
    device.write("SELect:CH1 ON")
    device.write("SELect:CH2 OFF")
    device.write("SELect:CH3 OFF")
    device.write("SELect:CH4 OFF")
    # Note- "print(scope.commands.select.ch[1],'ON')" doesn't work but should
    
    # Set up timebase
    device.write("HORizontal:SCAle 200E-6")
    device.write("HORizontal:POSition 50")
    
    # Set up trigger
    device.write("TRIGger:A:EDGE:SOUrce CH1")
    device.write("TRIGger:A:EDGE:COUPling DC")
    device.write("TRIGger:A:EDGE:SLOpe RISE")
    device.write("TRIGger:A:LEVel:CH1 2.0")
    # Use "scope.write("TRIGGER:A SETLEVEL")" instead for mid-level"
    device.write("TRIGger:A:MODe NORMal")
    device.write("TRIGger:A:TYPe EDGE")

    # Set up vertical
    device.write("CH1:COUPling DC")
    device.write("CH1:PRObe:GAIN 0.1")
    device.write("CH1:TERmination MEG")
    device.write("CH1:SCALe 1")
    device.write("CH1:INVert OFF")
    device.write("CH1:POSition -2")
    device.write("CH1:OFFSet 0")
    device.write("CH1:BANdwidth 250E6")

    # Turn cursor display off, set up measurements, and trigger mode
    device.write("CURSor:FUNCtion OFF")
    device.write("MEASUrement:DELETEALL")
    device.write("MEASUrement:MEAS1:SOUrce1 CH1;STATE 1;TYPE PK2Pk")
    device.write("MEASUrement:MEAS2:SOUrce1 CH1;STATE 1;TYPE RMS")

    # Check adequate sample rate
    device.write("HORizontal:MODe AUTO")
    device.write("HORizontal:SAMPLERate:ANALYZemode:MINimum:OVERRide OFF")
    device.write("HORizontal:SAMPLERate:ANALYZemode:MINimum:VALue 3e9")

    device.write("ACQuire:STATE 0")
    device.write("ACQuire:MODe SAMPLE")
    device.write("ACQuire:STOPAfter SEQuence")

    # Wait for scope to finish setting up
    device.commands.opc.query()


# ************** MAIN    
counter = 0  # trigger counter to track data record
with DeviceManager(verbose=True) as device_manager:

    # Open device
    scope:MSO4 = device_manager.add_scope(visaResourceAddr)    # CHANGE FOR YOUR PARTICULAR SCOPE USING Intellisense!
    print(scope.idn_string)

    # Set up scope capture for specific event(s)
    set_up_scope(scope)
    
    # Generate a filename based on the current date & time
    dt = datetime.datetime.now()
    fileName = dt.strftime("%Y%m%d.txt")

    # Trigger scope and start recording
    scope.write("ACQuire:STATE 1")

    # Create save directory if it doesn't exist
    os.makedirs(savePath, exist_ok=True)
    print(f"Data and images will be saved to: {savePath}")
    print("-" * 30)
    # Create data file with header
    with open(os.path.join(savePath , fileName), "w") as datafile:
        datafile.write("Count, Time, Vpk2pk, Vrms\n")
        datafile.close()

        # Trigger Capture Loop
        while (True):

            # Check if user keyboard press
            if keyboard.is_pressed('q'):
                print("Loop terminated by user.")
                break

            Status = scope.query('ACQuire:STATE?')

            if Status == '0' :  
                # Scope triggered
                scope.write("TRIGger:A:LEVel:CH1 4.0")
                print ("triggered")
                counter += 1

                # Get time
                dt = datetime.datetime.now()

                # Get measured data and display for user
                Vp2p = float(scope.query("MEASUREMENT:MEAS1:VALue?"))
                Vrms = float(scope.query("MEASUREMENT:MEAS2:VALue?"))

                print(f"counter: {counter} Vpk2pk: {Vp2p:.3f}, Vrms: {Vrms:.3f}")

                # Append measured data to data file
                with open(os.path.join(savePath , fileName), "a") as datafile:
                    datafile.write(f"{counter:4.0f}, {dt.hour:02d}:{dt.minute:02d}:{dt.second:02d}, {Vp2p:.3f}, {Vrms:.3f}\n")
                    datafile.close()

                # Grab screenshot and save to file
                scope.write('SAVE:IMAGe:FILEFormat PNG')
                scope.write('HARDCOPY:INKSAVER OFF')
                scope.write('HARDCOPY:PORT ETHERNET')
                scope.write('HARDCopy START')
                raw_data = scope.read_raw()
                imgfileName = dt.strftime("%Y%m%d_%H.%M.%S.png")
                imgfilePath = os.path.join(savePath , imgfileName)
                fid = open(imgfilePath, 'wb')
                fid.write(raw_data)
                fid.close()
                # CAUTION- This routine tested on DPO4 series only.  Not tested on newer scopes.
                # HARDCOPY may be depricated on newer scopes.
                # SHOULD WORK BUT DOESN'T.. . is scope.save_screenshot("example.png")

                # Allow time for save before allowing re-triggering, single-mode
                time.sleep(0.5)
                scope.write("ACQuire:STATE 1")

                # Allow time for scope to set up for trigger 
                time.sleep(0.5)

            else:   # Still waiting for a trigger
                # Notify user of status and/or allow user input for other functions
                print ("not triggered")
                scope.write("TRIGger:A:LEVel:CH1 2.0")
                time.sleep(0.5)
                               