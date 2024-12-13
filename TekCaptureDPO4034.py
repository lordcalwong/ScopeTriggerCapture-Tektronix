# For DPO4034 Series Scope
# Connect to scope to set up, trigger, and save image/screenshot (hardcopy) to PC for 4 Series MSO Oscilloscopes

# Select the PyVISA-py backend
import os       # interact with the operating system and file management
import pyvisa   # control of instruments over wide range of interfaces
rm = pyvisa.ResourceManager('@py')
counter = 0

# Use Python device management package from Tektronix
from tm_devices import DeviceManager
from tm_devices.drivers import MSO5                            # CHANGE FOR YOUR PARTICULAR SCOPE USING Intellisense!
# from tm_devices.helpers import PYVISA_PY_BACKEND, SYSTEM_DEFAULT_VISA_BACKEND

# List available resources
rm.list_resources()
os.environ["TM_OPTIONS"] = "STANDALONE"

# Use time, date packages
import time
import datetime

# Use keyboard package
import keyboard

# Modify following section to configure this script for scope or interface
#==============================================
visaResourceAddr = '10.101.100.176'     #DPO4034                # CHANGE FOR YOUR PARTICULAR SCOPE!
# visaResourceAddr = '10.101.100.236'   #MSO58
#visaResourceAddr = 'TCPIP::10.101.100.236::INSTR'
savePath = "C:\\Users\\Calvert.Wong\\OneDrive - qsc.com\\Desktop\\"
#==============================================


with DeviceManager(verbose=True) as device_manager:

    # Open device
    scope:MSO5 = device_manager.add_scope(visaResourceAddr)    # CHANGE FOR YOUR PARTICULAR SCOPE USING Intellisense!
    print(scope.idn_string)

    # Set up scope
    # Set up dispaly channels
    scope.write("SELect:CH1 ON")
    scope.write("SELect:CH2 OFF")
    scope.write("SELect:CH3 OFF")
    scope.write("SELect:CH4 OFF")
    # Note- "print(scope.commands.select.ch[1],'ON')" doesn't work but should
    
    # Set up timebase
    scope.write("HORizontal:SCAle 200E-6")
    scope.write("HORizontal:POSition 50")
    
    # Set up trigger
    scope.write("TRIGger:A:EDGE:SOUrce CH1")
    scope.write("TRIGger:A:EDGE:COUPling DC")
    scope.write("TRIGger:A:EDGE:SLOpe RISE")
    scope.write("TRIGger:A:LEVel:CH1 2.0")
    # Use "scope.write("TRIGGER:A SETLEVEL")" instead for mid-level"
    scope.write("TRIGger:A:MODe NORMal")
    scope.write("TRIGger:A:TYPe EDGE")

    # Set up vertical
    scope.write("CH1:COUPling DC")
    scope.write("CH1:PRObe:GAIN 0.1")
    scope.write("CH1:TERmination MEG")
    scope.write("CH1:SCALe 0.5")
    scope.write("CH1:INVert OFF")
    scope.write("CH1:POSition -2")
    scope.write("CH1:OFFSet 0")
    scope.write("CH1:BANdwidth 250E6")

    # Turn cursor display off, set up measurements, and trigger mode
    scope.write("CURSor:FUNCtion OFF")
    scope.write("MEASUrement:MEAS1:SOUrce1 CH1;STATE 1;TYPE PK2Pk")
    scope.write("MEASUrement:MEAS2:SOUrce1 CH1;STATE 1;TYPE RMS")
    scope.write("ACQuire:STATE 0")
    scope.write("ACQuire:MODe SAMPLE")
    scope.write("ACQuire:STOPAfter SEQuence")

    # Wait for scope to finish setting up
    scope.commands.opc.query()
    
    # Generate a filename based on the current date & time
    dt = datetime.datetime.now()
    fileName = dt.strftime("%Y%m%d.txt")

    # Trigger scope and start recording
    scope.write("ACQuire:STATE 1")

    # Create data file with header
    with open(os.path.join(savePath , fileName), "w") as datafile:
        datafile.write("Count, Time, Vpk2pk, Vrms\n")
        datafile.close()

        # Trigger Capture Loop
        while (True):
            # Slow script down for interrupts
            time.sleep(1)

            Status = scope.query('ACQuire:STATE?')
            if Status == '0' :  
                # Scope triggered
                print ("triggered")
                counter += 1

                # Get time
                dt = datetime.datetime.now()

                # Get measured data and display for user
                Vp2p = float(scope.query("MEASUREMENT:MEAS1:VALue?"))
                Vrms = float(scope.query("MEASUREMENT:MEAS2:VALue?"))
                print(f"counter: {counter} Vpk2pk: {Vp2p:.3f}, Vrms: {Vrms:.3f}")

                # Grab screenshot and save to file
                scope.write('SAVE:IMAGe:FILEFormat PNG')
                scope.write('HARDCOPY:INKSAVER OFF')
                scope.write('HARDCOPY:PORT ETHERNET')
                scope.write('HARDCopy START')
                raw_data = scope.read_raw()
                fid = open('my_image.png', 'wb')
                fid.write(raw_data)
                fid.close()
                # Note- This routine works but has been deprecated in their newer scopes.
                # CAUTION- Use only on Tek DPO4000/MSO4000 scopes and earlier

                # SHOULD WORK BUT DOESN'T.. .
                # scope.save_screenshot("example.png")

                # Append measured data to data file
                with open(os.path.join(savePath , fileName), "a") as datafile:
                    datafile.write(f"{counter:4.0f}, {dt.hour:02d}.{dt.minute:02d}.{dt.second:02d}, {Vp2p:.3f}, {Vrms:.3f}\n")
                    datafile.close()

                # Allow time for save before allowing re-triggering, single-mode
                time.sleep(2)
                scope.write("ACQuire:STATE 1")

                # Allow time for scope to set up for trigger 
                time.sleep(2)

                # Check if user keyboard press
                if keyboard.is_pressed('q'):
                    print("Loop terminated by user.")
                    scope.close()
                    rm.close()
                    break

            else:   # Still waiting for a trigger
                # Notify user of status and/or allow user input for other functions
                print ("not triggered")
                time.sleep(1)
                               



