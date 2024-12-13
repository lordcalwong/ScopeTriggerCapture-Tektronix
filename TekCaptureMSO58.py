# For MSO58 Series Scope
# Connect to scope to set up, trigger, and save image.

# Select the PyVISA-py backend
import os       # interact with the operating system and file management
import pyvisa   # control of instruments over wide range of interfaces
rm = pyvisa.ResourceManager('@py')
counter = 0

# Use Python device management package from Tektronix
from tm_devices import DeviceManager
from tm_devices.drivers import MSO5B                        # CHANGE FOR YOUR PARTICULAR SCOPE USING Intellisense!
# from tm_devices.helpers import PYVISA_PY_BACKEND, SYSTEM_DEFAULT_VISA_BACKEND

# List available resources
rm.list_resources()
os.environ["TM_OPTIONS"] = "STANDALONE"

# Use time, date, and file utility packages
import time
import datetime

# Modify the following lines to configure this script 
# for your needs or particular instrument
#==============================================
# visaResourceAddr = '10.101.100.254'   #DPO4034
# visaResourceAddr = '10.101.100.236'   #MSO58                # CHANGE FOR YOUR PARTICULAR SCOPE!
visaResourceAddr = '10.101.101.101'   #MSO58                # CHANGE FOR YOUR PARTICULAR SCOPE!
#visaResourceAddr = 'TCPIP::10.101.100.236::INSTR'
savePath = "C:\\Users\\Calvert.Wong\\OneDrive - qsc.com\\Desktop\\"
#==============================================

with DeviceManager(verbose=True) as device_manager:
    
    scope:MSO5B = device_manager.add_scope(visaResourceAddr)  # CHANGE FOR YOUR PARTICULAR SCOPE USING Intellisense!
    print()
    print(scope.idn_string)

    scope.write("SELect:CH1 ON")
    scope.write("SELect:CH2 OFF")
    scope.write("SELect:CH3 OFF")
    scope.write("SELect:CH4 OFF")
    # Note- "print(scope.commands.select.ch[1],'ON')" doesn't work but should
    scope.write('CH1:LABel:NAMe \"Vout\"')

    scope.write("HORizontal:SCAle 200E-6")
    scope.write("HORizontal:POSition 50")
    
    scope.write("TRIGger:A:EDGE:SOUrce CH1")
    scope.write("TRIGger:A:EDGE:COUPling DC")
    scope.write("TRIGger:A:EDGE:SLOpe RISE")
    scope.write("TRIGger:A:LEVel:CH1 2.0")
    # Use "scope.write("TRIGGER:A SETLEVEL")" instead for mid-level"
    scope.write("TRIGger:A:MODe NORMal")
    scope.write("TRIGger:A:TYPe EDGE")

    scope.write("CH1:COUPling DC")
    scope.write("CH1:PRObe:GAIN 0.1")
    scope.write("CH1:TERmination MEG")
    scope.write("CH1:SCALe 0.5")
    scope.write("CH1:INVert OFF")
    scope.write("CH1:POSition -2")
    scope.write("CH1:OFFSet 0")
    scope.write("CH1:BANdwidth 250E6")

    # Set up measurements.
    # Clear, add measurements and set up acqusition mode.
    # Note- This capability and nomenclature may vary by scope MFR for cursor display and making measurements.
    scope.write("CURSor:FUNCtion OFF")
    scope.write("MEASUrement:DELETEALL")
    scope.write("MEASUrement:MEAS1:SOUrce1 CH1;STATE 1;TYPE PK2Pk")
    scope.write("MEASUrement:MEAS2:SOUrce1 CH1;STATE 1;TYPE RMS")
    scope.write("ACQuire:STATE 0")
    scope.write("ACQuire:MODe SAMPLE")
    scope.write("ACQuire:STOPAfter SEQuence")

    # Wait for scope to finsh
    scope.commands.opc.query()
    
    # Generate a filename based on the current Date & Time
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
        # slow script down for interrupts
        time.sleep(1)
        Status = scope.query('ACQuire:STATE?')
        if Status == '0' :  
            # Scope triggered
            print ("triggered")
            counter += 1

            # get time
            dt = datetime.datetime.now()

            # get measured data and display for user
            Vp2p = float(scope.query("MEASUREMENT:MEAS1:VALue?"))
            Vrms = float(scope.query("MEASUREMENT:MEAS2:VALue?"))
            print(f"counter: {counter} Vpk2pk: {Vp2p:.3f}, Vrms: {Vrms:.3f}")

            # # This routine works but has been deprecated in the newer scopes
            # # Use only on Tek DPO4000/MSO4000 scopes and earlier
            # # grab screenshot and save to file
            # scope.write('SAVE:IMAGe:FILEFormat PNG')
            # scope.write('HARDCOPY:INKSAVER OFF')
            # scope.write('HARDCOPY:PORT ETHERNET')
            # scope.write('HARDCopy START')
            # raw_data = scope.read_raw()
            # fid = open('my_image.png', 'wb')
            # fid.write(raw_data)
            # fid.close()

            # This routine works for the MSO5 Series
            # Create image filename 
            imagefilename = os.path.join(savePath , 'myimage.png')
            print('imagefile = ', imagefilename)
            # Save image to instrument's local disk, flash drive, or TekDrive
            scope.write('SAVE:IMAGe \"C:/Temp.png\"')
            # Wait for instrument to finish writing image to disk
            scope.query('*OPC?')
            # Generate a filename based on the current Date & Time
            # dt = datetime.now()
            # fileName = dt.strftime("%YY%mm%dd_%HH%MM%SS.png")
            # Read image file from instrument
            scope.write('FILESystem:READfile \"C:/Temp.png\"')
            scope.chunk_size = 40960
            image_data = scope.read_raw(640*480)
            # Save image data to local disk
            file = open(imagefilename, "wb")
            file.write(image_data)
            file.close()


            # SHOULD WORK BUT DOESN'T.. .
            # scope.commands.save.waveform.write('CH1,"example.wfm"')
            # scope.save_screenshot("example.png")

            # append measured data to data file
            with open(os.path.join(savePath , fileName), "a") as datafile:
                datafile.write(f"{counter:4.0f}, {dt.hour:02d}.{dt.minute:02d}.{dt.second:02d}, {Vp2p:.3f}, {Vrms:.3f}\n")
                datafile.close()

            # Allow time for save before allowing re-triggering, single-mode
            time.sleep(2)
            scope.write("ACQuire:STATE 1")
            # Allow time for scope to set up for trigger 
            time.sleep(2)
        else:   # Still waiting for a trigger
            # notify user of status and/or allow user input for other functions
            print ("not triggered")
            time.sleep(1)

scope.close()
rm.close()
