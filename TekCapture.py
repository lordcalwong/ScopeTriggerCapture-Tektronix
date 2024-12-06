# Connect to scope to set up, trigger, and save image.

# Select the PyVISA-py backend
import os
import pyvisa 
rm = pyvisa.ResourceManager('@py')
counter = 0

# Use Python device management package from Tektronix
from tm_devices import DeviceManager
from tm_devices.drivers import DPO4K

# List available resources
rm.list_resources()
os.environ["TM_OPTIONS"] = "STANDALONE"

# Some file utility packages
import time
import datetime
import os

# Modify the following lines to configure this script 
# for your needs or particular instrument
#==============================================
visaResourceAddr = '10.101.100.254'
#visaResourceAddr = 'TCPIP::10.101.100.236::INSTR'
savePath = "C:\\Users\\Calvert.Wong\\OneDrive - qsc.com\\Desktop\\"
#==============================================



with DeviceManager(verbose=True) as device_manager:
    scope:DPO4K = device_manager.add_scope(visaResourceAddr)
    print()
    print(scope.idn_string)

    scope.write("SELect:CH1 ON")
    scope.write("SELect:CH2 OFF")
    scope.write("SELect:CH3 OFF")
    scope.write("SELect:CH4 OFF")
    # Note- "print(scope.commands.select.ch[1],'ON')" doesn't work but should

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

    # This capability and nomenclature may vary by scope MFR for 
    # cursor displayand making measurements
    scope.write("CURSor:FUNCtion OFF")
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
        # slow script down
        time.sleep(5)
        Status = scope.query('ACQuire:STATE?')
        if Status == '0' :  # Scope triggered
            print ("triggered")
            counter += 1
            # get time
            dt = datetime.datetime.now()
            # grab measured data and display for user
            Vp2p = float(scope.query("MEASUREMENT:MEAS1:VALue?"))
            Vrms = float(scope.query("MEASUREMENT:MEAS2:VALue?"))
            print(f"Vpk2pk: {Vp2p:.3f}, Vrms: {Vrms:.3f}")
            # append data to data file
            with open(os.path.join(savePath , fileName), "a") as datafile:
                datafile.write(f"{counter}, {dt.hour}.{dt.minute}.{dt.second}, {Vp2p:.3f}, {Vrms:.3f}\n")
                datafile.close()
            # Wait 5 sec before allowing re-triggering, single-mode
            time.sleep(5)
            scope.write("ACQuire:STATE 1")
        else:   # Still waiting for a trigger
            # notify user of status and/or allow user input for other functions
            print ("not triggered")
            time.sleep(5)

scope.close()
rm.close()

