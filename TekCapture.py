# Connect to scope to set up, trigger, and save image.

# Select the PyVISA-py backend
import os
import pyvisa 
rm = pyvisa.ResourceManager('@py')

# List available resources
rm.list_resources()
os.environ["TM_OPTIONS"] = "STANDALONE"

from tm_devices import DeviceManager
from tm_devices.drivers import DPO4K

# # import time
# # import numpy

# Modify the following lines to configure this script for your instrument
#==============================================
visaResourceAddr = '10.101.100.254'
#visaResourceAddr = 'TCPIP::10.101.100.236::INSTR'
#fileSaveLocation = 'C:\Users\Calvert.Wong\OneDrive - qsc.com\Desktop\'
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

    scope.write("MEASUrement:MEAS1:SOUrce1 CH1;STATE 1;TYPE PK2Pk")
    scope.write("MEASUrement:MEAS2:SOUrce1 CH1;STATE 1;TYPE RMS")

    scope.write("ACQuire:STATE 0")
    scope.write("ACQuire:MODe SAMPLE")
    scope.write("ACQuire:STOPAfter SEQuence")

    scope.commands.opc.query()
    
    # Trigger


    # Measure
    # voltage = scope.query("MEASUrement:MEAS1:SOUrce1 CH1;STATE 1;TYPE PK2Pk")
    # ch1pk2pk = float(scope.commands.measurement.meas[1].results.allacqs.mean.query())
    # print(f'Channel 1 pk2pk: {ch1pk2pk}')

# scope.commands.measurement.addmeas.write('PK2Pk')

# scope.commands.trigger.a.type.write("EDGE")

# scope.commands.opc.query()

# scope.write("SAVe:IMAGe:FILEFormat PNG")
# scope.write("SAVe:IMAGe:INKSaver OFF")
# scope.write("HARDCopy STARt")
# scope.query("*OPC?")  # Wait for the operation to complete

#imgData = scope.read_raw()

# # Generate a filename based on the current Date & Time
# dt = datetime.now()
# fileName = dt.strftime("%Y%m%d_%H%M%S.png")

# imgFile = open(fileSaveLocation + fileName, "wb")
# imgFile.write(imgData)
# imgFile.close()

# scope.close()
# rm.close()

