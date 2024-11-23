# Connect to scope to set up, trigger, and save image.

import pyvisa_py as visa
from tm_devices import DeviceManager
from tm_devices.drivers import MSO5B
with DeviceManager(verbose=False) as device_manager:
    scope :MSO5B= device_manager.add_scope("10.101.100.236")
# Modify the following lines to configure this script for your instrument
#==============================================
#visaResourceAddr = 'TCPIP::10.101.100.236::INSTR'
#fileSaveLocation = 'C:\Users\Calvert.Wong\OneDrive - qsc.com\Desktop\'
#==============================================

# rm = visa.ResourceManager()
scope.query("*IDN?")
# scope = rm.open_resource(visaResourceAddr)

# print(scope.query('*IDN?'))

scope.commands.display.select.source.write("CH2")

scope.commands.trigger.a.type.write("EDGE")


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

