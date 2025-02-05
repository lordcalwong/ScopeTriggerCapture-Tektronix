# For MSO58 Series Scope
# Connect to scope to set up, trigger, and save image.

# Use time, date, and file utility packages
import time
import os

# Select the PyVISA-py backend
# from gpib_ctypes import make_default_gpib
# make_default_gpib()
import pyvisa   # control of instruments over wide range of interfaces
rm = pyvisa.ResourceManager('@py')

# Use Python device management package from Tektronix
from tm_devices import DeviceManager
from tm_devices.drivers import MSO5B                        # CHANGE FOR YOUR PARTICULAR SCOPE USING Intellisense!
# from tm_devices.helpers import PYVISA_PY_BACKEND, SYSTEM_DEFAULT_VISA_BACKEND

# List available resources
rm.list_resources()
equipment = rm.list_resources()
for i in range(len(equipment)):
    print(equipment[i])
os.environ["TM_OPTIONS"] = "STANDALONE"
print()

# Modify the following lines to configure this script 
# for your needs or particular instrument
#==============================================
# visaResourceAddr = '10.101.100.254'   #DPO4034
# visaResourceAddr = '10.101.100.236'   #MSO58                # CHANGE FOR YOUR PARTICULAR SCOPE!
visaResourceAddr = '10.101.100.104'   #MSO58                # CHANGE FOR YOUR PARTICULAR SCOPE!
#visaResourceAddr = 'TCPIP::10.101.100.236::INSTR'
savePath = "C:\\Users\\Calvert.Wong\\OneDrive - qsc.com\\Desktop\\"
#==============================================

with DeviceManager(verbose=True) as device_manager:

    scope:MSO5B = device_manager.add_scope(visaResourceAddr)  # CHANGE FOR YOUR PARTICULAR SCOPE USING Intellisense!
    print(scope.idn_string)

    # This routine works for the MSO5 Series
    # Create image filename 
    imagefilename = os.path.join(savePath , 'myimage.png')
    print('imagefile = ', imagefilename)
    # Save image to instrument's local disk, flash drive, or TekDrive
    scope.write('SAVE:IMAGe \"C:/Temp.png\"')
    # Wait for instrument to finish writing image to disk
    time.sleep(5)
    scope.query('*OPC?')

    scope.write('FILESystem:READfile \"C:/Temp.png\"')
    
    # scope.chunk_size = 40960
    # image_data = scope.read_raw(640*480)
    image_data = scope.read_raw()
    
    # Save image data to local disk
    file = open(imagefilename, 'wb')
    file.write(image_data)
    file.close()

    scope.device_clear()
    scope.close

rm.close


