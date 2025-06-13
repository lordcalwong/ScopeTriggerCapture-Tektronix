# For MSO58 Series Scope
# Connect to scope to set up, trigger, and save image.

import time
import os
import pyvisa
rm = pyvisa.ResourceManager('@py')
from tm_devices import DeviceManager
from tm_devices.drivers import MSO5B  # CHANGE FOR YOUR PARTICULAR SCOPE USING Intellisense!


# List available resources
rm.list_resources()
os.environ["TM_OPTIONS"] = "STANDALONE"


# Configure visaResourceAddr, e.g., 'TCPIP::10.101.100.236::INSTR',  '10.101.100.236', '10.101.100.254', '10.101.100.176'
visaResourceAddr = '10.101.100.151'   # CHANGE FOR YOUR PARTICULAR SCOPE!
savePath = "C:\\Users\\Calvert.Wong\\OneDrive - qsc.com\\Desktop\\"


with DeviceManager(verbose=True) as device_manager:
    
    scope:MSO5B = device_manager.add_scope(visaResourceAddr)  # CHANGE FOR YOUR PARTICULAR SCOPE USING Intellisense!
    print()
    print(scope.idn_string)

    # This routine works for the MSO5 Series
    # Create image filename 
    imagefilename = os.path.join(savePath , 'myimage.png')
    print('imagefile = ', imagefilename)
    # Save image to instrument's local disk, flash drive, or TekDrive
    scope.write('SAVE:IMAGe \"C:/Temp.png\"')
    # Wait for instrument to finish writing image to disk
    time.sleep(5)

    scope.write('FILESystem:READfile \"C:/Temp.png\"')
    
    image_data = scope.read_raw()
 
    # Save image data to local disk
    file = open(imagefilename, "wb")
    file.write(image_data)
    file.close()

