# For MSO58 Series Scope
# Connect to scope to set up, trigger, and save image.

import time
import datetime
import os
import pyvisa
from tm_devices import DeviceManager
from tm_devices.drivers import MSO5B   # CHANGE FOR YOUR PARTICULAR SCOPE

# List available resources
rm = pyvisa.ResourceManager('@py')
print()
equipment = rm.list_resources()
for i in range(len(equipment)):
    print(equipment[i])
print()

# Configure visaResourceAddr, e.g., 'TCPIP::10.101.100.236::INSTR',  '10.101.100.236', '10.101.100.254', '10.101.100.176'
visaResourceAddr = '10.101.100.151'   # CHANGE FOR YOUR PARTICULAR SCOPE!
SAVE_PATH = r"C:\Users\Calvert.Wong\OneDrive - qsc.com\Desktop\ScopeData" # Ensure this directory exists or create it

def set_up_scope(scope_device):
    """
    Configures the oscilloscope settings for measurement.

    Args:
        scope_device: An instance of the MSO5B oscilloscope device.
    """
    print("Setting up oscilloscope...")
    scope_device.write("SELect:CH1 ON")
    scope_device.write("SELect:CH2 OFF; CH3 OFF; CH4 OFF; CH5 OFF; CH6 OFF; CH7 OFF; CH8 OFF")
    scope_device.write('CH1:LABel:NAMe \"CH1 Vout\"')
    scope_device.write("CH1:COUPling DC")
    scope_device.write("CH1:PRObe:GAIN 0.1")
    scope_device.write("CH1:TERmination MEG")
    scope_device.write("CH1:SCALe 0.5")
    scope_device.write("CH1:INVert OFF")
    scope_device.write("CH1:POSition -2")
    scope_device.write("CH1:OFFSet 0")
    scope_device.write("CH1:BANdwidth 250E6")
    scope_device.write("HORizontal:SCAle 200E-6")
    scope_device.write("HORizontal:POSition 50")
    scope_device.write("TRIGger:A:EDGE:SOUrce CH1")
    scope_device.write("TRIGger:A:EDGE:COUPling DC")
    scope_device.write("TRIGger:A:EDGE:SLOpe RISE")
    scope_device.write("TRIGger:A:LEVel:CH1 2.0")  # TRIGGER:A SETLEVEL may be easier for a mid-level trigger
    scope_device.write("TRIGger:A:MODe NORMal")
    scope_device.write("TRIGger:A:TYPe EDGE")
    scope_device.write("HORizontal:MODe AUTO")
    scope_device.write("HORizontal:SAMPLERate:ANALYZemode:MINimum:VALue 250e6")   # set MIN sample rate
    scope_device.write("HORizontal:SAMPLERate:ANALYZemode:MINimum:OVERRide ON")   # but allow horizontal scale override
    scope_device.write("HORizontal:MODe SCALE 400E-6")
    scope_device.write("DISplay:WAVEView1:CH1:VERTical:SCAle 0.5")
    scope_device.write("DISplay:WAVEView1:CURSor:CURSOR1:STATE 0")
    scope_device.write("DISplay:WAVEView1:INTENSITy:GRATicule 100")
    scope_device.write('MEASUrement:DELETE "MEAS1"')
    scope_device.write('MEASUrement:DELETE "MEAS2"')
    scope_device.write("MEASUrement:MEAS1:SOUrce CH1; TYPE PK2Pk")
    scope_device.write("MEASUrement:MEAS2:SOUrce CH1; TYPE RMS")
    scope_device.write("ACQuire:STATE 0")
    scope_device.write("ACQuire:MODe SAMPLE")
    scope_device.write("ACQuire:STOPAfter SEQuence")

    # List of commands that don't work for MSO5 Series
    # scope_device.write("CURSor:FUNCtion OFF")    
    # scope_device.write("MEASUrement:DELETEALL") 
    # scope_device.write("MEASUrement:MEAS1:STATE OFF")

    # Wait for scope to finish setting up
    scope_device.commands.opc.query()

def save_screen_png(scope_device):


# ************** MAIN    
counter = 0  # trigger counter to track data record
with DeviceManager(verbose=True) as device_manager:
    
    scope:MSO5B = device_manager.add_scope(visaResourceAddr)  # CHANGE FOR YOUR PARTICULAR SCOPE USING Intellisense!
    print(scope.idn_string)

    # Set up scope capture subroutine for specific event(s)
    set_up_scope(scope)
    
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

            # get time stamp and append measurement data
            dt = datetime.datetime.now()
            Vp2p = float(scope.query("MEASUREMENT:MEAS1:VALue?"))
            Vrms = float(scope.query("MEASUREMENT:MEAS2:VALue?"))
            print(f"counter: {counter} Vpk2pk: {Vp2p:.3f}, Vrms: {Vrms:.3f}")
            with open(os.path.join(SAVE_PATH, fileName), "a") as datafile:
                datafile.write(f"{counter:4.0f}, {dt.hour:02d}.{dt.minute:02d}.{dt.second:02d}, {Vp2p:.3f}, {Vrms:.3f}\n")
                datafile.close()

            # # save jpeg
            #     # Create image filename 
            #     imagefilename = os.path.join(SAVE_PATH, 'myimage.png')
            #     print('imagefile = ', imagefilename)
            #     # Save image to instrument's local disk, flash drive, or TekDrive
            #     scope.write('SAVE:IMAGe \"C:/Temp.png\"')
            #     # Wait for instrument to finish writing image to disk
            #     time.sleep(5)
            #     scope.query('*OPC?')
            #     # Generate a filename based on the current Date & Time
            #     # dt = datetime.now()
            #     # fileName = dt.strftime("%YY%mm%dd_%HH%MM%SS.png")
            #     # Read image file from instrument
            #     scope.write('FILESystem:READfile \"C:/Temp.png\"')
            #     # scope.chunk_size = 40960
            #     # image_data = scope.read_raw(640*480)
            #     image_data = scope.read_raw()
            #     # Save image data to local disk
            #     file = open(imagefilename, "wb")
            #     file.write(image_data)
            #     file.close()
            #     # clear output buffers
            #     scope.device_clear()
            #     scope.close

            # ready next trigger
            scope.write("ACQuire:STATE ON")
            time.sleep(2)  # wait before checking again

        else:   # Still waiting for a trigger
            # print ("not triggered")
            scope.write("TRIGger:A:LEVel:CH1 2.0")
            scope.write("ACQuire:STATE ON")
            time.sleep(2)  # wait before checking again

rm.close()
