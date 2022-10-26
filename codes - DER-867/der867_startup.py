# NOT INCLUDED IN THE FINAL TEST DATA

"""IMPORT DEPENDENCIES"""
from time import time, sleep
import sys
import os
import math
from powi.equipment import ACSource, PowerMeter, ElectronicLoad, Oscilloscope, LEDControl
from powi.equipment import headers, create_folder, footers, waveform_counter, soak, convert_argv_to_int_list, tts, prompt
from filemanager import path_maker, remove_file
import winsound as ws
from playsound import playsound
waveform_counter = 0

##################################################################################

"""COMMS"""
ac_source_address = 5
source_power_meter_address = 1 
load_power_meter_address = 2
eload_address = 8
scope_address = "10.125.10.184"

"""USER INPUT"""
vin_list = [230]
vout = 3.5

test = "Startup Sequence"
waveforms_folder = f'C:/Users/ccayno/Desktop/DER-867 LNK-TNZ ONOFF Switch/Test Data/{test}'

"""DO NOT EDIT BELOW THIS LINE"""
##################################################################################

"""EQUIPMENT INITIALIZE"""
ac = ACSource(ac_source_address)
# pms = PowerMeter(source_power_meter_address)
# pml = PowerMeter(load_power_meter_address)
# eload = ElectronicLoad(eload_address)
scope = Oscilloscope(scope_address)


"""GENERIC FUNCTIONS"""

# def scope_settings():
#     scope.position_scale(time_position = 0.5, time_scale = 0.01)

def operation():

    global waveform_counter
        
    for vin in vin_list:

        scope.run_single()
        scope.edge_trigger(2, 1, 'POS')

        sleep(3)

        ac.voltage = vin
        ac.turn_on()


        sleep(3)

        
        print("ON relay")

        

        a = 'b'
        while a != 'q':
            temp_filename = f"{test}"
            i = 0
            waveform_counter = 0
            print("\n>>> Press ENTER to start capturing waveforms on SCOPE.")
            while '' == input():
                if i == 0: filename = temp_filename + f", {vin}Vac.png"
                else: filename = temp_filename + f", {vin}Vac, ({i}).png"
                print(filename)
                path = path_maker(waveforms_folder)
                i += 1
                scope.get_screenshot(filename, path)
                waveform_counter += 1




        ac.turn_off()
    
def main():
    global waveform_counter
    operation()
        
if __name__ == "__main__":
    headers(test)
    main()
    footers(waveform_counter)