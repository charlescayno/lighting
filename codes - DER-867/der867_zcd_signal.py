
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
vin_list = [90,220,277]
iout_list = [0.06, 0.005]
vout = 3.5

test = "ZCD signal (MOSFET Config)"
waveforms_folder = f'C:/Users/ccayno/Desktop/DER-867 LNK-TNZ ONOFF Switch/Test Data/{test}'

"""DO NOT EDIT BELOW THIS LINE"""
##################################################################################

"""EQUIPMENT INITIALIZE"""
ac = ACSource(ac_source_address)
pms = PowerMeter(source_power_meter_address)
pml = PowerMeter(load_power_meter_address)
eload = ElectronicLoad(eload_address)
scope = Oscilloscope(scope_address)


"""GENERIC FUNCTIONS"""

def discharge_output():
    ac.turn_off()
    for i in range(1,9):
        eload.channel[i].cc = 1
        eload.channel[i].turn_on()
        eload.channel[i].short_on()
    sleep(2)
    for i in range(1,9):
        eload.channel[i].turn_off()
        eload.channel[i].short_off()
    sleep(1)

# def scope_settings():
#     scope.position_scale(time_position = 0.5, time_scale = 0.01)

def operation():

    global waveform_counter

    # scope_settings()
    
    for iout in iout_list:
        
        for vin in vin_list:

                ac.voltage = vin
                ac.turn_on()

                eload.channel[1].cc = iout
                eload.channel[1].turn_on()

                sleep(3)
        
                scope.edge_trigger(4, 0, 'NEG')
                scope.run_single()
                
                sleep(6)

                filename = f'{vin}Vac, {iout}A, {test}.png'
                path = path_maker(f'{waveforms_folder}')
                scope.get_screenshot(filename, path)
                print(filename)
                waveform_counter += 1
                
                discharge_output()
    
def main():
    global waveform_counter
    operation()
        
if __name__ == "__main__":
    headers(test)
    discharge_output()
    main()
    discharge_output()
    footers(waveform_counter)