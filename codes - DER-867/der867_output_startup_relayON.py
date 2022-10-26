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
vin_list = [90,120,220,265,277]

trigger_channel = 4
trigger_level = 1
trigger_edge = 'EITH'


test = "Output Waveforms, Startup, Relay ON"
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

def operation():
    global waveform_counter
    
    scope.channel_label(2, label="Q1 regulator output (3V LDO input)", rel_x_position=25)
    scope.time_position(20)
    scope.stop()

    input("Change input voltage probe points.")

    input("On relay. Press enter to continue...")
    
    for vin in vin_list:

        scope.edge_trigger(trigger_channel, trigger_level, trigger_edge)
        scope.run_single()

        sleep(3)

        ac.voltage = vin
        ac.turn_on()

        sleep(3)
        ac.turn_off()

        input("Capture waveform?")

        filename = f'{test}, {vin}Vac.png'
        path = path_maker(f'{waveforms_folder}')
        scope.get_screenshot(filename, path)
        print(filename)
        waveform_counter += 1
          
def main():
    global waveform_counter
    operation()
        
if __name__ == "__main__":
    headers(test)
    main()
    footers(waveform_counter)