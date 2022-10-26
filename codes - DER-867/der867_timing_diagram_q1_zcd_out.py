# NOT INCLUDED IN THE FINAL TEST DATA

"""IMPORT DEPENDENCIES"""
from calendar import c
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
vin_list = [120,230]
# vin_list = [230]
test = "Timing Diagram ZCD_OUT and Q1"
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

def scope_settings():
    global condition
    scope.channel_settings(state='ON', channel=1, scale=200, position=0, label="Input Voltage", color='YELLOW', rel_x_position=10, bandwidth=20, coupling='AC', offset=0)
    scope.channel_settings(state='ON', channel=2, scale=2, position=3, label="ZCD_IN", color='ORANGE', rel_x_position=30, bandwidth=20, coupling='DCLimit', offset=0)
    scope.channel_settings(state='ON', channel=3, scale=2, position=-4, label="Q1", color='GREEN', rel_x_position=50, bandwidth=20, coupling='DCLimit', offset=0)
    scope.channel_settings(state='ON', channel=4, scale=2, position=0, label="ZCD_OUT", color='LIGHT_BLUE', rel_x_position=70, bandwidth=20, coupling='DCLimit', offset=0)
    
    scope.edge_trigger(2, 1,'POS')
    scope.time_position(50)
    scope.time_scale(100E-3)
    scope.record_length(50E6)

    scope.measure(1, "RMS")
    scope.measure(2, "MAX,MIN")
    scope.measure(3, "MAX,MIN")
    scope.measure(4, "MAX,MIN")


    

    scope.stop()

    sleep(2)

def operation():
    global waveform_counter
    global load
    
    scope_settings()



    
    
    for vin in vin_list:

        ac.voltage = vin
        ac.turn_on()
        sleep(3)
        input("Turn ON switch")
        scope.run_single()
        sleep(3)

        input("Press ENTER to capture waveform...")

        
        filename = f'{test}, {vin}Vac, {ac.frequency}Hz, {load}.png'
        path = path_maker(f'{waveforms_folder}')
        scope.get_screenshot(filename, path)
        print(filename)
        waveform_counter += 1

        scope.run_single()
        sleep(3)
          
def main():
    global waveform_counter
    global load
    load = input("Load: ")
    operation()
    
        
if __name__ == "__main__":
    headers(test)
    main()
    footers(waveform_counter)