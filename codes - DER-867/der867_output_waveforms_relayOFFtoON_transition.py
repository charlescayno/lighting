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
vin_list = [90,120,230,265]

test = "Output Waveforms, Relay OFF to ON transition"
waveforms_folder = f'C:/Users/ccayno/Desktop/DER/DER-867/Final Test Data/{test}'

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
    scope.channel_settings(state='ON', channel=1, scale=200, position=-3, label="Input Voltage", color='YELLOW', rel_x_position=10, bandwidth=20, coupling='AC', offset=0)
    scope.channel_settings(state='ON', channel=2, scale=2, position=-1, label="3V LDO Input", color='ORANGE', rel_x_position=40, bandwidth=20, coupling='DCLimit', offset=0)
    scope.channel_settings(state='ON', channel=3, scale=1, position=-1, label="3V LDO Output", color='GREEN', rel_x_position=65, bandwidth=20, coupling='DCLimit', offset=0)
    scope.channel_settings(state='ON', channel=4, scale=2, position=-2, label="Relay ON pulse", color='BLUE', rel_x_position=90, bandwidth=20, coupling='DCLimit', offset=0)
    
    scope.edge_trigger(4, 1,'NEG')
    scope.time_position(10)
    scope.time_scale(30E-3)
    scope.record_length(50E6)

    scope.measure(1, "MAX,RMS,FREQ")
    scope.measure(2, "MAX,MIN,MEAN")
    scope.measure(3, "MAX,MIN,MEAN")
    scope.measure(4, "MAX")

    scope.stop()

def operation():
    global waveform_counter
    
    scope_settings()

    ac.voltage = 90
    ac.turn_on()
    sleep(2)
    
    for vin in vin_list:

        ac.voltage = vin
        ac.turn_on()
        
        sleep(3)

        scope.run_single()

        input("ON relay. Press enter to continue...")

        filename = f'{vin}Vac, {ac.frequency}Hz.png'
        path = path_maker(f'{waveforms_folder}')
        scope.get_screenshot(filename, path)
        print(filename)
        waveform_counter += 1

        discharge_output()
          
def main():
    global waveform_counter
    discharge_output()
    input(">> Put input voltage to L_OUT.")
    input(">> Attached probe to relay ON.")
    operation()
        
if __name__ == "__main__":
    headers(test)
    main()
    footers(waveform_counter)