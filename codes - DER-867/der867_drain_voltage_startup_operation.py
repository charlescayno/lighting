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
vin_list = [120,230,265,277]
# vin_list = convert_argv_to_int_list(sys.argv[1])

test = "Drain Voltage Starup Operation (50VOR)"
waveforms_folder = f'C:/Users/ccayno/Desktop/DER/DER-867/Final Test Data/{test}'

"""DO NOT EDIT BELOW THIS LINE"""
###################################################################################

"""EQUIPMENT INITIALIZE"""
ac = ACSource(ac_source_address)
pms = PowerMeter(source_power_meter_address)
pml = PowerMeter(load_power_meter_address)
eload = ElectronicLoad(eload_address)
scope = Oscilloscope(scope_address)

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
    scope.channel_settings(state='OFF', channel=1, scale=100, position=0, label="Input Voltage", color='YELLOW', rel_x_position=40, bandwidth=20, coupling='AC', offset=0)
    scope.channel_settings(state='ON', channel=2, scale=100, position=-4, label="Drain Voltage", color='YELLOW', rel_x_position=50, bandwidth=500, coupling='DCLimit', offset=0)
    scope.channel_settings(state='OFF', channel=3, scale=2, position=-4, label="Q1 Regulator", color='GREEN', rel_x_position=60, bandwidth=20, coupling='DCLimit', offset=0)
    scope.channel_settings(state='OFF', channel=4, scale=1, position=0, label="RelayOFF Pulse", color='PINK', rel_x_position=80, bandwidth=20, coupling='DCLimit', offset=0)
    
    # scope.measure(1, "MAX")
    scope.measure(2, "MAX")
    # scope.measure(3, "MAX,MIN")
    # scope.measure(4, "MAX,RMS,FREQ")

    scope.time_position(10)
    scope.time_scale(20E-3)
    scope.record_length(50E6)
    
    # scope.add_zoom(rel_pos=50, rel_scale=0.01)
    scope.remove_zoom()
    
    trigger_channel = 2
    trigger_level = 100
    trigger_edge = 'POS'
    scope.edge_trigger(trigger_channel, trigger_level, trigger_edge)

    scope.stop()

    sleep(2)
    print()

def operation():
    global waveform_counter
    
    scope_settings()
    
    for vin in vin_list:

        scope.run_single()
        sleep(3)

        ac.voltage = vin
        ac.turn_on()

        # eload.channel[1].cc = 0.06
        # eload.channel[1].turn_on()

        sleep(5)
        
        filename = f'{vin}Vac, {ac.frequency}Hz.png'
        path = path_maker(f'{waveforms_folder}')
        scope.get_screenshot(filename, path)
        print(filename)
        waveform_counter += 1

        discharge_output()


def main():
    global waveform_counter
    discharge_output()
    operation()
        
if __name__ == "__main__":
    headers(test)
    main()
    footers(waveform_counter)