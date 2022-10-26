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

from datetime import datetime
now = datetime.now()
date = now.strftime('%Y_%m_%d')	

##################################################################################

"""COMMS"""
ac_source_address = 5
source_power_meter_address = 1 
load_power_meter_address = 2
eload_address = 8
scope_address = "10.125.10.184"

"""USER INPUT"""
vin_list = [120,230]
led_list = convert_argv_to_int_list(sys.argv[1])
vin_list = convert_argv_to_int_list(sys.argv[2])

test = input("Enter test name: ")
waveforms_folder = f'C:/Users/ccayno/Desktop/DER-945 Marketing Sample Units/Test Data/{test}'

"""DO NOT EDIT BELOW THIS LINE"""
##################################################################################

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
    scope.channel_settings(state='ON', channel=1, scale=0.3, position=-4, label="Output Current", color='PINK', rel_x_position=20, bandwidth=20, coupling='DCLimit', offset=0)
    scope.measure(1, "MAX,MIN")

    scope.channel_settings(state='ON', channel=2, scale=200, position=-4, label="VDS", color='YELLOW', rel_x_position=40, bandwidth=500, coupling='DCLimit', offset=0)
    scope.measure(2, "MAX,MIN")

    scope.channel_settings(state='ON', channel=3, scale=10, position=-4, label="Output Voltage", color='GREEN', rel_x_position=60, bandwidth=20, coupling='DCLimit', offset=0)
    scope.measure(3, "MAX,MIN")

    scope.channel_settings(state='ON', channel=4, scale=10, position=-4, label="VR", color='LIGHT_BLUE', rel_x_position=80, bandwidth=20, coupling='DCLimit', offset=0)
    scope.measure(4, "MAX,MIN")
       

    scope.time_position(10)
    scope.record_length(50E6)
    scope.time_scale(1)
    # scope.time_scale(200E-3)
    # scope.add_zoom(rel_pos=24.1, rel_scale=1)
    # scope.remove_zoom()


    trigger_channel = 2
    trigger_level = 100
    trigger_edge = 'POS'
    scope.edge_trigger(trigger_channel, trigger_level, trigger_edge)

    scope.stop()

    sleep(2)

def operation():
    global waveform_counter
    
    scope_settings()

    for led in led_list:
    
        for vin in vin_list:

            scope.run_single()
            sleep(3)

            ac.voltage = vin
            ac.turn_on()

            sleep(11)

            ac.turn_off()
            # input("Capture waveform?")

            filename = f'{test}, {led}V, {vin}Vac, {ac.frequency}Hz.png'
            path = path_maker(f'{waveforms_folder}')
            scope.get_screenshot(filename, path)
            print(filename)
            waveform_counter += 1

            discharge_output()
            

            # capturing_condition = input("Press ENTER to continue capture waveform. Press anything else to stop capturing waveforms. ")
            # i = 1
            # while capturing_condition == '':
            #     filename = f'{test}, {vin}Vac, {ac.frequency}Hz, ({i}).png'
            #     path = path_maker(f'{waveforms_folder}')
            #     scope.get_screenshot(filename, path)
            #     print(filename)
            #     waveform_counter += 1
            #     i += 1
            #     capturing_condition = input("Press ENTER to continue capture waveform. Press anything else to stop capturing waveforms. ")





def main():
    global waveform_counter
    operation()
        
if __name__ == "__main__":
    headers(test)
    discharge_output()
    main()
    discharge_output()
    footers(waveform_counter)