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
source_power_meter_address = 30 
load_power_meter_address = 2
dimming_power_meter_address = 10
eload_address = 8
scope_address = "10.125.11.10"
dc_source_address = '4'

"""USER INPUT"""
vin_list = [90,110,115,230,277]
# vin_list = [110]
iout = 0.2
led_list = [48, 36, 24]
# vin_list = convert_argv_to_int_list(sys.argv[1])

test = input("Enter test name: ")
waveforms_folder = f'C:/Users/ccayno/Desktop/DER/DER-727/Test Data/{test}'
path = path_maker(f'{waveforms_folder}')

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
    sleep(3)
    for i in range(1,9):
        eload.channel[i].turn_off()
        eload.channel[i].short_off()
    sleep(1)

def scope_settings():
    global condition
    
    scope.channel_settings(state='OFF', channel=1, scale=100, position=-4, label="Output Voltage", color='LIGHT_BLUE', rel_x_position=20, bandwidth=20, coupling='DC', offset=0)
    scope.channel_settings(state='ON', channel=2, scale=2, position=-4, label="Vbpp", color='YELLOW', rel_x_position=40, bandwidth=20, coupling='DCLimit', offset=0)
    scope.channel_settings(state='ON', channel=3, scale=3, position=-4, label="Vbias_bulk", color='LIGHT_BLUE', rel_x_position=60, bandwidth=20, coupling='DCLimit', offset=0)
    scope.channel_settings(state='ON', channel=4, scale=10, position=-4, label="Output Voltage", color='PINK', rel_x_position=80, bandwidth=20, coupling='DCLimit', offset=0)
    
    # scope.measure(1, "MAX,RMS,FREQ")
    scope.measure(2, "MAX,RMS")
    scope.measure(3, "MAX,RMS")
    scope.measure(4, "MAX,RMS")

    scope.time_position(10)
    
    scope.record_length(50E6)
    
    scope.time_scale(100E-3)

    # scope.remove_zoom()
    # scope.add_zoom(rel_pos=21.727, rel_scale=1)
    
    trigger_channel = 2
    trigger_level = 1
    trigger_edge = 'POS'
    scope.edge_trigger(trigger_channel, trigger_level, trigger_edge)

    scope.trigger_mode('AUTO')

    scope.stop()

def browning(start, end, slew, frequency):

    if start > end:
        print(f"brownout: {start} -> {end} Vac")
        for voltage in range(start, end+1, -slew):
            ac.voltage = voltage
            ac.frequency = frequency
            ac.turn_on()
            sleep(1)
    if start < end:
        print(f"brownin: {start} -> {end} Vac")
        for voltage in range(start, end+1, slew):
            ac.voltage = voltage
            ac.frequency = frequency
            ac.turn_on()
            sleep(1)

def screenshot(filename, path):
    global waveform_counter

    scope.get_screenshot(filename, path)
    print(filename)
    waveform_counter += 1

def operation():
    global waveform_counter

    scope_settings()

    for led in led_list:

        input(f"Change LED load to {led}V")

        for vin in vin_list:

            
            ac.voltage = vin
            ac.turn_on()

            sleep(3)

            scope.run_single()
            sleep(3)


            # input("Capture waveform?")

            filename = f'{vin}Vac, {ac.frequency}Hz, {led}V.png'
            
            screenshot(filename, path)

            discharge_output()

        discharge_output()
        
        
            
            
            # capturing_condition = input("Press ENTER to continue capture waveform. Press anything else to stop capturing waveforms. ")
            # i = 1
            # while capturing_condition == '':
            #     filename = f'{vin}Vac, {ac.frequency}Hz, ({i}).png'
            #     screenshot(filename, path)
            #     i += 1
            #     capturing_condition = input("Press ENTER to continue capture waveform. Press anything else to stop capturing waveforms. ")



def main():
    global waveform_counter
    discharge_output()
    operation()
        
if __name__ == "__main__":
    headers(test)
    main()
    footers(waveform_counter)