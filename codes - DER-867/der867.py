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
scope_address = "10.125.11.20"

"""USER INPUT"""
vin_list = [120, 230]
# vin_list = [230] # led load
# vin_list = convert_argv_to_int_list(sys.argv[1])

test = input("Enter test name: ")
waveforms_folder = f'C:/Users/ccayno/Desktop/DER/DER-867/Final Test Data/{test}'


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
    
    scope.channel_settings(state='ON', channel=1, scale=100, position=-1, label="Input Voltage", color='YELLOW', rel_x_position=20, bandwidth=20, coupling='AC', offset=0)
    scope.channel_settings(state='ON', channel=2, scale=2, position=3, label="ZCD_IN", color='ORANGE', rel_x_position=40, bandwidth=20, coupling='DCLimit', offset=0)
    if choice == 1: scope.channel_settings(state='ON', channel=3, scale=1, position=-1, label="RelayON Pulse", color='LIGHT_BLUE', rel_x_position=60, bandwidth=20, coupling='DCLimit', offset=0)
    else: scope.channel_settings(state='ON', channel=3, scale=1, position=-1, label="RelayOFF Pulse", color='GREEN', rel_x_position=60, bandwidth=20, coupling='DCLimit', offset=0)
    scope.channel_settings(state='ON', channel=4, scale=10, position=-1, label="Input Current", color='PINK', rel_x_position=80, bandwidth=20, coupling='DCLimit', offset=0)
    # scope.channel_settings(state='ON', channel=4, scale=5, position=-1, label="Input Current", color='PINK', rel_x_position=80, bandwidth=20, coupling='DCLimit', offset=0) # led load
    
    scope.measure(1, "MAX,RMS,FREQ")
    scope.measure(2, "MAX,MIN")
    scope.measure(3, "MAX,MIN")
    scope.measure(4, "MAX,MIN")

    scope.time_position(20)
    
    scope.record_length(50E6)
    
    scope.time_scale(1)

    # scope.remove_zoom()
    # scope.add_zoom(rel_pos=21.727, rel_scale=1)
    
    trigger_channel = 2
    trigger_level = 1
    trigger_edge = 'POS'
    scope.edge_trigger(trigger_channel, trigger_level, trigger_edge)

    scope.stop()

def operation():
    global waveform_counter
    global choice
    
    
    choice = int(input("Press 1 for RelayON_Pulse, Press 2 for RelayOFF_Pulse: "))
    scope_settings()
    
    for vin in vin_list:

        scope.run_single()
        sleep(3)

        ac.voltage = vin
        ac.turn_on()

        sleep(3)
        print("Turn switch on")
        sleep(3)
        print("Turn switch off")

        sleep(4)
        ac.turn_off()

        input("Capture waveform?")

        if choice == 1: j = "RelayON_pulse"
        else: j = "RelayOFF_pulse"

        filename = f'{j} - {vin}Vac, {ac.frequency}Hz.png'
        path = path_maker(f'{waveforms_folder}')
        scope.get_screenshot(filename, path)
        print(filename)
        waveform_counter += 1
        
        
        capturing_condition = input("Press ENTER to continue capture waveform. Press anything else to stop capturing waveforms. ")
        i = 1
        while capturing_condition == '':
            filename = f'{j} - {vin}Vac, {ac.frequency}Hz, ({i}).png'
            path = path_maker(f'{waveforms_folder}')
            scope.get_screenshot(filename, path)
            print(filename)
            waveform_counter += 1
            i += 1
            capturing_condition = input("Press ENTER to continue capture waveform. Press anything else to stop capturing waveforms. ")





def main():
    global waveform_counter
    discharge_output()
    operation()
        
if __name__ == "__main__":
    headers(test)
    main()
    footers(waveform_counter)