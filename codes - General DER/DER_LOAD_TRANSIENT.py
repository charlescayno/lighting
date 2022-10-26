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
"""COMMS ADDRESS"""
ac_source_address = 5
source_power_meter_address = 2 
load_power_meter_address = 1
eload_address = 8
scope_address = ""

"""USER INPUT"""
vin_list = [180,230,265]
eload_channel = 1
delay = 0.25 #s
max_iout = 4 #A
min_iout = 0 #A

batch = input("Batch: ")
unit = input("Unit: ")
der = 958

test = "LOAD TRANSIENT"
waveforms_folder = f'C:/Users/jbembate/Desktop/DER-{der}/Batch {batch}/Unit {unit}/{test}/'
path = path_maker(f'{waveforms_folder}')


"""DO NOT EDIT BELOW THIS LINE"""
##################################################################################
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
    sleep(1)
    for i in range(1,9):
        eload.channel[i].turn_off()
        eload.channel[i].short_off()
    sleep(1)

def scope_settings():
    """CHANNEL SETTINGS"""
    scope.channel_settings(state='ON', channel=1, scale=0.2, position=-4, label="BAGUHIN MO NA LANG TO", color='GREEN', rel_x_position=20, bandwidth=20, coupling='DCLimit', offset=0)
    scope.channel_settings(state='ON', channel=2, scale=2, position=-1, label="IKAW NA BAHALA SA SCOPE SETTINGS SA CODE", color='BLUE', rel_x_position=40, bandwidth=20, coupling='DCLimit', offset=0)
    scope.channel_settings(state='ON', channel=3, scale=10, position=-4, label="OUTPUT VOLTAGE", color='PINK', rel_x_position=60, bandwidth=20, coupling='DCLimit', offset=0)
    scope.channel_settings(state='OFF', channel=4)
    
    """MEASURE SETTINGS"""
    scope.measure(1, "MAX,MIN,RMS")
    scope.measure(2, "MAX,MIN,RMS")
    scope.measure(3, "MAX,MIN,RMS")
    # scope.measure(4, "MAX,MIN,PDELTA,MEAN,RMS")

    """HORIZONTAL SETTINGS"""
    scope.time_position(50)
    scope.record_length(50E6)
    scope.time_scale(delay)

    """ZOOM SETTINGS"""
    # scope.remove_zoom()
    # scope.add_zoom(rel_pos=21.727, rel_scale=1)
    
    """TRIGGER SETTINGS"""
    trigger_channel = 2
    trigger_level = 0.5
    trigger_edge = 'POS'
    scope.edge_trigger(trigger_channel, trigger_level, trigger_edge)

    scope.stop()
    sleep(2)

def operation():
    global waveform_counter

    scope_settings()

    for vin in vin_list:

        ac.voltage = vin
        ac.turn_on()

        sleep(4)

        for i in range(5):
            eload.channel[eload_channel].cc = min_iout
            eload.channel[eload_channel].turn_on()
            sleep(delay)
            eload.channel[eload_channel].cc = max_iout
            eload.channel[eload_channel].turn_on()
            sleep(delay)


        scope.run_single()

        for i in range(5):
            eload.channel[eload_channel].cc = min_iout
            eload.channel[eload_channel].turn_on()
            sleep(delay)
            eload.channel[eload_channel].cc = max_iout
            eload.channel[eload_channel].turn_on()
            sleep(delay)

        scope.stop()
        
        filename = f'{vin}Vac, {ac.frequency}Hz, {delay}s.png'
        scope.get_screenshot(filename, waveforms_folder)
        waveform_counter+=1
        print(filename)
        
        discharge_output()




def main():
    global waveform_counter
    discharge_output()
    operation()
        
if __name__ == "__main__":
    headers(test)
    main()
    footers(waveform_counter)