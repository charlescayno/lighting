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
vin_list = [180,230,265]
led_list = [42,36,30]


batch = input("Batch: ")
unit = input("Unit: ")
der = 935

print("Charles Cayno | 03-18-2022")
print()
print("Set the following:")
print("Channel 1 - Output Current")
print("Channel 3 - Output Voltage")
input("Press ENTER to continue...")
print()

test = "OUTPUT CURRENT RIPPLE"
waveforms_folder = f'C:/Users/ccayno/Desktop/DER-{der} Marketing Samples/Batch {batch}/Unit {unit}/{test}/'
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
    sleep(2)
    for i in range(1,9):
        eload.channel[i].turn_off()

def scope_settings():
    """CHANNEL SETTINGS"""
    scope.channel_settings(state='ON', channel=1, scale=0.2, position=-4, label="OUTPUT CURRENT", color='GREEN', rel_x_position=40, bandwidth=20, coupling='DCLimit', offset=0)
    scope.channel_settings(state='OFF', channel=2)
    scope.channel_settings(state='ON', channel=3, scale=10, position=-1, label="OUTPUT VOLTAGE", color='PINK', rel_x_position=60, bandwidth=20, coupling='DCLimit', offset=0)
    scope.channel_settings(state='OFF', channel=4)
    
    """MEASURE SETTINGS"""
    scope.measure(1, "MAX,MIN,PDELTA,MEAN,RMS")
    # scope.measure(2, "MAX,MIN")
    scope.measure(3, "MAX,MIN,PDELTA,MEAN,RMS")
    # scope.measure(4, "MAX,MIN,PDELTA,MEAN,RMS")

    """HORIZONTAL SETTINGS"""
    scope.time_position(50)
    scope.record_length(50E6)
    scope.time_scale(0.004)

    """ZOOM SETTINGS"""
    # scope.remove_zoom()
    # scope.add_zoom(rel_pos=21.727, rel_scale=1)
    
    """TRIGGER SETTINGS"""
    trigger_channel = 1
    trigger_level = 1
    trigger_edge = 'POS'
    scope.edge_trigger(trigger_channel, trigger_level, trigger_edge)

def operation():
    global waveform_counter
    
    scope_settings()
    
    for led in led_list:

        input(f"Change led load to {led}V")

        for vin in vin_list:

            ac.voltage = vin
            ac.turn_on()

            sleep(5)

            scope.run_single()

            input("Capture waveform?")

            filename = f'{led}V, {vin}Vac, {ac.frequency}Hz.png'
            
            scope.get_screenshot(filename, path)
            print(filename)
            waveform_counter += 1

        discharge_output()
        
            
def main():
    global waveform_counter
    operation()
        
if __name__ == "__main__":
    headers(test)
    main()
    footers(waveform_counter)