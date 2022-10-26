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
scope_address = "10.125.10.184"

"""USER INPUT"""
led_list = [42,30]
vin_list = [180,230,265]
timeperdiv = 2 # s/div

batch = input("Batch: ")
unit = input("Unit: ")
der = 935

print("Charles Cayno | 03-18-2022")
print()
print("Set the following:")
print("Channel 1 - Output Current")
print("Channel 2 - SHORT CIRCUIT CURRENT")
print("Channel 3 - Output Voltage")
input("Press ENTER to continue...")
print()

test = "OUTPUT RUNNING SHORT (TSP15H150S S1G)"
waveforms_folder = f'C:/Users/ccayno/Desktop/DER/DER-{der}/2 Marketing Sample Units/Batch {batch}/Unit {unit}/{test}/'
path = path_maker(f'{waveforms_folder}')



"""DO NOT EDIT BELOW THIS LINE"""
##################################################################################
ac = ACSource(ac_source_address)
pms = PowerMeter(source_power_meter_address)
pml = PowerMeter(load_power_meter_address)
eload = ElectronicLoad(eload_address)
scope = Oscilloscope(scope_address)
# led = LEDControl()


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
    scope.channel_settings(state='ON', channel=1, scale=0.2, position=-4, label="OUTPUT CURRENT", color='GREEN', rel_x_position=20, bandwidth=20, coupling='DCLimit', offset=0)
    scope.channel_settings(state='ON', channel=2, scale=10, position=-4, label="SHORT CIRCUIT CURRENT", color='BLUE', rel_x_position=50, bandwidth=20, coupling='DCLimit', offset=0)
    scope.channel_settings(state='ON', channel=3, scale=10, position=-4, label="OUTPUT VOLTAGE", color='PINK', rel_x_position=80, bandwidth=20, coupling='DCLimit', offset=0)
    scope.channel_settings(state='OFF', channel=4)
    
    """MEASURE SETTINGS"""
    scope.measure(1, "MAX,MIN,RMS")
    scope.measure(2, "MAX,MIN,RMS")
    scope.measure(3, "MAX,MIN,RMS")
    # scope.measure(4, "MAX,MIN,PDELTA,MEAN,RMS")

    """HORIZONTAL SETTINGS"""
    scope.time_position(10)
    scope.record_length(50E6)
    scope.time_scale(timeperdiv)

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

    for led in led_list:
        input(f"Change LED to {led}V")

        for vin in vin_list:

            ac.voltage = vin
            ac.turn_on()
            
            scope.run()
            sleep(4)
            
            soak(2*timeperdiv) 
            
            print(">> Short Output...")
            soak(4*timeperdiv)
            print(">> Remove Short...")
            soak(4*timeperdiv)
            
            scope.stop()

            filename = f'{led}V, {vin}Vac, {ac.frequency}Hz, Running Short.png'
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