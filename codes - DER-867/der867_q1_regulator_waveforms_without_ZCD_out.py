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
# vin_list = [230]
# vin_list = convert_argv_to_int_list(sys.argv[1])




test = "Q1 Regulator Waveforms (disabled ZCD_OUT)"
waveforms_folder = f'C:/Users/ccayno/Desktop/DER/DER-867/Final Test Data/{test}'

"""DO NOT EDIT BELOW THIS LINE"""
##################################################################################

"""EQUIPMENT INITIALIZE"""
ac = ACSource(ac_source_address)
pms = PowerMeter(source_power_meter_address)
pml = PowerMeter(load_power_meter_address)
eload = ElectronicLoad(eload_address)
scope = Oscilloscope(scope_address)


def scope_settings():
    global condition
    scope.channel_settings(state='ON', channel=1, scale=200, position=3, label="Input Voltage", color='YELLOW', rel_x_position=30, bandwidth=20, coupling='AC', offset=0)
    scope.channel_settings(state='ON', channel=2, scale=2, position=-1, label="ZCD_IN", color='ORANGE', rel_x_position=50, bandwidth=20, coupling='DCLimit', offset=0)
    scope.channel_settings(state='ON', channel=3, scale=5, position=-4, label="Q1 Regulator", color='GREEN', rel_x_position=70, bandwidth=20, coupling='DCLimit', offset=0)
    scope.channel_settings(state='OFF', channel=4, scale=1, position=0, label="RelayOFF Pulse", color='PINK', rel_x_position=80, bandwidth=20, coupling='DCLimit', offset=0)
    scope.measure(1, "RMS,PER,FREQ")
    scope.measure(2, "MAX,PER,FREQ")
    scope.measure(3, "MAX,PER,FREQ")
    # scope.measure(4, "MAX,RMS,FREQ")

    scope.time_position(50)
    scope.time_scale(10E-3)
    scope.record_length(50E6)

    trigger_channel = 3
    trigger_level = 1
    trigger_edge = 'POS'

    scope.edge_trigger(trigger_channel, trigger_level, trigger_edge)

    scope.remove_zoom()

    scope.stop()

def operation():
    global waveform_counter
    
    scope_settings()
    
    a = 0

    for vin in vin_list:

        ac.voltage = vin
        ac.turn_on()

        if a == 0:
            input("Turn on relay. Press ENTER to continue...")
            a += 1

        sleep(2)

        scope.run_single()
        sleep(3)

        input("Adjust cursor.")

        filename = f'{vin}Vac, {ac.frequency}Hz.png'
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