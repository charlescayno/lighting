# NOT INCLUDED IN THE FINAL TEST DATA


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
vin_list = [90,277]
iout_list = [0.06, 0.005]
vout = 3.5

test = "ZCD (MOSFET)"
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


# def message():
#     print(">> Set CH1 to Output Current")
#     print(">> Set CH4 to Input Voltage")
#     input("Press ENTER to continue")


# def scope_settings():
#     global ch1, ch2, ch3, ch4
#     global ton_time


#     scope.stop()
#     scope.remove_zoom()
#     scope.cursor(channel=2, cursor_set=1, X1=0, X2=0, Y1=0, Y2=0, type='HOR')
#     scope.cursor(channel=2, cursor_set=2, X1=0, X2=0, Y1=0, Y2=0, type='HOR')

#     ch1 = scope.channel_settings(state='OFF')
#     ch2 = scope.channel_settings(state='OFF')
#     ch3 = scope.channel_settings(state='OFF')
#     ch4 = scope.channel_settings(state='OFF')

#     scope.position_scale(time_position = 10, time_scale = 0.100)
#     scope.edge_trigger(4, 10, 'POS')

#     ch1 = scope.channel_settings(state='ON', channel=1, scale=0.5, position=1, label="Output Current", color='ORANGE', rel_x_position=30, bandwidth=500, coupling='DCLimit', offset=0)
#     ch4 = scope.channel_settings(state='ON', channel=4, scale=200, position=-2, label="Input Voltage", color='PINK', rel_x_position=70, bandwidth=500, coupling='AC', offset=0)
#     scope.cursor(channel=1, cursor_set=1, X1=0, X2=0.4, Y1=0, Y2=iout, type='HOR')
    
#     ton_time = 0.4 # initial setting
#     scope.cursor(channel=1, cursor_set=2, X1=0, X2=ton_time, Y1=0, Y2=1, type='PAIR')


def operation():

    global waveform_counter
    
    for iout in iout_list:
        
        for vin in vin_list:

                ac.voltage = vin
                ac.turn_on()

                eload.channel[1].cc = iout
                eload.channel[1].turn_on()
        
                scope.edge_trigger(4, 0, 'NEG')
                scope.run()
                input("Capture Rising? ")
                scope.stop()
                filename = f'{vin}Vac, {iout}A, rising.png'
                path = path_maker(f'{waveforms_folder}')
                scope.get_screenshot(filename, path)
                print(filename)
                waveform_counter += 1
                
                discharge_output()
    
    for iout in iout_list:
        eload.channel[1].cc = iout
        eload.channel[1].turn_on()
        
        for vin in vin_list:

                ac.voltage = vin
                ac.turn_on()

                eload.channel[1].cc = iout
                eload.channel[1].turn_on()
                

                scope.edge_trigger(4, 0, 'POS')
                scope.run()
                input("Capture Falling? ")
                scope.stop()
                filename = f'{vin}Vac, {iout}A, falling.png'
                path = path_maker(f'{waveforms_folder}')
                scope.get_screenshot(filename, path)
                print(filename)
                waveform_counter += 1

                discharge_output()

def main():
    global waveform_counter
    operation()
        
if __name__ == "__main__":
    headers(test)
    discharge_output()
    main()
    discharge_output()
    footers(waveform_counter)