## AUGUST 12, 2022 - THIS CODE CAN BE A DEFAULT STARTUP TEMPLATE
"""IMPORT DEPENDENCIES"""
from time import time, sleep
import sys
import os
import math
from powi.equipment import ACSource, PowerMeter, ElectronicLoad, Oscilloscope, LEDControl
from powi.equipment import headers, create_folder, footers, waveform_counter, soak, convert_argv_to_int_list, tts, prompt
from powi.equipment import excel_to_df, df_to_excel, image_to_excel, col_row_extractor, get_anchor
from powi.default_scope_settings import *
import shutil
import openpyxl
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
from openpyxl.utils import get_column_letter
import matplotlib.pyplot as plt
from openpyxl import Workbook, load_workbook
from filemanager import path_maker, remove_file
import winsound as ws
from playsound import playsound
import numpy as np
####################################################################################################################################################################

test = "STARTUP PROFILE"
operation = "CV"
waveforms_folder = f'C:/Users/ccayno/Desktop/APPS EVAL/Single Stage GAN/Test Data ({operation})/{test}'
path = path_maker(f'{waveforms_folder}')

"""COMMS"""
ac_source_address = 5
source_power_meter_address = 30 
load_power_meter_address = 2
eload_address = 8
scope_address = "10.125.10.129"
sig_gen_address = '11'
dimming_power_meter_address = 21
eload_channel = 5

"""USER INPUT"""


vin_list = [90,115,130,180,230,265,300]
iout_list = [1.1]


"""BROWN-IN SETTINGS"""
start_voltage = 0
end_voltage = 277
slew_rate = 0.5
frequency = 60
time_fixvoltage = 60



"""SCOPE SETTINGS"""
trig_channel = 2
trig_level = 10
trig_edge = 'POS'

time_position = 10
time_scale = 0.100

"""
MEASUREMENT SETTINGS OPTIONS: "MAX,MIN,RMS,MEAN,PDELta"
"""

ch1_enable = 'ON'
ch2_enable = 'ON'
ch3_enable = 'ON'
ch4_enable = 'ON'

"""CHANNEL 1"""
ch1_scale = 1
ch1_position = -3
ch1_bw = 20
ch1_rel_x_position = 20
ch1_label = "BPP"
ch1_measure = "MAX"
ch1_color = "GREEN"
ch1_coupling = "DCLimit"


"""CHANNEL 2"""
ch2_scale = 200
ch2_position = -4
ch2_bw = 500
ch2_rel_x_position = 40
ch2_label = "PRI_VDS"
ch2_measure = "MAX"
ch2_color = "LIGHT_BLUE"
ch2_coupling = "DCLimit"



"""CHANNEL 3"""
ch3_scale = 10
ch3_position = -4
ch3_bw = 20
ch3_rel_x_position = 60
ch3_label = "VR PIN"
ch3_measure = "MAX"
ch3_color = "PINK"
ch3_coupling = "DCLimit"



"""CHANNEL 4"""
ch4_scale = 5
ch4_position = -3
ch4_bw = 20
ch4_rel_x_position = 80
ch4_label = "HBP"
ch4_measure = "MAX"
ch4_color = "YELLOW"
ch4_coupling = "DCLimit"


"""DO NOT EDIT BELOW THIS LINE"""
####################################################################################################################################################################

"""EQUIPMENT INITIALIZE"""
ac = ACSource(ac_source_address)
pms = PowerMeter(source_power_meter_address)
pml = PowerMeter(load_power_meter_address)
eload = ElectronicLoad(eload_address)
scope = Oscilloscope(scope_address)
# led_ctrl = LEDControl()

"""GENERIC FUNCTIONS"""

def discharge_output(times):
    ac.turn_off()
    for i in range(times):
        for i in range(1,9):
            eload.channel[i].cc = 1
            eload.channel[i].turn_on()
            eload.channel[i].short_on()
        sleep(1)

        for i in range(1,9):
            eload.channel[i].turn_off()
            eload.channel[i].short_off()
        sleep(2)

def scope_settings():
    
    scope.channel_settings(state=ch1_enable, channel=1, scale=ch1_scale, position=ch1_position, label=ch1_label,
                            color=ch1_color, rel_x_position=ch1_rel_x_position, bandwidth=ch1_bw, coupling=ch1_coupling, offset=0)
        
    scope.channel_settings(state=ch2_enable, channel=2, scale=ch2_scale, position=ch2_position, label=ch2_label,
                            color=ch2_color, rel_x_position=ch2_rel_x_position, bandwidth=ch2_bw, coupling=ch2_coupling, offset=0)
    
    scope.channel_settings(state=ch3_enable, channel=3, scale=ch3_scale, position=ch3_position, label=ch3_label,
                            color=ch3_color, rel_x_position=ch3_rel_x_position, bandwidth=ch3_bw, coupling=ch3_coupling, offset=0)
    
    scope.channel_settings(state=ch4_enable, channel=4, scale=ch4_scale, position=ch4_position, label=ch4_label,
                            color=ch4_color, rel_x_position=ch4_rel_x_position, bandwidth=ch4_bw, coupling=ch4_coupling, offset=0)
    
    if ch1_enable != 'OFF': scope.measure(1, ch1_measure)
    if ch2_enable != 'OFF': scope.measure(2, ch2_measure)
    if ch3_enable != 'OFF': scope.measure(3, ch3_measure)
    if ch4_enable != 'OFF': scope.measure(4, ch4_measure)

    scope.record_length(50E6)
    scope.time_position(time_position)
    scope.time_scale(time_scale)

    # scope.remove_zoom()
    # scope.add_zoom(rel_pos=50, rel_scale=0.04)
    
    trigger_channel = trig_channel
    trigger_level = trig_level
    trigger_edge = trig_edge
    scope.edge_trigger(trigger_channel, trigger_level, trigger_edge)

    scope.stop()
    
def screenshot(filename, path):
    global waveform_counter

    scope.get_screenshot(filename, path)
    print(filename)
    waveform_counter += 1

def browning(start, end, slew, frequency):

    if start > end:
        print(f"brownout: {start} -> {end} Vac")
        for voltage in np.arange(start, end+1, -slew):
            ac.voltage = voltage
            ac.frequency = frequency
            ac.turn_on()
            sleep(1)
    if start < end:
        print(f"brownin: {start} -> {end} Vac")
        for voltage in np.arange(start, end+1, slew):
            ac.voltage = voltage
            ac.frequency = frequency
            ac.turn_on()
            sleep(1)

def fixvoltage(voltage, soak):
    ac.voltage = voltage
    ac.frequency = ac.set_freq(voltage)
    ac.turn_on()
    sleep(soak)

def roundup(x):
    return int(math.ceil(x / 100.0)) * 100

def brown_in(led):

    test_time = ((end_voltage-start_voltage)/slew_rate)+time_fixvoltage
    scope_time = roundup(test_time)
    delay = scope_time-test_time
    time_scale = scope_time/10
    print(f"Estimated time: {scope_time/60} mins.")
    scope.time_scale(time_scale)
    scope.run()
    soak(int(delay))

    # start of test
    browning(start_voltage, end_voltage, slew_rate, frequency)
    fixvoltage(end_voltage,time_fixvoltage)
    scope.stop()

    discharge_output()

    filename = f"{led}V, {start_voltage}-{end_voltage} Vac, {slew_rate}V per s.png"
    screenshot(filename, path)

def brown_out(led):

    test_time = ((end_voltage-start_voltage)/slew_rate)+time_fixvoltage
    scope_time = roundup(test_time)
    delay = scope_time-test_time
    time_scale = scope_time/10
    print(f"Estimated time: {scope_time/60} mins.")
    scope.time_scale(time_scale)
    scope.run()
    
    # start of test
    fixvoltage(end_voltage, time_fixvoltage)
    browning(end_voltage, start_voltage, slew_rate, frequency)
    soak(int(delay))
    scope.stop()
    discharge_output()

    filename = f"{led}V, {end_voltage}-{start_voltage} Vac, {slew_rate}V per s.png"
    screenshot(filename, path)

def brown_in_brown_out(led):

    test_time = 2*((end_voltage-start_voltage)/slew_rate)
    scope_time = roundup(test_time)
    delay = scope_time-test_time
    time_scale = scope_time/10
    print(f"Estimated time: {scope_time/60} mins.")
    scope.time_scale(time_scale)
    scope.run()
    soak(int(delay/2))
    # start of test
    browning(start_voltage, end_voltage, slew_rate, frequency)
    browning(end_voltage, start_voltage, slew_rate, frequency)
    soak(int(delay/2))
    scope.stop()
    discharge_output()

    filename = f"{led}V, {start_voltage}-{end_voltage}-{start_voltage} Vac, {slew_rate}V per s.png"
    screenshot(filename, path)

def brown_out_brown_in(led):

    test_time = 2*((end_voltage-start_voltage)/slew_rate)
    scope_time = roundup(test_time)
    delay = scope_time-test_time
    time_scale = scope_time/10
    print(f"Estimated time: {scope_time/60} mins.")
    scope.time_scale(time_scale)
    scope.run()
    t_fixvoltage = int(delay/2)
    # start of test
    fixvoltage(end_voltage, t_fixvoltage)
    browning(end_voltage, start_voltage, slew_rate, frequency)
    browning(start_voltage, end_voltage, slew_rate, frequency)
    fixvoltage(end_voltage, t_fixvoltage)
    scope.stop()
    discharge_output()

    filename = f"{led}V, {end_voltage}-{start_voltage}-{end_voltage} Vac, {slew_rate}V per s.png"
    screenshot(filename, path)



def operation():

    for iout in iout_list:
        
        for vin in vin_list:

            input(">> Continue test?")

            eload.channel[eload_channel].cc = iout
            eload.channel[eload_channel].turn_on()

            scope.run_single()
            sleep(3)

            ac.voltage = vin
            ac.turn_on()
            sleep(5)
            
            discharge_output(times=1)

            filename = f"Startup {time_scale}s_per_div - {vin}Vac, {iout}A.png"
            screenshot(filename, path)
            

def main():
    global waveform_counter
    discharge_output(times=1)
    scope_settings()
    operation()
        
if __name__ == "__main__":
    headers(test)
    main()
    footers(waveform_counter)