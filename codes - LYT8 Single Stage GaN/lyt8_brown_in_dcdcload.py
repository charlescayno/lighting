##################################################################################
"""IMPORT DEPENDENCIES"""
from time import time, sleep
import sys
import os
import math
import numpy as np
import shutil
import os
import pandas as pd

from powi.equipment import ACSource, PowerMeter, ElectronicLoad, Oscilloscope, LEDControl, Keithley_DC_2230G
from powi.equipment import headers, create_folder, footers, waveform_counter, soak, convert_argv_to_int_list, tts, prompt
from powi.equipment import excel_to_df, df_to_excel, image_to_excel, col_row_extractor, get_anchor
from powi.equipment import create_header_list, export_to_excel, export_screenshot_to_excel
from powi.equipment import path_maker, remove_file

import getpass
username = getpass.getuser().lower()

from datetime import datetime
now = datetime.now()
date = now.strftime('%m%d')
##################################################################################

operation = "CV"
test = "Brown In and Brown Out"
condition = input(">> Condition: ")
load_type = input(">> Load Type (cc/cr/led): ")
if load_type == 'cc' or load_type == 'cr': von = int(input(">> Enter Von point: "))

waveforms_folder = f'C:/Users/{username}/Desktop/APPS EVAL/Single Stage GAN/Test Data ({operation})/{date}/{test}/{condition}/{load_type}'
path = path_maker(f'{waveforms_folder}')

"""COMMS"""
ac_source_address = 5
source_power_meter_address = 30 
load_power_meter_address = 2
eload_address = 8
scope_address = "10.125.10.101"
sig_gen_address = '11'
dimming_power_meter_address = 21
eload_channel = 5

"""USER INPUT"""

# vout = 50
# iout_nom = 1.5
vout = 40
iout_nom = 1.25
cr_load = vout/iout_nom

percent_list = [1, 0.75, 0.5, 0.25, 0.1, 0]
percent_list = [1]


zoom_enable = False

iout_list = [percent * iout_nom for percent in percent_list]
cr_list = [(cr_load/percent if percent != 0 else 0) for percent in percent_list]
led_list = [40]

if load_type == 'cc': load_list = iout_list
if load_type == 'cr': load_list = cr_list
if load_type == 'led': load_list = led_list

"""BROWN-IN SETTINGS"""
start_voltage = 0
end_voltage = 300
slew_rate = 0.5
frequency = 60
time_fixvoltage = 60

"""SCOPE SETTINGS"""
trig_channel = 3
trig_level = 0.4
trig_edge = 'POS'

time_position = 10
time_scale = 0.1

"""
MEASUREMENT SETTINGS OPTIONS: "MAX,MIN,RMS,MEAN,PDELta"
"""

ch1_enable = 'ON'
ch2_enable = 'ON'
ch3_enable = 'ON'
ch4_enable = 'ON'

"""CHANNEL 1"""
ch1_scale = 0.3
ch1_position = -4
ch1_bw = 20
ch1_rel_x_position = 20
ch1_label = "DC-DC IOUT"
ch1_measure = "MAX"
ch1_color = "GREEN"
ch1_coupling = "DCLimit"

"""CHANNEL 2"""
ch2_scale = 150
ch2_position = 0
ch2_bw = 20
ch2_rel_x_position = 40
ch2_label = "VIN"
ch2_measure = "MAX,RMS"
ch2_color = "YELLOW"
ch2_coupling = "AC"

"""CHANNEL 3"""
ch3_scale = 2
ch3_position = -4
ch3_bw = 500
ch3_rel_x_position = 60
ch3_label = "PRI_IDS"
ch3_measure = "MAX"
ch3_color = "LIGHT_BLUE"
ch3_coupling = "DCLimit"

"""CHANNEL 4"""
ch4_scale = 10
ch4_position = -4
ch4_bw = 20
ch4_rel_x_position = 80
ch4_label = "DC-DC VOUT"
ch4_measure = "MAX"
ch4_color = "PINK"
ch4_coupling = "DCLimit"


"""DO NOT EDIT BELOW THIS LINE"""
####################################################################################################################################################################

"""EQUIPMENT INITIALIZE"""
ac = ACSource(ac_source_address)
pms = PowerMeter(source_power_meter_address)
pml = PowerMeter(load_power_meter_address)
eload = ElectronicLoad(eload_address)
scope = Oscilloscope(scope_address)


"""GENERIC FUNCTIONS"""

def discharge_output(times=1):
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

def screenshot_looper(filename, path):
    iter = 0
    while '' == input(">> Press ENTER to continue capture. To stop press other keys to stop: "):
        fname = filename.split('.png')[0] + f'_{iter}' + '.png'
        iter += 1
        screenshot(fname, path)

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

def brown_in(load):

    test = "Brown In"


    if load == 0:
            eload.channel[eload_channel].turn_off()
            if load_type == 'cc': filename = f"{test}, {start_voltage}-{end_voltage}Vac, {load} A (0 CC load).png"
            if load_type == 'cr': filename = f"{test}, {start_voltage}-{end_voltage}Vac, {load} Ohms (0 CR load).png"
    else:
        if load_type == 'cc':
            eload.channel[eload_channel].von = von
            eload.channel[eload_channel].cc = load
            eload.channel[eload_channel].turn_on()
            percent = int(load*100/iout_nom)
            load = round(load, 2)
            filename = f"{test}, {start_voltage}-{end_voltage}Vac, {load} A ({percent} CC load), Von = {von} V.png"
        if load_type == 'cr':
            eload.channel[eload_channel].cr = load
            eload.channel[eload_channel].turn_on()
            percent = int(cr_load*100/load)
            load = round(load, 2)
            filename = f"{test}, {start_voltage}-{end_voltage}Vac, {load} Ohms ({percent} CR load).png"
        if load_type == 'led':
            filename = f"{test}, {start_voltage}-{end_voltage}Vac, {load}V.png"


    test_time = ((end_voltage-start_voltage)/slew_rate)+time_fixvoltage
    scope_time = roundup(test_time)
    delay = scope_time-test_time
    time_scale = scope_time/10
    print(filename)
    print(f"Estimated time: {scope_time/60} mins.")
    scope.time_scale(time_scale)
    scope.run()
    soak(int(delay))

    browning(start_voltage, end_voltage, slew_rate, frequency)
    fixvoltage(end_voltage,time_fixvoltage)
    scope.stop()

    discharge_output()

    input(">> Capture Vin with GATE.")
    screenshot(filename, path)
    screenshot_looper(filename, path) 

def brown_out(load):

    test = "Brown Out"

    

    if load == 0:
        eload.channel[eload_channel].turn_off()
        if load_type == 'cc': filename = f"{test}, {end_voltage}-{start_voltage}Vac, {load} A (0 CC load).png"
        if load_type == 'cr': filename = f"{test}, {end_voltage}-{start_voltage}Vac, {load} Ohms (0 CR load).png"
    else:
        if load_type == 'cc':
            eload.channel[eload_channel].von = von
            eload.channel[eload_channel].cc = load
            eload.channel[eload_channel].turn_on()
            percent = int(load*100/iout_nom)
            load = round(load, 2)
            filename = f"{test}, {end_voltage}-{start_voltage}Vac, {load} A ({percent} CC load), Von = {von} V.png"
        if load_type == 'cr':
            eload.channel[eload_channel].cr = load
            eload.channel[eload_channel].turn_on()
            percent = int(cr_load*100/load)
            load = round(load, 2)
            filename = f"{test}, {end_voltage}-{start_voltage}Vac, {load} Ohms ({percent} CR load).png"
        if load_type == 'led':
            filename = f"{test}, {start_voltage}-{end_voltage}Vac, {load}V.png"





    test_time = ((end_voltage-start_voltage)/slew_rate)+time_fixvoltage
    scope_time = roundup(test_time)
    delay = scope_time-test_time
    time_scale = scope_time/10
    print(filename)
    print(f"Estimated time: {scope_time/60} mins.")
    scope.time_scale(time_scale)
    scope.run()
    
    # start of test
    fixvoltage(end_voltage, time_fixvoltage)
    browning(end_voltage, start_voltage, slew_rate, frequency)
    soak(int(delay))
    scope.stop()
    discharge_output()
    input(">> Capture Vin with GATE.")
    screenshot(filename, path)
    screenshot_looper(filename, path)

def brown_in_brown_out(load):

    if load == 0:
        eload.channel[eload_channel].turn_off()
        if load_type == 'cc': filename = f"{test}, {start_voltage}-{end_voltage}Vac, {load} A (0 CC load).png"
        if load_type == 'cr': filename = f"{test}, {start_voltage}-{end_voltage}Vac, {load} Ohms (0 CR load).png"
    else:
        if load_type == 'cc':
            eload.channel[eload_channel].cc = load
            eload.channel[eload_channel].turn_on()
            percent = int(load*100/iout_nom)
            filename = f"{test}, {start_voltage}-{end_voltage}Vac, {load} A ({percent} CC load).png"
        if load_type == 'cr':
            eload.channel[eload_channel].cr = load
            eload.channel[eload_channel].turn_on()
            percent = int(cr_load*100/load)
            filename = f"{test}, {start_voltage}-{end_voltage}Vac, {load} Ohms ({percent} CR load).png"
        if load_type == 'led':
            filename = f"{test}, {start_voltage}-{end_voltage}Vac, {load}V.png"



    print(filename)

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

    input(">> Capture Vin with GATE.")
    screenshot(filename, path)
    screenshot_looper(filename, path)

def brown_out_brown_in(load):


    if load == 0:
        eload.channel[eload_channel].turn_off()
        if load_type == 'cc': filename = f"{test}, {start_voltage}-{end_voltage}Vac, {load} A (0 CC load).png"
        if load_type == 'cr': filename = f"{test}, {start_voltage}-{end_voltage}Vac, {load} Ohms (0 CR load).png"
    else:
        if load_type == 'cc':
            eload.channel[eload_channel].cc = load
            eload.channel[eload_channel].turn_on()
            percent = int(load*100/iout_nom)
            filename = f"{test}, {start_voltage}-{end_voltage}Vac, {load} A ({percent} CC load).png"
        if load_type == 'cr':
            eload.channel[eload_channel].cr = load
            eload.channel[eload_channel].turn_on()
            percent = int(cr_load*100/load)
            filename = f"{test}, {start_voltage}-{end_voltage}Vac, {load} Ohms ({percent} CR load).png"
        if load_type == 'led':
            filename = f"{test}, {start_voltage}-{end_voltage}Vac, {load}V.png"



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

    input(">> Capture Vin with GATE.")
    screenshot(filename, path)
    screenshot_looper(filename, path)


def operation():

    for load in load_list:
        soak(2)
        brown_in(load)
        brown_out(load)
        # brown_in_brown_out(load)
        # brown_out_brown_in(load)
        

def main():
    global waveform_counter
    discharge_output(times=3)
    scope_settings()
    operation()
        
if __name__ == "__main__":
    headers(test)
    main()
    footers(waveform_counter)