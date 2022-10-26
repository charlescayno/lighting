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

test = f"Startup Waveforms"
operation = "CV"


condition = "DCDC Load - Aux (Battery Powered)"
condition = "DCDC Load - Aux (Battery Powered) - without LYT8 IOUT"
condition = "DCDC Load - Aux (Battery Powered) - without LYT8 IOUT - with DIM=10V"
condition = "DCDC Load - Aux (Battery Powered) - without LYT8 IOUT - DIM connected to HBP"
condition = "DCDC Load - Aux (Battery Powered) - with LYT8 IOUT - DIM connected to HBP"

load_type = "led"
excel_name = "summary"
waveforms_folder = f'C:/Users/{username}/Desktop/APPS EVAL/Single Stage GAN/Test Data ({operation})/{date}/{test}/{condition}'
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

iout_channel = 1
vin_channel = 2
ids_channel = 3
vout_channel = 4

"""USER INPUT"""

vin_list = [90,115,130,180,230,265,277,300]

led_list = [40,25]

zoom_enable = True
"""SCOPE SETTINGS"""
trig_channel = 4
trig_level = 2.5
trig_edge = 'POS'

time_position = 10
time_scale = 0.1

zoom_enable = False



ch1_enable = 'ON'
ch2_enable = 'ON'
ch3_enable = 'ON'
ch4_enable = 'ON'

"""CHANNEL 1"""
ch1_scale = 1
ch1_position = -4
ch1_bw = 20
ch1_rel_x_position = 20 #from left
ch1_label = "LYT8 IOUT" 
ch1_measure = "MAX,MIN,MEAN"
ch1_color = "GREEN"
ch1_coupling = "DCLimit"

"""CHANNEL 2"""
ch2_scale = 10
ch2_position = -4
ch2_bw = 20 #20, 200, 500 in Mhz
ch2_rel_x_position = 40
ch2_label = "LYT8 VOUT"
ch2_measure = "MAX,MIN,MEAN"
ch2_color = "YELLOW"
ch2_coupling = "DCLimit"

"""CHANNEL 3"""
ch3_scale = 0.2 # 0.2
ch3_position = -4
ch3_bw = 20
ch3_rel_x_position = 60
ch3_label = "DC-DC IOUT"
ch3_measure = "MAX,MIN,MEAN"
ch3_color = "LIGHT_BLUE"
ch3_coupling = "DCLimit"

"""CHANNEL 4"""
ch4_scale = 10
ch4_position = -4
ch4_bw = 20
ch4_rel_x_position = 80
ch4_label = "DC-DC VOUT"
ch4_measure = "MAX,MIN,MEAN"
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
# led_ctrl = LEDControl()

"""GENERIC FUNCTIONS"""

def discharge_output(times):    
    ac.turn_off()
    for i in range(times):
        for i in range(1,9):
            eload.channel[i].cr = 10
            eload.channel[i].turn_on()
        sleep(0.5)

        for i in range(1,9):
            eload.channel[i].turn_off()
            eload.channel[i].short_off()
        sleep(1)

        for i in range(1,9):
            eload.channel[i].cc = 1
            eload.channel[i].short_on()
            eload.channel[i].turn_on()
        sleep(0.5)

        for i in range(1,9):
            eload.channel[i].turn_off()
            eload.channel[i].short_off()
        sleep(0.5)

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

##################################################################################################


def operation():



    for led in led_list:
        input(f">> Change LED load to {led}V: ")
        
        for vin in vin_list:

            scope.run_single()
            sleep(3)

            ac.voltage = vin
            ac.turn_on()
            sleep(2)
            
            discharge_output(times=1)

            filename = f"Startup - {vin}Vac, LED = {led}V.png"
            screenshot(filename, path)
            
    

        

def main():
    global waveform_counter
    discharge_output(times=1)
    scope_settings()
    operation()
        
if __name__ == "__main__":
    headers(test)
    print(condition)
    main()
    footers(waveform_counter)