##################################################################################
"""IMPORT DEPENDENCIES"""
from time import time, sleep
import sys
import os
import math
# from lyt8_load_step import screenshot
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


"""COMMS"""
ac_source_address = 5
source_power_meter_address = 30 
load_power_meter_address = 2
eload_address = 8
scope_address = "10.125.10.101"
sig_gen_address = '11'
dimming_power_meter_address = 21
eload_channel = 5
dc_source_address = 4 



vin_list = [90,115,230,265,277,300]
led_list = [48] # V
vout = 40 # V
iout = 1.25 # A
soak_time = 2 # s

start_bit = 0
end_bit = 255
inc_bit = 1


"""SCOPE CHANNEL"""
iout_channel = 1
# vin_channel = 2
ids_channel = 3
vout_channel = 2

"""SCOPE SETTINGS"""
trig_channel = 3
trig_level = 13
trig_edge = 'POS'

time_position = 30
time_scale = 0.1

"""ZOOM SETTINGS"""
zoom_enable = False
zoom_pos = 20
zoom_rel_scale = 1


"""
MEASUREMENT SETTINGS OPTIONS: "MAX,MIN,RMS,MEAN,PDELta"
"""

ch1_enable = 'ON'
ch2_enable = 'ON'
ch3_enable = 'ON'
ch4_enable = 'ON'

"""CHANNEL 1"""
ch1_scale = 10
ch1_position = -4
ch1_bw = 20
ch1_rel_x_position = 20 #from left
ch1_label = "DC-DC VOUT"
ch1_measure = "MAX,MIN,MEAN"
ch1_color = "PINK"
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
ch2_offset = 0

"""CHANNEL 3"""
ch3_scale = 2 # 0.2
ch3_position = -4
ch3_bw = 20
ch3_rel_x_position = 60
ch3_label = "DIM"
ch3_measure = "MAX,MIN,MEAN"
ch3_color = "LIGHT_BLUE"
ch3_coupling = "DCLimit"

"""CHANNEL 4"""
ch4_scale = 0.2
ch4_position = -4
ch4_bw = 20
ch4_rel_x_position = 80
ch4_label = "DC-DC IOUT" 
ch4_measure = "MAX,MIN,MEAN"
ch4_color = "GREEN"
ch4_coupling = "DCLimit"
ch4_offset = 0


test = "DALI"
operation = "CV"
condition = input(">> Condition: ")
# condition = "test"
# load_type = input(">> Load Type (cc/cr/led): ")
load_type = "cc"
excel_name = "summary"
waveforms_folder = f'C:/Users/{username}/Desktop/APPS EVAL/Single Stage GAN/Test Data ({operation})/{date}/{test}/{condition}'
path = path_maker(f'{waveforms_folder}')
"""DO NOT EDIT BELOW THIS LINE"""
##################################################################################

"""EQUIPMENT INITIALIZE"""
ac = ACSource(ac_source_address)
pms = PowerMeter(source_power_meter_address)
pml = PowerMeter(load_power_meter_address)
# # pmdc = PowerMeter(dimming_power_meter_address)
eload = ElectronicLoad(eload_address)
scope = Oscilloscope(scope_address)
dc = Keithley_DC_2230G(dc_source_address)
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
    # INCLUDE
    
    scope.channel_settings(state=ch1_enable, channel=1, scale=ch1_scale, position=ch1_position, label=ch1_label,
                            color=ch1_color, rel_x_position=ch1_rel_x_position, bandwidth=ch1_bw, coupling=ch1_coupling, offset=0)
        
    scope.channel_settings(state=ch2_enable, channel=2, scale=ch2_scale, position=ch2_position, label=ch2_label,
                            color=ch2_color, rel_x_position=ch2_rel_x_position, bandwidth=ch2_bw, coupling=ch2_coupling, offset=ch2_offset)
    
    scope.channel_settings(state=ch3_enable, channel=3, scale=ch3_scale, position=ch3_position, label=ch3_label,
                            color=ch3_color, rel_x_position=ch3_rel_x_position, bandwidth=ch3_bw, coupling=ch3_coupling, offset=0)
    
    scope.channel_settings(state=ch4_enable, channel=4, scale=ch4_scale, position=ch4_position, label=ch4_label,
                            color=ch4_color, rel_x_position=ch4_rel_x_position, bandwidth=ch4_bw, coupling=ch4_coupling, offset=ch4_offset)
    
    if ch1_enable != 'OFF': scope.measure(1, ch1_measure)
    if ch2_enable != 'OFF': scope.measure(2, ch2_measure)
    if ch3_enable != 'OFF': scope.measure(3, ch3_measure)
    if ch4_enable != 'OFF': scope.measure(4, ch4_measure)

    scope.record_length(50E6)
    scope.time_position(time_position)
    scope.time_scale(time_scale)

    if zoom_enable:
        scope.remove_zoom()
        scope.add_zoom(rel_pos=zoom_pos, rel_scale=zoom_rel_scale)
    
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

def set_dim_voltage(dim):
    dc.set_volt_curr(channel ='CH1', voltage = dim, current = 1.0)
    dc.channel_state(channel ='CH1', state = 'ON')
    soak(1)

def dali(bit):
    a = bit-1
    b = 253/3
    y = (10**(a/b - 1))/100 # dali equation in percentage
    dim = y*10 # dim voltage equivalent at 10 V

    set_dim_voltage(dim) # control dc source
    return dim

def reset_dc_source():
    set_dim_voltage(1)
    soak(2)
    set_dim_voltage(0)

def operation():

    # start_bit = 0
    end_bit = 254
    dim_list = [0,1,2,3,4,5,6,7,8,9,10]
    # vin_list = [90,115,230,265,277,300]
    vin_list = [90,115,230,300]
    led_list = [40] # V

    for led in led_list:
        input(f"Change led load to {led}V")

        for dim in dim_list:

            for vin in vin_list:


                

                set_dim_voltage(dim)

                ac.voltage = vin
                ac.turn_on()

                sleep(2)

                scope.run_single()
                
                
                sleep(3)
                
                set_dim_voltage(10)
                scope.stop()
                
                # sleep(2)

                # input(">> Adjust cursor. Press ENTER to capture waveform... ")

                # filename = f"DALI {start_bit}-{end_bit} bit, {vin}Vac, {led}V.png"
                filename = f"Analog Dim {dim}-10 V, {vin}Vac, {led}V.png"
                screenshot(filename, path)

                discharge_output(1)








def main():
    global waveform_counter
    discharge_output(1)
    scope_settings()
    operation()
        
if __name__ == "__main__":
    headers(test)
    main()
    footers(waveform_counter)





           