## 10142022
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
from powi.equipment import path_maker, remove_file, start_timer, end_timer

import getpass
username = getpass.getuser().lower()

from datetime import datetime
now = datetime.now()
date = now.strftime('%m%d')
##################################################################################


test = f"Output Voltage Ripple"
condition = "1007 rework"
project = "APPS EVAL/Single Stage GAN/Test Data (CV)"
excel_name = f"{test}_{condition}"
waveforms_folder = f'C:/Users/{username}/Desktop/{project}/{date}/{test}/'
path = path_maker(f'{waveforms_folder}')


"""COMMS"""
ac_source_address = 5
source_power_meter_address = 30 
load_power_meter_address = 2
eload_address = 8
scope_address = "10.125.10.101"
sig_gen_address = '11'
dimming_power_meter_address = 21
eload_channel = 1

"""USER INPUT"""



vin_list = [90,100,110,115,120,130,132,180,200,230,265,277,300]
vout = 50
iout_nom = 1.5


percent_list = [100,75,50,25,10,0]
percent_list = [0]

iout_list = [(percent/100) * iout_nom for percent in percent_list]
cr_load = vout/iout_nom
cr_list = [(cr_load/(percent/100) if percent != 0 else 0) for percent in percent_list]

iout_channel = 1
vin_channel = 2
ids_channel = 3
vout_channel = 4




"""SCOPE SETTINGS"""
trig_channel = 4
trig_level = vout
trig_edge = 'POS'

time_position = 50
time_scale = 0.1

"""
MEASUREMENT SETTINGS OPTIONS: "MAX,MIN,RMS,MEAN,PDELta"
"""

ch1_enable = 'OFF'
ch2_enable = 'OFF'
ch3_enable = 'OFF'
ch4_enable = 'ON'

"""CHANNEL 1"""
ch1_scale = 0.015
ch1_position = 0
ch1_bw = 20
ch1_rel_x_position = 20
ch1_label = "OUTPUT CURRENT"
ch1_measure = "MAX,MIN,RMS,MEAN,PDELta"
ch1_color = "GREEN"
ch1_coupling = "DCLimit"
ch1_offset = iout_nom

"""CHANNEL 2"""
ch2_scale = 100
ch2_position = -4
ch2_bw = 500
ch2_rel_x_position = 40
ch2_label = "PRI_VDS"
ch2_measure = "MAX"
ch2_color = "YELLOW"
ch2_coupling = "DCLimit"

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
ch4_scale = 1
ch4_position = 0
ch4_bw = 20
ch4_rel_x_position = 65.5
ch4_label = "OUTPUT VOLTAGE RIPPLE"
ch4_measure = "MAX,MIN,RMS,MEAN,PDELta"
ch4_color = "PINK"
ch4_coupling = "DCLimit"
ch4_offset = 54


"""DO NOT EDIT BELOW THIS LINE"""
##################################################################################

"""EQUIPMENT INITIALIZE"""
ac = ACSource(ac_source_address)
pms = PowerMeter(source_power_meter_address)
pml = PowerMeter(load_power_meter_address)
eload = ElectronicLoad(eload_address)
scope = Oscilloscope(scope_address)
# led_ctrl = LEDControl()

"""GENERIC FUNCTIONS"""

def discharge_output():
    ac.turn_off()
    for i in range(1):
        for i in range(1,9):
            eload.channel[i].cc = 1
            eload.channel[i].turn_on()
            eload.channel[i].short_on()
        sleep(2)

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
                            color=ch4_color, rel_x_position=ch4_rel_x_position, bandwidth=ch4_bw, coupling=ch4_coupling, offset=ch4_offset)
    
    if ch1_enable != 'OFF': scope.measure(1, ch1_measure)
    if ch2_enable != 'OFF': scope.measure(2, ch2_measure)
    if ch3_enable != 'OFF': scope.measure(3, ch3_measure)
    if ch4_enable != 'OFF': scope.measure(4, ch4_measure)

    scope.record_length(50E6)
    scope.time_position(time_position)
    scope.time_scale(time_scale)

    # scope.remove_zoom()
    # scope.add_zoom(rel_pos=50, rel_scale=0.5)
    
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

def find_trigger(channel, trigger_delta):

    scope.edge_trigger(channel, 0, 'POS')

    # finding trigger level
    scope.run_single()
    scope.trigger_mode('AUTO')
    soak(5)
    scope.trigger_mode('NORM')
    

    # get initial peak-to-peak measurement value
    labels, values = scope.get_measure(channel)
    max_value = float(values[0])
    max_value = float(f"{max_value:.4f}")

    # set max_value as initial trigger level
    trigger_level = max_value
    scope.edge_trigger(channel, trigger_level, 'POS')

    # check if it triggered within 3 seconds
    scope.run_single()
    soak(3)
    trigger_status = scope.trigger_status()

    # increase trigger level until it reaches the maximum trigger level
    while (trigger_status == 1):
        trigger_level += trigger_delta
        scope.edge_trigger(channel, trigger_level, 'POS')

        # check trigger status
        scope.run_single()
        soak(3)
        trigger_status = scope.trigger_status()

    # decrease trigger level below to get the maximum trigger possible
    trigger_level -= 1*trigger_delta
    scope.edge_trigger(channel, trigger_level, 'POS')
    trigger_status = scope.trigger_status()
    while (trigger_status != 1):
        # decrease trigger level below to get the maximum trigger possible
        trigger_level -= 1*trigger_delta
        scope.edge_trigger(channel, trigger_level, 'POS')

        # check trigger status
        scope.run_single()
        soak(3)
        trigger_status = scope.trigger_status()


def collect_data(percent, v_input):
    
    """COLLECT DATA"""
    labels, values = scope.get_measure(vout_channel)
    vout_max = float(values[0])
    vout_min = float(values[1])
    vout_pkpk = float(values[2])
    vout_mean = float(values[3])
    vout_rms = float(values[4])
    vout_ripple = 100*vout_pkpk/vout_mean
    output_list = [percent, v_input, vout_max, vout_min, vout_pkpk, vout_mean, vout_rms, vout_ripple]
    
    return output_list

def operation():

    ################### COPY PASTE THIS CODE IN HEADER OF MAIN ###################
    df_header_list = ['CR (Ohms)', 'Vin (Vac)', 'Vmax (V)', 'Vmin (V)', 'Vpk-pk (V)','Vmean (V)', 'Vrms (V)','%Peak to Peak (%)']
    df = create_header_list(df_header_list)
    ##############################################################################

    scope.cursor(channel=4, cursor_set=4, X1=1, X2=1, Y1=vout*0.92, Y2=vout*1.08, type='HOR')

    for cr in cr_list:
    
        for vin in vin_list:

            if cr == 0:
                eload.channel[eload_channel].turn_off()
            else:
                eload.channel[eload_channel].cr = cr
                eload.channel[eload_channel].turn_on()


            percent = int(cr_load*100/cr) if cr != 0 else 0


            ac.voltage = vin
            ac.turn_on()
            sleep(5)

            
            sleep(3)
            find_trigger(channel=vout_channel, trigger_delta=0.05)
                       

            discharge_output()

            crr = round(cr, 2)
            
            filename = f"Vout Ripple - {vin}Vac, {crr}Ohms ({percent}percent CR load).png"
            screenshot(filename, path)

            output_list = collect_data(percent, vin)
            export_to_excel(df, waveforms_folder, output_list, excel_name, sheet_name=f"{percent}%", anchor="B2")
        
        ################### COPY PASTE THIS CODE IN HEADER OF MAIN ###################
        df_header_list = ['Load Percent (%)', 'Vin (Vac)', 'Vmax (V)', 'Vmin (V)', 'Vpk-pk (V)','Vmean (V)', 'Vrms (V)','%Peak to Peak (%)']
        df = create_header_list(df_header_list)
        ##############################################################################
    
    


def main():
    global waveform_counter
    discharge_output()
    scope_settings()
    operation()
        
if __name__ == "__main__":
    headers(test)
    main()
    footers(waveform_counter)