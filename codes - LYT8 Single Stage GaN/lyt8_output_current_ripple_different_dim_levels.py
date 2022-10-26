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

"""COMMS"""
ac_source_address = 5
source_power_meter_address = 30 
load_power_meter_address = 2
eload_address = 8
scope_address = "10.125.10.129"
sig_gen_address = '11'
dimming_power_meter_address = 21
eload_channel = 5
dc_source_address = '4'

"""USER INPUT"""

vin_list = [90,115,130,180,230,265,277,300]
vout = 50
iout_nom = 1.5
led_list = [50,36,24]
dim_list = [1,2,3,4,5,6,7,8,9,10]
iout_channel = 1

test = "Iout Ripple (at Different Dim)"
operation = "CC"
condition = input(">> Condition: ")
excel_name = f"{condition}"
waveforms_folder = f'C:/Users/{username}/Desktop/APPS EVAL/Single Stage GAN/Test Data ({operation})/{date}/{test}'
path = path_maker(f'{waveforms_folder}')

"""SCOPE SETTINGS"""
trig_channel = 1
trig_level = 1.4
trig_edge = 'POS'

time_position = 50
time_scale = 0.002

"""
MEASUREMENT SETTINGS OPTIONS: "MAX,MIN,RMS,MEAN,PDELta"
"""

ch1_enable = 'ON'
ch2_enable = 'OFF'
ch3_enable = 'OFF'
ch4_enable = 'OFF'

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
ch4_scale = 10
ch4_position = -4
ch4_bw = 20
ch4_rel_x_position = 80
ch4_label = "VOUT"
ch4_measure = "MAX"
ch4_color = "PINK"
ch4_coupling = "DCLimit"


"""DO NOT EDIT BELOW THIS LINE"""
##################################################################################

"""EQUIPMENT INITIALIZE"""
ac = ACSource(ac_source_address)
pms = PowerMeter(source_power_meter_address)
pml = PowerMeter(load_power_meter_address)
eload = ElectronicLoad(eload_address)
scope = Oscilloscope(scope_address)
led_ctrl = LEDControl()
dc = Keithley_DC_2230G(dc_source_address)

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
                            color=ch1_color, rel_x_position=ch1_rel_x_position, bandwidth=ch1_bw, coupling=ch1_coupling, offset=ch1_offset)
        
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

    scope.run_single()
    soak(3)
    trigger_status = scope.trigger_status()
    while (trigger_status != 1):
        # decrease trigger level below to get the maximum trigger possible
        trigger_level -= 1*trigger_delta
        scope.edge_trigger(channel, trigger_level, 'POS')

        # check trigger status
        scope.run_single()
        soak(3)
        trigger_status = scope.trigger_status()


def set_dim_voltage(dim):
    dc.set_volt_curr(channel ='CH1', voltage = dim, current = 1.0)
    dc.channel_state(channel ='CH1', state = 'ON')
    soak(2)


def collect_data(led, vin, dim):
    
    """COLLECT DATA"""
    labels, values = scope.get_measure(iout_channel)
    iout_max = float(values[0])
    iout_min = float(values[1])
    iout_pkpk = float(values[2])
    iout_mean = float(values[3])
    iout_rms = float(values[4])
    iout_ripple = 100*iout_pkpk/iout_mean
    iout_flicker = 100*iout_pkpk/(2*iout_mean)
    output_list = [led, vin, dim, iout_max, iout_min, iout_pkpk, iout_mean, iout_rms, iout_ripple, iout_flicker]
    output_list = [round(x, 4) for x in output_list]
    
    return output_list

def operation():

    ################### COPY PASTE THIS CODE IN HEADER OF MAIN ###################
    df_header_list = ['LED (V)', 'Vin (VAC)', 'Dim (V)', 'Iout_max (A)', 'Iout_min (A)', 'Iout_pk-pk (A)', 'Iout_mean (A)', 'Iout_rms (A)', '%Ripple', 'Flicker(%)'] 
    df = create_header_list(df_header_list)
    ##############################################################################

    for led in led_list:

        led_ctrl.voltage(led)

        for vin in vin_list:

            for dim in dim_list:

                set_dim_voltage(dim)

                scope.channel_offset(iout_channel, iout_nom)

                ac.voltage = vin
                ac.turn_on()
                sleep(5)

                find_trigger(channel=iout_channel, trigger_delta=0.001)
                
                scope.run_single()
                sleep(3)
                scope.stop()

                discharge_output()

                filename = f"Iout Ripple - {led}V, {vin}Vac, {dim}V.png"
                screenshot(filename, path)

                output_list = collect_data(led, vin, dim)
                export_to_excel(df, waveforms_folder, output_list, excel_name, sheet_name=f"{led}V", anchor="B2")

        print(f"\n\nFinal Data: ")
        print(df)
    
    


def main():
    global waveform_counter
    discharge_output()
    scope_settings()
    operation()
    discharge_output()
        
if __name__ == "__main__":
    headers(test)
    main()
    footers(waveform_counter)