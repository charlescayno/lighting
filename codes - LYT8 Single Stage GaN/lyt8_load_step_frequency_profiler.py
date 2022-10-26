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

test = "LOAD STEP FREQUENCY PROFILER"
operation = "CV"
condition = input(">> Type test condition: ")
excel_name = f"{condition}"
waveforms_folder = f'C:/Users/{username}/Desktop/APPS EVAL/Single Stage GAN/Test Data ({operation})/{date}/{test}/{condition}/'
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
# vin_list = [90,115,130,180,230,265,277,300]
vin_list = [230]
iout_nom = 1.5
vout = 50
iout_list = [iout_nom]

percent_list = [0,100]

cr_load = vout/iout_nom

"""SCOPE CHANNEL"""
iout_channel = 1
# vin_channel = 2
ids_channel = 3
vout_channel = 4

"""SCOPE SETTINGS"""
trig_channel = 1
trig_level = 0.4
trig_edge = 'POS'

time_position = 10
time_scale = 0.02

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
ch1_scale = 0.2
ch1_position = -4
ch1_bw = 20
ch1_rel_x_position = 65.9 #from left
ch1_label = "LYT8 IOUT" 
ch1_measure = "MAX,MIN,MEAN"
ch1_color = "GREEN"
ch1_coupling = "DCLimit"

"""CHANNEL 2"""
ch2_scale = 2
ch2_position = -4
ch2_bw = 20
ch2_rel_x_position = 75.5
ch2_label = "VR"
ch2_measure = "MAX,MIN,MEAN"
ch2_color = "YELLOW"
ch2_coupling = "DCLimit"
ch2_offset = 0

"""CHANNEL 3"""
ch3_scale = 1
ch3_position = -4
ch3_bw = 500
ch3_rel_x_position = 83
ch3_label = "PRI_IDS"
ch3_measure = "MAX,MIN,MEAN"
ch3_color = "LIGHT_BLUE"
ch3_coupling = "DCLimit"

"""CHANNEL 4"""
ch4_scale = 5
ch4_position = -4
ch4_bw = 20
ch4_rel_x_position = 97.2
ch4_label = "LYT8 VOUT"
ch4_measure = "MAX,MIN"
ch4_color = "PINK"
ch4_coupling = "DCLimit"
ch4_offset = 20

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

def screenshot_looper(filename, path):
        
    iter = 0
    while '' == input(">> Press ENTER to continue capture. To stop press other keys to stop: "):

        fname = filename.split('.png')[0] + f'_{iter}' + '.png'
        iter += 1
        screenshot(fname, path)

def collect_data(input_voltage, start_percent, end_percent):


    ################ SUBJECT TO CHANGE ################
    input_step = input_voltage

    labels, values = scope.get_measure(vout_channel)
    vout_max = round(float(values[0]),2)
    vout_min = round(float(values[1]),2)
    vout_mean = round(float(values[2]),2)

    labels, values = scope.get_measure(iout_channel)
    iout_max = round(float(values[0]),2)
    iout_min = round(float(values[1]),2)

    labels, values = scope.get_measure(ids_channel)
    ids_max = round(float(values[0]),2)

    vout_overshoot = round(abs(100*(vout_max-vout)/vout),2)
    vout_undershoot = round(abs(100*(vout-vout_min)/vout),2)

    if vout_overshoot > 8: vout_overshoot_result = 'FAIL'
    else: vout_overshoot_result = 'PASS'
    if vout_undershoot > 8: vout_undershoot_result = 'FAIL'
    else: vout_undershoot_result = 'PASS'

    start = cr_load/(start_percent/100) if start_percent != 0 else 0
    end = cr_load/(end_percent/100) if end_percent != 0 else 0

    start = round(start,2)
    end = round(end,2)

   
    output_list = [start, end, start_percent, end_percent, input_voltage,
                    vout_max, vout_overshoot, vout_overshoot_result,
                    vout_min, vout_undershoot, vout_undershoot_result,
                    vout_mean,
                    ids_max,
                    iout_max, 
                    iout_min]


    # output_list = [start_percent, end_percent, input_voltage,
    #                 vout_max, vout_overshoot, vout_overshoot_result,
    #                 vout_min, vout_undershoot, vout_undershoot_result,
    #                 vout_mean,
    #                 ids_max,
    #                 iout_max, 
    #                 iout_min]
    return output_list

def loadstep_cr(df, input, cr_load, start_percent, end_percent):

    start = cr_load/(start_percent/100) if start_percent != 0 else 0
    end = cr_load/(end_percent/100) if end_percent != 0 else 0
    
    start_iout = iout_nom*start_percent/100
    end_iout = iout_nom*end_percent/100
    trigger_level = (start_iout+end_iout)/2
    if start_iout < end_iout: trigger_slope = 'POS'
    else: trigger_slope = 'NEG'
    scope.edge_trigger(trig_channel, trigger_level, trigger_slope)
    
    
    eload.channel[eload_channel].von = 0
    eload.channel[eload_channel].cr = start
    eload.channel[eload_channel].turn_on() if start_percent != 0 else eload.channel[eload_channel].turn_off()

    ac.voltage = input
    ac.turn_on()

    soak(5)

    scope.run_single()

    soak(3)

    eload.channel[eload_channel].von = 0
    eload.channel[eload_channel].cr = end
    eload.channel[eload_channel].turn_on() if end_percent != 0 else eload.channel[eload_channel].turn_off()

    soak(2)

    discharge_output(times=2)

    start = round(start,2)
    end = round(end,2)

    filename = f"{input}Vac, {start}-{end} Ohms ({start_percent}-{end_percent} CR load).png"
    print(filename)
    screenshot(filename, path)
    screenshot_looper(filename, path)
    
    output_list = collect_data(input, start_percent, end_percent)
    export_to_excel(df, output_list, excel_name, sheet_name="Sheet1", anchor="B2")
    print()

def loadstep_cc(df, input, start_percent, end_percent):

    # start = cr_load/(start_percent/100) if start_percent != 0 else 0
    # end = cr_load/(end_percent/100) if end_percent != 0 else 0
    
    start_iout = iout_nom*start_percent/100
    end_iout = iout_nom*end_percent/100
    trigger_level = (start_iout+end_iout)/2
    if start_iout < end_iout: trigger_slope = 'POS'
    else: trigger_slope = 'NEG'
    scope.edge_trigger(trig_channel, trigger_level, trigger_slope)
    
    
    eload.channel[eload_channel].von = 0
    eload.channel[eload_channel].cc = start_iout
    eload.channel[eload_channel].turn_on() if start_percent != 0 else eload.channel[eload_channel].turn_off()

    ac.voltage = input
    ac.turn_on()

    soak(5)

    scope.run_single()

    soak(3)


    eload.channel[eload_channel].von = 0
    eload.channel[eload_channel].cc = end_iout
    eload.channel[eload_channel].turn_on() if end_percent != 0 else eload.channel[eload_channel].turn_off()

    soak(2)

    discharge_output(times=2)


    filename = f"{input}Vac, {start_iout}-{end_iout} A ({start_percent}-{end_percent} CC load).png"
    screenshot(filename, path)
    output_list = collect_data(input, start_percent, end_percent)
    export_to_excel(df, output_list, excel_name, sheet_name="Sheet1", anchor="B2")
    print()

def operation():

    ################### COPY PASTE THIS CODE IN HEADER OF MAIN ###################
    df_header_list = ['Start (Ohms)', 'End (Ohms)', 'Start (%)', 'End (%)', 'Input Voltage (VAC)', 
                        'Vout_max (V)', 'Overshoot (%)', 'PASS/FAIL',
                        'Vout_min (V)', 'Undershoot (%)', 'PASS/FAIL',
                        'Vout_mean (V)', 
                        'Ids_max (A)',
                        'Iout_max (A)',
                        'Iout_min (A)']
    df = create_header_list(df_header_list)
    ##############################################################################

    # ################### COPY PASTE THIS CODE IN HEADER OF MAIN ###################
    # df_header_list = ['Start (%)', 'End (%)', 'Input Voltage (VAC)', 
    #                     'Vout_max (V)', 'Overshoot (%)', 'PASS/FAIL',
    #                     'Vout_min (V)', 'Undershoot (%)', 'PASS/FAIL',
    #                     'Vout_mean (V)', 
    #                     'Ids_max (A)',
    #                     'Iout_max (A)',
    #                     'Iout_min (A)']
    # df = create_header_list(df_header_list)
    # ##############################################################################


    scope.cursor(channel=vout_channel, cursor_set=1, X1=1, X2=1, Y1=50*0.92, Y2=50*1.08, type='HOR')

    percent_list.sort()
    for i in range(len(percent_list)):
        for percent in percent_list:
            if percent_list[i] < percent:
                for vin in vin_list:
                    start = percent_list[i]
                    end = percent
                    loadstep_cr(df, input=vin, cr_load=cr_load, start_percent=start, end_percent=end)
                    # loadstep_cc(df, input=vin, start_percent=start, end_percent=end)


    # percent_list.sort(reverse=True)
    # for i in range(len(percent_list)):
    #     for percent in percent_list:
    #         if percent_list[i] > percent:
    #             for vin in vin_list:
    #                 start = percent_list[i]
    #                 end = percent
    #                 loadstep_cr(df, input=vin, cr_load=cr_load, start_percent=start, end_percent=end)


def main():
    global waveform_counter
    discharge_output(times=1)
    scope_settings()
    operation()
        
if __name__ == "__main__":
    headers(test)
    main()
    footers(waveform_counter)

































# """CHANNEL 1"""
# ch1_scale = 0.4
# ch1_position = -4
# ch1_bw = 20
# ch1_rel_x_position = 20
# ch1_label = "IOUT"
# ch1_measure = "MAX,MIN"
# ch1_color = "GREEN"
# ch1_coupling = "DCLimit"

# """CHANNEL 2"""
# ch2_scale = 300
# ch2_position = 2
# ch2_bw = 20
# ch2_rel_x_position = 40
# ch2_label = "VIN"
# ch2_measure = "MAX,RMS"
# ch2_color = "YELLOW"
# ch2_coupling = "DCLimit"

# """CHANNEL 3"""
# ch3_scale = 2
# ch3_position = -4
# ch3_bw = 500
# ch3_rel_x_position = 60
# ch3_label = "PRI_IDS"
# ch3_measure = "MAX,MIN"
# ch3_color = "LIGHT_BLUE"
# ch3_coupling = "DCLimit"

# """CHANNEL 4"""
# ch4_scale = 5
# ch4_position = -5
# ch4_bw = 20
# ch4_rel_x_position = 80
# ch4_label = "VOUT"
# ch4_measure = "MAX,MIN,MEAN"
# ch4_color = "PINK"
# ch4_coupling = "DCLimit"
# ch4_offset = 10

