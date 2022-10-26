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

load_type = input(">> Load Type (cc/cr/nl): ")
project = "APPS EVAL/Single Stage GAN/Test Data (CV)"
test = f"Startup Waveforms"
operation = "CV"
condition = "1007 rework"
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

iout_channel = 1
vr_channel = 2
ids_channel = 3
vout_channel = 4

"""USER INPUT"""
von_list = [50,40,30,25,20,0]
vin_list = [90,115,130,180,230,265,277,300]
percent_list = [100,75,50,25,10,0]

vout = 50
iout_nom = 1.5

cr_load = vout/iout_nom

iout_list = [iout_nom*percent/100 for percent in percent_list]
cr_list = [(cr_load/(percent/100) if percent != 0 else 0) for percent in percent_list]

zoom_enable = True

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
ch1_label = "IOUT"
ch1_measure = "MAX,MIN,MEAN"
ch1_color = "GREEN"
ch1_coupling = "DCLimit"
ch1_offset = 0

"""CHANNEL 2"""
ch2_scale = 2
ch2_position = -4
ch2_bw = 20
ch2_rel_x_position = 40
ch2_label = "VR"
ch2_measure = "MAX,MIN,MEAN"
ch2_color = "YELLOW"
ch2_coupling = "DCLimit"
ch2_offset = 0

"""CHANNEL 3"""
ch3_scale = 2
ch3_position = -4
ch3_bw = 500
ch3_rel_x_position = 60
ch3_label = "PRI_IDS"
ch3_measure = "MAX,MIN,MEAN"
ch3_color = "LIGHT_BLUE"
ch3_coupling = "DCLimit"
ch3_offset = 0

"""CHANNEL 4"""
ch4_scale = 10
ch4_position = -4
ch4_bw = 20
ch4_rel_x_position = 80
ch4_label = "VOUT"
ch4_measure = "MAX,MIN,MEAN"
ch4_color = "PINK"
ch4_coupling = "DCLimit"
ch4_offset = 0


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
                            color=ch1_color, rel_x_position=ch1_rel_x_position, bandwidth=ch1_bw, coupling=ch1_coupling, offset=ch1_offset)
        
    scope.channel_settings(state=ch2_enable, channel=2, scale=ch2_scale, position=ch2_position, label=ch2_label,
                            color=ch2_color, rel_x_position=ch2_rel_x_position, bandwidth=ch2_bw, coupling=ch2_coupling, offset=ch2_offset)
    
    scope.channel_settings(state=ch3_enable, channel=3, scale=ch3_scale, position=ch3_position, label=ch3_label,
                            color=ch3_color, rel_x_position=ch3_rel_x_position, bandwidth=ch3_bw, coupling=ch3_coupling, offset=ch3_offset)
    
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

def collect_data(input_voltage, load, von):

    ################ SUBJECT TO CHANGE ################
    input_step = input_voltage

    labels, values = scope.get_measure(vout_channel)
    vout_max = round(float(values[0]),2)
    vout_min = round(float(values[1]),2)
    vout_mean = round(float(values[2]),2)

    labels, values = scope.get_measure(iout_channel)
    try:
        iout_max = round(float(values[0]),2)
        iout_min = round(float(values[1]),2)
        iout_mean = round(float(values[2]),2)
    except:
        iout_max = 0
        iout_min = 0
        iout_mean = 0

    labels, values = scope.get_measure(ids_channel)
    ids_max = round(float(values[0]),2)

    vout_overshoot = round(abs(100*(vout_max-vout)/vout),2)
    vout_undershoot = round(abs(100*(vout-vout_min)/vout),2)

    if vout_overshoot > 13.4: vout_overshoot_result = 'FAIL'
    else: vout_overshoot_result = 'PASS'
    if vout_undershoot > 13.4: vout_undershoot_result = 'FAIL'
    else: vout_undershoot_result = 'PASS'


    if load_type == 'cc':
        load_percent = int(load*100/iout_nom)
        load_value = round(float(f"{load:.2f}"), 2)

    if load_type == 'cr':
        try:
            load_percent = int(cr_load*100/load)
            load_value = round(float(f"{load:.2f}"), 2)
        except:
            load_percent = 0
            load_value = 0

    output_list = [von, input_step, load_percent, load_value,
                    vout_max, vout_overshoot, vout_overshoot_result,
                    vout_min, vout_undershoot, vout_undershoot_result,
                    iout_max, ids_max, iout_mean, vout_mean]
    
    return output_list

def operation():

    ################### COPY PASTE THIS CODE IN HEADER OF MAIN ###################
    df_header_list = ['Von (V)', 'Input Voltage (VAC)', 'Load (%)', 'Load',
                        'Vout_max (V)', 'Overshoot (%)', 'PASS/FAIL',
                        'Vout_min (V)', 'Undershoot (%)', 'PASS/FAIL',
                        'Iout_max (A)', 'Ids_max (A)',
                        'Iout_mean (A)', 'Vout_mean (V)']
    df = create_header_list(df_header_list)
    ##############################################################################

    pms.current_range(2)
    pml.current_range(2)

    estimated_time = (len(von_list)*len(iout_list)*len(vin_list)*(5+5))/60
    print(f"Estimated time: {estimated_time} mins.")


    if load_type == 'cc':
        for von in von_list:
            for iout in iout_list:
                for vin in vin_list:

                    eload.channel[eload_channel].von = von

                    if iout == 0:
                        filename = f"Startup - {vin}Vac, No Load.png"
                        eload.channel[eload_channel].turn_off()
                    else:
                        percent = int(iout*100/iout_nom)
                        filename = f"Startup - {vin}Vac, {iout}A ({percent} CC load), Von = {von} V.png"
                        eload.channel[eload_channel].von = von
                        eload.channel[eload_channel].cc = iout
                        eload.channel[eload_channel].turn_on()

                    scope.run_single()
                    sleep(2)

                    ac.voltage = vin
                    ac.turn_on()
                    sleep(2)
                    
                    discharge_output(times=1)

                    screenshot(filename, path)
                    output_list = collect_data(vin, iout, von)
                    export_to_excel(df, waveforms_folder, output_list, excel_name, sheet_name=f"cc load", anchor="B2")
                    print(df)

        print(f"\n\nFinal Data: \n")
        print(df)
    
    if load_type == 'cr':
        for von in von_list:
            eload.channel[eload_channel].von = von
            for cr in cr_list:
                for vin in vin_list:
                    if cr == 0:
                        filename = f"Startup - {vin}Vac, No Load.png"
                        eload.channel[eload_channel].turn_off()
                    else:
                        percent = int(cr_load*100/cr) if cr != 0 else 0
                        crr = round(cr, 2)
                        filename = f"Startup - {vin}Vac, {crr}Ohms ({percent} CR load), Von = {von} V.png"
                        eload.channel[eload_channel].cr = cr
                        eload.channel[eload_channel].turn_on()

                    scope.run_single()
                    sleep(2)

                    ac.voltage = vin
                    ac.turn_on()
                    sleep(2)
                    
                    discharge_output(times=1)

                    screenshot(filename, path)
                    output_list = collect_data(vin, cr, von)
                    export_to_excel(df, waveforms_folder, output_list, excel_name, sheet_name=f"cr load", anchor="B2")
                    print(df)
        
        print(f"\n\nFinal Data: \n")
        print(df)
        

def main():
    global waveform_counter
    discharge_output(times=1)
    scope_settings()
    operation()
        
if __name__ == "__main__":
    headers(test)
    main()
    footers(waveform_counter)