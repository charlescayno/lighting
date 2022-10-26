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
from powi.equipment import path_maker, remove_file

import getpass
username = getpass.getuser().lower()

from datetime import datetime
now = datetime.now()
date = now.strftime('%m%d')
##################################################################################

load_type = input(">> Enter load type (cc/cr/dcdc): ")

test = f"Line Transient"
condition = "1007 rework"
project = "APPS EVAL/Single Stage GAN/Test Data (CV)"
excel_name = f"{test}_{condition}"
waveforms_folder = f'C:/Users/{username}/Desktop/{project}/{date}/{test}/{load_type}'
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
percent_list = [100,75,50,25,10,0]
# percent_list = [100]

if load_type != 'dcdc':
    vout = 50
    iout_nom = 1.5
    iout_list = [(percent/100) * iout_nom for percent in percent_list]
    cr_load = vout/iout_nom
    cr_list = [(cr_load/(percent/100) if percent != 0 else 0) for percent in percent_list]
else:
    vout = 50
    dcdc_vout = 38
    iout_nom = 1.05
    iout_list = [(percent/100) * iout_nom for percent in percent_list]
    cr_load = dcdc_vout/iout_nom
    cr_list = [(cr_load/(percent/100) if percent != 0 else 0) for percent in percent_list]




"""SCOPE SETTINGS"""
trig_channel = 4
trig_level = 40
trig_edge = 'NEG'

time_position = 20
time_scale = 1

"""ZOOM SETTINGS"""
zoom_enable = False
zoom_pos = 50
zoom_rel_scale = 1

"""
MEASUREMENT SETTINGS OPTIONS: "MAX,MIN,RMS,MEAN,PDELta"
"""

if load_type != 'dcdc':

    iout_channel = 1
    vin_channel = 2
    ids_channel = 3
    vout_channel = 4

    input("CH1 - IOUT, CH2 - VIN, CH3 - IDS, CH4 - VOUT")

    ch1_enable = 'ON'
    ch2_enable = 'ON'
    ch3_enable = 'ON'
    ch4_enable = 'ON'

    """CHANNEL 1"""
    ch1_scale = 0.3
    ch1_position = -4
    ch1_bw = 20
    ch1_rel_x_position = 40
    ch1_label = "IOUT"
    ch1_measure = "MAX,MIN"
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
    ch3_measure = "MAX,MIN"
    ch3_color = "LIGHT_BLUE"
    ch3_coupling = "DCLimit"

    """CHANNEL 4"""
    ch4_scale = 10
    ch4_position = -4
    ch4_bw = 20
    ch4_rel_x_position = 80
    ch4_label = "VOUT"
    ch4_measure = "MAX,MIN"
    ch4_color = "PINK"
    ch4_coupling = "DCLimit"
    ch4_offset = 0

else:

    input("CH1 - VIN, CH2 - DCDC VOUT, CH3 - IDS, CH4 - LYT8 VOUT")

    vin_channel = 1
    dcdc_vout_channel = 2
    ids_channel = 3
    vout_channel = 4

    ch1_enable = 'ON'
    ch2_enable = 'ON'
    ch3_enable = 'ON'
    ch4_enable = 'ON'

    """CHANNEL 1"""
    ch1_scale = 150
    ch1_position = 0
    ch1_bw = 20
    ch1_rel_x_position = 40
    ch1_label = "VIN"
    ch1_measure = "MAX,RMS"
    ch1_color = "YELLOW"
    ch1_coupling = "AC"

    """CHANNEL 2"""
    ch2_scale = 10
    ch2_position = -4
    ch2_bw = 20
    ch2_rel_x_position = 40
    ch2_label = "DCDC VOUT"
    ch2_measure = "MAX,MIN"
    ch2_color = "GREEN"
    ch2_coupling = "DCLimit"

    """CHANNEL 3"""
    ch3_scale = 2
    ch3_position = -4
    ch3_bw = 500
    ch3_rel_x_position = 60
    ch3_label = "PRI_IDS"
    ch3_measure = "MAX,MIN"
    ch3_color = "LIGHT_BLUE"
    ch3_coupling = "DCLimit"

    """CHANNEL 4"""
    ch4_scale = 10
    ch4_position = -4
    ch4_bw = 20
    ch4_rel_x_position = 80
    ch4_label = "VOUT"
    ch4_measure = "MAX,MIN"
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

def fixvoltage(voltage, soak):
    ac.voltage = voltage
    ac.frequency = ac.set_freq(voltage)
    ac.turn_on()
    sleep(soak)

def collect_data(start_voltage, end_voltage, load):


    ################ SUBJECT TO CHANGE ################
    input_step = f"{start_voltage}-{end_voltage} Vac"

    

    labels, values = scope.get_measure(ids_channel)
    ids_max = round(float(values[0]),2)

    

    if load_type == 'cc':

        labels, values = scope.get_measure(vout_channel)
        vout_max = round(float(values[0]),2)
        vout_min = round(float(values[1]),2)


        vout_overshoot = round(abs(100*(vout_max-vout)/vout),2)
        vout_undershoot = round(abs(100*(vout-vout_min)/vout),2)

        vout_overshoot_result = 'FAIL' if vout_overshoot > 13 else 'PASS'
        vout_undershoot_result = 'FAIL' if vout_undershoot > 13 else 'PASS'


        labels, values = scope.get_measure(iout_channel)
        iout_max = round(float(values[0]),2)

        load_percent = int(load*100/iout_nom)
        load_value = float(f"{load:.2f}")

        output_list = [input_step, load_percent, load_value,
                        vout_max, vout_overshoot, vout_overshoot_result,
                        vout_min, vout_undershoot, vout_undershoot_result,
                        iout_max, ids_max]

    if load_type == 'cr':

        labels, values = scope.get_measure(vout_channel)
        vout_max = round(float(values[0]),2)
        vout_min = round(float(values[1]),2)


        vout_overshoot = round(abs(100*(vout_max-vout)/vout),2)
        vout_undershoot = round(abs(100*(vout-vout_min)/vout),2)

        vout_overshoot_result = 'FAIL' if vout_overshoot > 13 else 'PASS'
        vout_undershoot_result = 'FAIL' if vout_undershoot > 13 else 'PASS'


    
        labels, values = scope.get_measure(iout_channel)
        iout_max = round(float(values[0]),2)

        try:
            load_percent = int(cr_load*100/load)
            load_value = float(f"{load:.2f}")
        except:
            load_percent = 0
            load_value = 0
    
        output_list = [input_step, load_percent, load_value,
                        vout_max, vout_overshoot, vout_overshoot_result,
                        vout_min, vout_undershoot, vout_undershoot_result,
                        iout_max, ids_max]

    if load_type == 'dcdc':

        labels, values = scope.get_measure(vout_channel)
        vout_max = round(float(values[0]),2)
        vout_min = round(float(values[1]),2)

        labels, values = scope.get_measure(dcdc_vout_channel)
        dcdc_vout_max = round(float(values[0]),2)
        dcdc_vout_min = round(float(values[1]),2)


        vout_overshoot = round(abs(100*(vout_max-vout)/vout),2)
        vout_undershoot = round(abs(100*(vout-vout_min)/vout),2)

        vout_overshoot_result = 'FAIL' if vout_overshoot > 13 else 'PASS'
        vout_undershoot_result = 'FAIL' if vout_undershoot > 13 else 'PASS'


        dcdc_vout_overshoot = round(abs(100*(vout_max-dcdc_vout)/dcdc_vout),2)
        dcdc_vout_undershoot = round(abs(100*(dcdc_vout-vout_min)/dcdc_vout),2)

        dcdc_vout_overshoot_result = 'FAIL' if dcdc_vout_overshoot > 13 else 'PASS'
        dcdc_vout_undershoot_result = 'FAIL' if dcdc_vout_undershoot > 13 else 'PASS'




        load_percent = int(load*100/iout_nom)
        load_value = float(f"{load:.2f}")
        iout_max = 'NA'


        output_list = [input_step, load_percent, load_value,
                    vout_max, vout_overshoot, vout_overshoot_result,
                    vout_min, vout_undershoot, vout_undershoot_result,
                    dcdc_vout_max, dcdc_vout_overshoot, dcdc_vout_overshoot_result,
                    dcdc_vout_min, dcdc_vout_undershoot, dcdc_vout_undershoot_result,
                    iout_max, ids_max]

    return output_list


def line_transient_test(df, start_voltage, end_voltage, load):

    scope.time_position(20)
    scope.time_scale(0.2)

    if start_voltage < end_voltage: scope.edge_trigger(vin_channel, end_voltage, 'POS')
    else: scope.timeout_trigger(vin_channel, timeout_range='LOW', timeout_time=20E-3)


    ac.voltage = 90
    ac.turn_on()
    eload.channel[eload_channel].cc = 0.1
    eload.channel[eload_channel].turn_on()
    soak(2)
    

    if load == 0:
        eload.channel[eload_channel].turn_off()
        filename = f"{test}, {start_voltage}-{end_voltage}Vac, No Load.png"
    else:
        if load_type == 'cc' or load_type == 'dcdc':
            eload.channel[eload_channel].cc = load
            eload.channel[eload_channel].turn_on()
            percent = int(load*100/iout_nom)
            filename = f"{test}, {start_voltage}-{end_voltage}Vac, {load} A ({percent} CC load).png"
        if load_type == 'cr':
            eload.channel[eload_channel].cr = load
            eload.channel[eload_channel].turn_on()
            percent = int(cr_load*100/load)
            crr = round(load, 2)
            filename = f"{test}, {start_voltage}-{end_voltage}Vac, {crr} Ohms ({percent} CR load).png"


    scope.run_single()
    fixvoltage(start_voltage, 5)
    fixvoltage(end_voltage, 5)

    discharge_output(times=2)

    screenshot(filename, path)
    output_list = collect_data(start_voltage, end_voltage, load)
    export_to_excel(df, waveforms_folder, output_list, excel_name, sheet_name=load_type, anchor="B2")

def operation():

    if load_type != 'dcdc':
        ################### COPY PASTE THIS CODE IN HEADER OF MAIN ###################
        df_header_list = ['Input Step (VAC)', 'Load (%)', 'Load',
                            'Vout_max (V)', 'Overshoot (%)', 'PASS/FAIL',
                            'Vout_min (V)', 'Undershoot (%)', 'PASS/FAIL',
                            'Iout_max (A)', 'Ids_max (A)']
        df = create_header_list(df_header_list)
        ##############################################################################
    else:
        ################### COPY PASTE THIS CODE IN HEADER OF MAIN ###################
        df_header_list = ['Input Step (VAC)', 'Load (%)', 'Load',
                            'Vout_max (V)', 'Overshoot (%)', 'PASS/FAIL',
                            'Vout_min (V)', 'Undershoot (%)', 'PASS/FAIL',
                            'DCDC_Vout_max (V)', 'Overshoot (%)', 'PASS/FAIL',
                            'DCDC_Vout_min (V)', 'Undershoot (%)', 'PASS/FAIL',
                            'Iout_max (A)', 'Ids_max (A)']
        df = create_header_list(df_header_list)
        ##############################################################################

    pms.reset()
    scope.cursor(channel=vout_channel, cursor_set=1, X1=1, X2=1, Y1=0.92*50, Y2=1.08*50, type='HOR')

    if load_type == 'cc':
        for iout in iout_list:
            line_transient_test(df, start_voltage=100, end_voltage=230, load=iout)
            line_transient_test(df, start_voltage=100, end_voltage=277, load=iout)
            line_transient_test(df, start_voltage=90, end_voltage=300, load=iout)

            line_transient_test(df, start_voltage=230, end_voltage=100, load=iout)
            line_transient_test(df, start_voltage=277, end_voltage=100, load=iout)
            line_transient_test(df, start_voltage=300, end_voltage=90, load=iout)

    if load_type == 'cr':
        for cr in cr_list:
            line_transient_test(df, start_voltage=100, end_voltage=230, load=cr)
            line_transient_test(df, start_voltage=100, end_voltage=277, load=cr)
            line_transient_test(df, start_voltage=90, end_voltage=300, load=cr)

            line_transient_test(df, start_voltage=230, end_voltage=100, load=cr)
            line_transient_test(df, start_voltage=277, end_voltage=100, load=cr)
            line_transient_test(df, start_voltage=300, end_voltage=90, load=cr)
            
    if load_type == 'dcdc':
        scope.cursor(channel=dcdc_vout_channel, cursor_set=2, X1=1, X2=1, Y1=0.95*dcdc_vout, Y2=1.05*dcdc_vout, type='HOR')
        for iout in iout_list:
            line_transient_test(df, start_voltage=100, end_voltage=230, load=iout)
            line_transient_test(df, start_voltage=100, end_voltage=277, load=iout)
            line_transient_test(df, start_voltage=90, end_voltage=300, load=iout)

            line_transient_test(df, start_voltage=230, end_voltage=100, load=iout)
            line_transient_test(df, start_voltage=277, end_voltage=100, load=iout)
            line_transient_test(df, start_voltage=300, end_voltage=90, load=iout)


def main():
    global waveform_counter
    discharge_output(times=1)
    scope_settings()
    operation()
        
if __name__ == "__main__":
    headers(test)
    main()
    footers(waveform_counter)