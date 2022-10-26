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




test_type = int(input(">> Press 1 - Running Short LFBR, 2 - Startup Short LFBR, 3 - Running Short LFBR NL, 4 - Startup Short LFBR NL: "))
if test_type == 1:
    test = "RUNNING SHORT LFBR"
elif test_type == 2:
    test = "STARTUP SHORT LFBR"
elif test_type == 3:
    test = "RUNNING SHORT LFBR NO LOAD"
elif test_type == 4:
    test = "STARTUP SHORT LFBR NO LOAD"
else:
    input("Invalid test type!")

operation = "CV"
condition = input(">> Rework condition: ")
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
vin_list = [90,115,130,180,230,265,277,300]
vin_list = [130,180,230,265,277,300]
# vin_list = [300]

iout_nom = 1.5
vout = 50
iout_list = [iout_nom]
timeperdiv = 3

percent_list = [0,100]

cr_load = vout/iout_nom

"""SCOPE CHANNEL"""
iout_channel = 1
vr_channel = 2
ids_channel = 3
vout_channel = 4

"""SCOPE SETTINGS"""
trig_channel = 3
trig_level = 0.4
trig_edge = 'POS'

time_position = 10
time_scale = 1

"""ZOOM SETTINGS"""
zoom_enable = False
zoom_pos = 16.75
zoom_rel_scale = 15


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
ch1_rel_x_position = 10
ch1_label = "LYT8 IOUT" 
ch1_measure = "MAX,MIN,MEAN"
ch1_color = "GREEN"
ch1_coupling = "DCLimit"

"""CHANNEL 2"""
ch2_scale = 10 # with linear reg
# ch2_scale = 10 # without linear reg
ch2_position = -1
ch2_bw = 20
ch2_rel_x_position = 20
ch2_label = "VR"
ch2_measure = "MAX,MIN,MEAN"
ch2_color = "YELLOW"
ch2_coupling = "DCLimit"
ch2_offset = 0

"""CHANNEL 3"""
ch3_scale = 2
ch3_position = -3
ch3_bw = 500
ch3_rel_x_position = 30
ch3_label = "PRI_IDS"
ch3_measure = "MAX,MIN,MEAN"
ch3_color = "LIGHT_BLUE"
ch3_coupling = "DCLimit"

"""CHANNEL 4"""
ch4_scale = 10
ch4_position = -4
ch4_bw = 20
ch4_rel_x_position = 45
ch4_label = "LYT8 VOUT"
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


def startup_short():


    # timperdiv = 0.1
    # scope.position_scale(time_position = 10, time_scale = timeperdiv)
    # scope.time_scale(0.04)

    scope.cursor(channel=ids_channel, cursor_set=1, X1=0, X2=0.052, Y1=0, Y2=0, type='VERT')
    scope.cursor(channel=ids_channel, cursor_set=2, X1=0.052, X2=0.314, Y1=0, Y2=0, type='VERT')
    # scope.cursor(channel=ids_channel, cursor_set=3, X1=1.134, X2=1.38, Y1=0, Y2=0, type='VERT')
    # scope.cursor(channel=ids_channel, cursor_set=4, X1=1.38, X2=4.87, Y1=0, Y2=0, type='VERT')

    scope.remove_zoom()
    scope.add_zoom(rel_pos=18, rel_scale=20)

    for vin in vin_list:
        eload.channel[eload_channel].cr = cr_load
        eload.channel[eload_channel].turn_on()

        scope.run_single()

        soak(2)

        prompt("Short output. ")

        soak(2)
        ac.voltage = vin
        ac.turn_on()

        soak(4*timeperdiv)

        discharge_output(1)

        prompt("Adjust cursor - Capture waveforms.")


        crr = round(cr_load,2)
        filename = f"{vin}Vac, {vout}V, {crr}Ohms (100 CR Load), {test}.png"
        screenshot(filename, path)
        screenshot_looper(filename, path)

def startup_short_NL():

    scope.time_scale(1.5)

    # scope.cursor(channel=ids_channel, cursor_set=1, X1=0, X2=0.249, Y1=0, Y2=0, type='VERT')
    # scope.cursor(channel=ids_channel, cursor_set=2, X1=0.567, X2=0.8145, Y1=0, Y2=0, type='VERT')
    # scope.cursor(channel=ids_channel, cursor_set=3, X1=1.134, X2=1.38, Y1=0, Y2=0, type='VERT')
    # scope.cursor(channel=ids_channel, cursor_set=4, X1=1.38, X2=4.87, Y1=0, Y2=0, type='VERT')

    scope.cursor(channel=ids_channel, cursor_set=1, X1=0, X2=0.099, Y1=0, Y2=0, type='VERT')
    scope.cursor(channel=ids_channel, cursor_set=2, X1=0.099, X2=0.414, Y1=0, Y2=0, type='VERT')
    scope.cursor(channel=ids_channel, cursor_set=3, X1=0.435, X2=0.753, Y1=0, Y2=0, type='VERT')
    scope.cursor(channel=vout_channel, cursor_set=4, X1=1.38, X2=4.87, Y1=57.235, Y2=59.807, type='HOR')

    scope.remove_zoom()
    scope.add_zoom(rel_pos=18, rel_scale=20)

    tts("Short output. ")

    for vin in vin_list:

        eload.channel[eload_channel].turn_off()

        scope.run_single()

        soak(2)

        

        ac.voltage = vin
        ac.turn_on()

        soak(6*timeperdiv)

        discharge_output(1)

        prompt("Adjust cursor - Capture waveforms.")

        filename = f"{vin}Vac, NL, {test}.png"
        screenshot(filename, path)



def running_short():



    timeperdiv = 1.5



    # scope.cursor(channel=ids_channel, cursor_set=1, X1=1.275, X2=2.565, Y1=0, Y2=0, type='VERT')
    # scope.cursor(channel=ids_channel, cursor_set=2, X1=3.855, X2=5.13, Y1=0, Y2=0, type='VERT')
    # scope.cursor(channel=ids_channel, cursor_set=3, X1=2.102, X2=5.6, Y1=0, Y2=0, type='VERT')


    scope.position_scale(time_position = 20, time_scale = timeperdiv)
    scope.edge_trigger(trigger_channel=vout_channel, trigger_level=40, trigger_edge='NEG')
    scope.remove_zoom()
    # scope.add_zoom(rel_pos=31.595, rel_scale=10)
        

    for vin in vin_list:

        # discharge_output(1)

        eload.channel[eload_channel].cr = cr_load
        eload.channel[eload_channel].turn_on()

        ac.voltage = vin
        ac.turn_on()

        scope.run_single()

        sleep(3*timeperdiv)

        tts("Short output. ")

        sleep(6*timeperdiv)

        tts("Remove short output. ")
        
        sleep(2*timeperdiv)

        discharge_output(1)

        prompt("Adjust cursor - Capture waveforms.")
        
        discharge_output(1)

        crr = round(cr_load,2)
        filename = f"{vin}Vac, {vout}V, {crr}Ohms (100 CR Load), {test}.png"
        screenshot(filename, path)


def running_short_NL():

    timeperdiv = 2

    scope.position_scale(time_position = 20, time_scale = timeperdiv)
    scope.edge_trigger(trigger_channel=ids_channel, trigger_level=2.5, trigger_edge='POS')
    scope.remove_zoom()
    # scope.add_zoom(rel_pos=28.992, rel_scale=10)

    scope.channel_settings(state=ch1_enable, channel=1, scale=1, position=ch1_position, label=ch1_label,
                        color=ch1_color, rel_x_position=ch1_rel_x_position, bandwidth=ch1_bw, coupling=ch1_coupling, offset=0)       


    for vin in vin_list:


        eload.channel[eload_channel].turn_off()

        ac.voltage = vin
        ac.turn_on()
        
        sleep(2)

        scope.run_single()

        sleep(3*timeperdiv)

        tts("Short output. ")

        sleep(3*timeperdiv)

        tts("Remove short output. ")
        
        sleep(4*timeperdiv)

        discharge_output(1)

        prompt("Adjust cursor. Capture waveforms.")
        
        discharge_output(1)

        filename = f"{vin}Vac, NL, {test}.png"
        screenshot(filename, path)

    scope.channel_settings(state=ch1_enable, channel=1, scale=ch1_scale, position=ch1_position, label=ch1_label,
                        color=ch1_color, rel_x_position=ch1_rel_x_position, bandwidth=ch1_bw, coupling=ch1_coupling, offset=0)    


def operation():

    if test_type == 1: running_short()
    elif test_type == 2: startup_short()
    elif test_type == 3: running_short_NL()
    elif test_type == 4: startup_short_NL()
    else:
        input("Invalid test type.")
    
    
    





def main():
    global waveform_counter
    discharge_output(times=1)
    scope_settings()
    operation()
        
if __name__ == "__main__":
    headers(test)
    main()
    footers(waveform_counter)












