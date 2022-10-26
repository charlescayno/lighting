"""IMPORT DEPENDENCIES"""
from time import time, sleep
import sys
import os
import math
import numpy as np
from powi.equipment import ACSource, PowerMeter, ElectronicLoad, Oscilloscope, LEDControl, Keithley_DC_2230G
from powi.equipment import headers, create_folder, footers, waveform_counter, soak, convert_argv_to_int_list, tts, prompt
from filemanager import path_maker, remove_file
import winsound as ws
from playsound import playsound
waveform_counter = 0

from datetime import datetime
now = datetime.now()
# date = now.strftime('%Y_%m_%d')
date = now.strftime('%Y_%m_%d_%H%M')

import shutil
from openpyxl import Workbook, load_workbook
import pandas as pd
from powi.equipment import excel_to_df, df_to_excel, image_to_excel, col_row_extractor, get_anchor

##################################################################################

"""COMMS"""
ac_source_address = 5
source_power_meter_address = 30 
load_power_meter_address = 2
eload_address = 8
scope_address = "10.125.11.20"
dc_source_address = '1'
dimming_power_meter_address = 11

"""USER INPUT"""
vout = 48
iout = 1.59

vin = float(sys.argv[1])
dim = float(sys.argv[2])


test = f"Analog"
waveforms_folder = f'C:/Users/ccayno/Desktop/DER/DER-727/Test Data/Dimming/{test}'
path = path_maker(f'{waveforms_folder}')


"""DO NOT EDIT BELOW THIS LINE"""
##################################################################################

"""EQUIPMENT INITIALIZE"""
ac = ACSource(ac_source_address)
pms = PowerMeter(source_power_meter_address)
pml = PowerMeter(load_power_meter_address)
pmdc = PowerMeter(dimming_power_meter_address)
eload = ElectronicLoad(eload_address)
scope = Oscilloscope(scope_address)
dc = Keithley_DC_2230G(dc_source_address)

def discharge_output():
    ac.turn_off()
    for i in range(1,9):
        eload.channel[i].cc = 1
        eload.channel[i].turn_on()
        eload.channel[i].short_on()
    sleep(1)
    for i in range(1,9):
        eload.channel[i].turn_off()
        eload.channel[i].short_off()
    sleep(1)

def screenshot(filename, path):
    global waveform_counter

    scope.get_screenshot(filename, path)
    print(filename)
    waveform_counter += 1


def operation():

    dc.set_volt_curr(channel ='CH1', voltage = dim, current = 1.0)
    dc.channel_state(channel ='CH1', state = 'ON')

    soak(1)

    ac.voltage = vin
    ac.turn_on()

    input()

    discharge_output()



def main():
    global waveform_counter
    # discharge_output()
    operation()
    # discharge_output()
        
if __name__ == "__main__":
    headers(test)
    main()
    footers(waveform_counter)