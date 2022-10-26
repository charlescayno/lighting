from xml.sax.handler import feature_external_ges
from powi.equipment import Keithley_DC_2230G
import pyautogui
from time import sleep, time
import pandas as pd
import os
import shutil
import numpy as np
# C:\Users\ccayno\Anaconda3\Lib\site-packages\powi
# C:\Users\ccayno\Desktop\Audible Automation
#########################################################################################
"""IMPORT DEPENDENCIES"""
from time import time, sleep
import sys
import os
import math
import numpy as np
import shutil
import os
import pandas as pd
import pyautogui

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
#########################################################################################
"""COMMS"""
ac_source_address = 5
source_power_meter_address = 30 
load_power_meter_address = 2
eload_address = 8
eload_channel = 5
#########################################################################################
"""USER INPUTS"""
test = "Audible Noise using E-load"
print(test)
condition = input(">> Project/IC: ")
vin_list = convert_argv_to_int_list(input(">> Enter input voltage list, vin_list (VAC): "))
percent_list = range(100, -1, -1) # 100% to 0% loading
vout = float(input(">> Enter Vout (V): "))
iout = float(input(">> Enter Vout (V): "))
cr_load = (vout/iout)
cr_list = [(cr_load/percent if percent != 0 else 0) for percent in percent_list]
#########################################################################################
ac = ACSource(ac_source_address)
pms = PowerMeter(source_power_meter_address)
pml = PowerMeter(load_power_meter_address)
eload = ElectronicLoad(eload_address)
#########################################################################################

def exists(filename):
    return os.path.isfile(filename)

def get_file_size(filename):
    if exists():        
        statinfo = os.stat(filename)
        return statinfo.st_size
    return 0

def file_to_df(filename, column_name):
    df = pd.read_excel(filename, skiprows=3)
    df.set_index('Hz', inplace=True)
    df.index.name = 'Frequency'
    df.columns = [column_name]
    return df

def append_df_to_file(filename, df):
    if exists(filename):
        base_df = pd.read_excel(filename, index_col=0)
        base_df = pd.concat([base_df, df], axis=1)
    else:
        base_df = df
    base_df.to_excel(filename)

def copy_file(src, dst):
    shutil.copyfile(src, dst)

def delete_file(filename):
    try:
        os.remove(filename)
    except:
        pass


def operation(percent, raw_acoustic_file, output_file):
    
    sleep(1)
    pyautogui.moveTo(button_position)
    sleep(1)
    pyautogui.click()

    while not exists(raw_acoustic_file):
        sleep(1)
    raw_df = file_to_df(filename=raw_acoustic_file, column_name=f'{percent}')
    append_df_to_file(output_file, raw_df)
    delete_file(raw_acoustic_file)
    sleep(3)




def main():

    global button_position
     
    start = time()

    raw_acoustic_file = 'FrequencyResponse.xlsx'
    button_position = (30, 135)

    for vin in vin_list:

        ac.voltage = vin
        ac.turn_on()
        sleep(5)

        for percent in percent_list:

            cr = cr_load/percent if percent != 0 else 0

            if cr == 0:
                eload.channel[eload_channel].turn_off()

            else:
                eload.channel[eload_channel].cr = cr
                eload.channel[eload_channel].turn_on()
            
            sleep(3)

            output_file = f'{vin}VAC.xlsx'

            delete_file(output_file)

            operation(percent, raw_acoustic_file, output_file)

            break

    end = time()
    print(f'Elapsed: {end-start} s')

        
if __name__ == "__main__":
    headers(test)
    main()
    footers(waveform_counter)
