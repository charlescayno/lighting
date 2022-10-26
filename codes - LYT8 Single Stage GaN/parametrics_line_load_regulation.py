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
scope_address = "10.125.10.101"
sig_gen_address = '11'
dimming_power_meter_address = 21
eload_channel = 5

##################################################################################

vin_list = [90,100,110,115,120,132,180,200,230,265,277,300]
vin_list = [132,180,200,230,265,277,300]
vout = 50
iout_nom = 1.5
cr_load = vout/iout_nom
line_soak_time = int(input(">> Soak Time per Line (sec): "))
load_soak_time = int(input(">> Soak Time per Load (sec): "))
percent_list = range(100, -1, -1) # 100% to 0% loading

iout_list = [percent * iout_nom for percent in percent_list]
cr_list = [(cr_load*100/percent if percent != 0 else 0) for percent in percent_list]

test = f"Parametrics - Line Load Regulation"
condition = "FINAL - LINE LOAD REGULATION"
project = "APPS EVAL/Single Stage GAN/Test Data (CV)"
excel_name = f"{condition}"

waveforms_folder = f'C:/Users/{username}/Desktop/{project}/{date}/{test}/'

path = path_maker(f'{waveforms_folder}')

"""EQUIPMENT INITIALIZE"""
ac = ACSource(ac_source_address)
pms = PowerMeter(source_power_meter_address)
pml = PowerMeter(load_power_meter_address)
# pml1 = PowerMeter(load_power_meter_address_1)
eload = ElectronicLoad(eload_address)

"""DISCHARGE FUNCTION"""
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

def collect_data(vin, percent):

    ################ SUBJECT TO CHANGE ################
    vac = vin
    freq = ac.set_freq(vin)
    vac_actual = float(f"{pms.voltage:.6f}")
    iin = float(f"{pms.current*1000:.6f}")
    pin = float(f"{pms.power:.6f}")
    pf = float(f"{pms.pf:.6f}")
    thd = float(f"{pms.thd:.6f}")
    vo1 = float(f"{pml.voltage:.6f}")
    io1 = float(f"{pml.current*1000:.6f}")
    po1 = float(f"{pml.power:.6f}")
    vreg1 = float(f"{100*(float(vo1)-vout)/float(vo1):.6f}")
    iout1 = (float(io1)/1000)
    try: ireg1 = float(f"{100*(iout1-iout_nom)/iout1:.6f}")
    except: ireg1 = 0
    eff = float(f"{100*(float(po1))/float(pin):.6f}")
    # vpfc = float(f"{pml1.voltage:.6f}")

   
    output_list = [percent, vac, freq, vac_actual, iin, pin, pf, thd, vo1, io1, po1, vreg1, ireg1, eff]

    return output_list

def initialize_dataframe():
    ################### COPY PASTE THIS CODE IN HEADER OF MAIN ###################
    df_header_list = ['Load (%)', 'Vin', 'Freq (Hz)', 'Vac (VAC)', 'Iin (mA)', 'Pin (W)', 'PF', 'THD (%)', 'Vo (V)', 'Io1 (mA)', 'Po (W)', 'Vreg (%)', 'Ireg (%)', 'Eff (%)'] 
    df = create_header_list(df_header_list)
    ##############################################################################

    return df

def operation():

    df = initialize_dataframe()
    

    eload.channel[eload_channel].cr = cr_load
    eload.channel[eload_channel].turn_on()

    ac.voltage = vin_list[0]
    ac.turn_on()

    soak(2)

    eload.channel[eload_channel].cr = cr_load
    eload.channel[eload_channel].turn_on()

    for vin in vin_list:

        df = initialize_dataframe()
        print(df)

        ac.voltage = vin
        ac.turn_on()
    
        soak(line_soak_time)

        for cr in cr_list:
            percent = cr_load*100/cr if cr != 0 else 0
            print()
            print(vin, percent, cr)
            if cr != 0:
                eload.channel[eload_channel].cr = cr
                eload.channel[eload_channel].turn_on()
            else:
                eload.channel[eload_channel].turn_off()

            soak(load_soak_time)
            output_list = collect_data(vin, percent)
            export_to_excel(df, waveforms_folder, output_list, excel_name, sheet_name=f"{vin}Vac", anchor="B2")
            print(df)

        discharge_output(1)
        
        

    discharge_output(2)

    
    
def main():
    global waveform_counter
    discharge_output(1)
    operation()
        
if __name__ == "__main__":
    estimated_time = 2 + len(vin_list)*line_soak_time + len(vin_list)*len(cr_list)*load_soak_time + 3*2
    print(f"Estimated Testing Time: {math.ceil(estimated_time/(60))} mins. or ~{round((estimated_time/(60*60)),1)} hrs.")


    headers(test)
    main()
    footers(waveform_counter)

    print(f"Estimated Testing Time: {math.ceil(estimated_time/(60))} mins. or ~{round((estimated_time/(60*60)),1)} hrs.")
    
