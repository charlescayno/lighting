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
dimming_power_meter_address = 10
eload_address = 8
scope_address = "10.125.11.10"
dc_source_address = '4'

"""USER INPUT"""
vin_list = [90,115,230,277]
# vin_list = [115,230,277]
led_list = [36,24]
# led_list = [48]

vout = 48
iout = 1.59

test_condition = input(">> UNIT #: ")

start = 0
mid = 0.01
end = 0.5

step1 = 0.001 # ideally 0.1 V
step2 = 0.01 # ideally 0.5 V
soak_time_1 = 10 # ideally 20 s
soak_time_2 = 10 # ideally 20 s

test = f"Analog"
waveforms_folder = f'C:/Users/ccayno/Desktop/DER/DER-727/Test Data/Final Data/Dimming/{test}/{test_condition}'
path = path_maker(f'{waveforms_folder}')


"""DO NOT EDIT BELOW THIS LINE"""
##################################################################################

"""EQUIPMENT INITIALIZE"""
ac = ACSource(ac_source_address)
pms = PowerMeter(source_power_meter_address)
pml = PowerMeter(load_power_meter_address)
pmdc = PowerMeter(dimming_power_meter_address)
eload = ElectronicLoad(eload_address)
dc = Keithley_DC_2230G(dc_source_address)
led_ctrl = LEDControl()

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

def collect_data(dim, v_input):
    """COLLECT DATA"""
    dim_voltage = dim
    actual_dim_voltage = float(f"{pmdc.voltage:.2f}")
    vac = v_input
    freq = ac.set_freq(v_input)
    try:
        vin = float(f"{pms.voltage:.4f}")
    except:
        vin = 0
    iin = float(f"{pms.current*1000:.4f}")
    pin = float(f"{pms.power:.4f}")
    pf = float(f"{pms.pf:.4f}")
    thd = float(f"{pms.thd:.4f}") # AC
    vo1 = float(f"{pml.voltage:.4f}")
    io1 = float(f"{pml.current*1000:.6f}")
    po1 = float(f"{pml.power:.4f}")
    vreg1 = float(f"{100*(float(vo1)-vout)/float(vo1):.4f}")
    iout1 = float(io1)/1000
    try: ireg1 = float(f"{100*(iout1-iout)/iout1:.4f}")
    except: ireg1 = 0
    eff = float(f"{100*(float(po1))/float(pin):.4f}")

    output_list = [dim_voltage, actual_dim_voltage, vac, freq, vin, iin, pin, pf, thd, vo1, io1, po1, vreg1, ireg1, eff] # AC
    
    return output_list


def set_dim_voltage(dim):
    dc.set_volt_curr(channel ='CH1', voltage = dim, current = 1.0)
    dc.channel_state(channel ='CH1', state = 'ON')
    soak(1)


def export_df_to_excel(v_input, df, led):
    
    if v_input == 90:
        anchor = 'B2' # starting anchor
    if v_input == 115:
        anchor = 'R2' # starting anchor
    if v_input == 230:
        anchor = 'AH2' # starting anchor
    if v_input == 277:
        anchor = 'AX2' # starting anchor

    # create an excel file output
    src = f"{os.getcwd()}/template_analog_dimming.xlsx"
    dst = f"{waveforms_folder}/{test}.xlsx"
    if not os.path.exists(dst): shutil.copyfile(src, dst)

    wb = load_workbook(dst)
    df_to_excel(wb, f"{led}", df, anchor)
    wb.save(dst)

    # moving anchor
    col, row = col_row_extractor(anchor)
    col = col + len(df.columns) + 1
    anchor = get_anchor(col, row)

    df = pd.DataFrame(columns = ['Dim', 'Actual Dim', 'Vin', 'Freq', 'Vac', 'Iin', 'Pin', 'PF', 'THD', 'Vo1', 'Io1', 'Po1', 'Vreg1', 'Ireg1', 'Eff']) # reset df

    return df




def operation():

    for led in led_list:

        led_ctrl.voltage(led)

        set_dim_voltage(0)

        
        df = pd.DataFrame(columns = ['Dim', 'Actual Dim', 'Vin', 'Freq', 'Vac', 'Iin', 'Pin', 'PF', 'THD', 'Vo1', 'Io1', 'Po1', 'Vreg1', 'Ireg1', 'Eff'])
        df_header = df.columns.values.tolist()


        for v_input in vin_list:

            for dim in np.arange(start, mid+step1, step1):

                set_dim_voltage(dim)

                ac.voltage = v_input
                ac.turn_on()

                if dim == 0:
                    soak(60)
                else:
                    soak(soak_time_1)

                output_list = collect_data(dim, v_input)

                df.loc[len(df)] = output_list # save to dataframe
                print(df)

            for dim in np.arange(mid, end+step2, step2):

                set_dim_voltage(dim)

                ac.voltage = v_input
                ac.turn_on()

                soak(soak_time_2)

                output_list = collect_data(dim, v_input)

                df.loc[len(df)] = output_list
                print(df)

            df = export_df_to_excel(v_input, df, led)
        
            discharge_output()
            set_dim_voltage(0)

            soak(5)

        set_dim_voltage(0)
        dc.channel_state(channel ='CH1', state = 'OFF')

        soak(3)


def main():
    global waveform_counter
    discharge_output()
    operation()
    discharge_output()
        
if __name__ == "__main__":
    headers(test)
    main()
    footers(waveform_counter)