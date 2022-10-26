"""IMPORT DEPENDENCIES"""
from time import time, sleep
import sys
import os
import math
from powi.equipment import ACSource, PowerMeter, ElectronicLoad, Oscilloscope, LEDControl
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
from powi.equipment import excel_to_df, df_to_excel, image_to_excel

##################################################################################

"""COMMS"""
ac_source_address = 5
source_power_meter_address = 1 
load_power_meter_address = 2
eload_address = 8
scope_address = "10.125.11.20"

"""USER INPUT"""
vin_list = [90,100,110,115,120,132,180,200,230,265,277]
# vin_list = [400]
iout = 1.59
iout_list = [iout]
led_list = [48,36,24]
vout = 48

special_test_condition = sys.argv[1]
test = "LineRegulation_LED"
waveforms_folder = f'C:/Users/ccayno/Desktop/DER/DER-727/Test Data/{test}'
path = path_maker(f'{waveforms_folder}')


"""DO NOT EDIT BELOW THIS LINE"""
##################################################################################

"""EQUIPMENT INITIALIZE"""
ac = ACSource(ac_source_address)
pms = PowerMeter(source_power_meter_address)
pml = PowerMeter(load_power_meter_address)
eload = ElectronicLoad(eload_address)
# scope = Oscilloscope(scope_address)

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

def test_line_regulation():

    df = pd.DataFrame(columns = ['LED', 'Vin', 'Freq', 'Vdc', 'Iin', 'Pin', 'PF', 'Vo1', 'Io1', 'Po1', 'Vreg1', 'Ireg1', 'Eff']) # DC

    for led in led_list:

        input(f"Change led mode to {led} V.")

        for vin in vin_list:
            ac.voltage = 400
            ac.coupling = 'DC'
            ac.turn_on()
        
            soak(10)

            led_load = led
            vac = vin
            vac = 400 # dc
            freq = ac.set_freq(vin)
            vin = float(f"{pms.voltage:.2f}")
            iin = float(f"{pms.current*1000:.2f}")
            pin = float(f"{pms.power:.3f}")
            pf = float(f"{pms.pf:.4f}")
            # thd = float(f"{pms.thd:.2f}") # DC
            vo1 = float(f"{pml.voltage:.3f}")
            io1 = float(f"{pml.current*1000:.2f}")
            po1 = float(f"{pml.power:.3f}")
            vreg1 = float(f"{100*(float(vo1)-vout)/float(vo1):.4f}")
            iout1 = float(io1)/1000
            ireg1 = float(f"{100*(iout1-iout)/iout1:.4f}")
            eff = float(f"{100*(float(po1))/float(pin):.4f}")

            output_list = [led_load, vac, freq, vin, iin, pin, pf, vo1, io1, po1, vreg1, ireg1, eff] # DC

            df.loc[len(df)] = output_list

            print(df)

            discharge_output()

    # replace existing file to the template
    src = f"{os.getcwd()}/template_line_regulation.xlsx"
    dst = f"{waveforms_folder}/{test}_{special_test_condition}_{date}.xlsx"
    shutil.copyfile(src, dst)
    wb = load_workbook(dst)
    df_to_excel(wb, "Line Regulation", df, 'A2')
    wb.save(dst)

    soak(3)


def main():
    global waveform_counter
    discharge_output()
    test_line_regulation()
    discharge_output()
        
if __name__ == "__main__":
    headers(test)
    main()
    footers(waveform_counter)