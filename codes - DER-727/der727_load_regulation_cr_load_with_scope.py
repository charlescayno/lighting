
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
vin_list = [230]
iout = 1.59

special_test_condition = sys.argv[1]
test = "LoadRegulation_CR"
waveforms_folder = f'C:/Users/ccayno/Desktop/DER/DER-727/Test Data/{test}'
path = path_maker(f'{waveforms_folder}')


"""DO NOT EDIT BELOW THIS LINE"""
##################################################################################

"""EQUIPMENT INITIALIZE"""
ac = ACSource(ac_source_address)
pms = PowerMeter(source_power_meter_address)
pml = PowerMeter(load_power_meter_address)
eload = ElectronicLoad(eload_address)
scope = Oscilloscope(scope_address)

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


def test_line_regulation():

    # df = pd.DataFrame(columns = ['CR', 'Vin', 'Freq', 'Vac', 'Iin', 'Pin', 'PF', 'THD', 'Vo1', 'Io1', 'Po1', 'Vreg1', 'Ireg1', 'Eff']) # AC
    df = pd.DataFrame(columns = ['CR (Ohms)', 'Vin', 'Freq', 'Vdc (V)', 'Iin (mA)', 'Pin (W)', 'PF', 'Vo1 (V)', 'Io1 (mA)', 'Po1 (W)', 'Vreg1 (%)', 'Ireg1 (%)', 'Eff (%)', 'ton (us)', 'tfly (us)', 'toff (us)', 'Kp', 'Fsw (kHz)']) # DC
    df_header = df.columns.values.tolist()

    for vout in range(48, 22, -1):

        eload.channel[5].cr = vout/iout
        eload.channel[5].turn_on()

        for vin in vin_list:
            dc = 410
            ac.voltage = dc
            ac.coupling = 'DC'
            ac.turn_on()
        
            soak(5)

            cr_load = float(f"{vout/iout:.2f}")
            vac = vin
            vac = dc # dc
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

            

            scope.run_single()

            input("Adjust cursors. Press ENTER to continue. ")
            
            ton = float(scope.get_cursor(cursor=1)['delta x'])
            tfly = float(scope.get_cursor(cursor=2)['delta x'])
            toff = float(scope.get_cursor(cursor=3)['delta x'])
            tsw = float(scope.get_cursor(cursor=4)['delta x'])
            
            kp = toff/tfly
            fsw = (1/tsw)/1000

            ton = float(f"{ton*1E6:.2f}")
            tfly = float(f"{tfly*1E6:.2f}")
            toff = float(f"{toff*1E6:.2f}")
            
            kp = float(f"{kp:.2f}")
            fsw = float(f"{fsw:.2f}")
            
            filename = f"{vac}VDC, {cr_load}_ohms, {po1}W, {vo1}V, {io1}mA, kp={kp}, fsw={fsw:.2f}kHz.png"
            print(filename)
            scope.get_screenshot(filename, path)

            # output_list = [cr_load, vac, freq, vin, iin, pin, pf, thd, vo1, io1, po1, vreg1, ireg1, eff] # AC
            output_list = [cr_load, vac, freq, vin, iin, pin, pf, vo1, io1, po1, vreg1, ireg1, eff, ton, tfly, toff, kp, fsw] # DC

            df.loc[len(df)] = output_list
            print(df)

    # replace existing file to the template
    src = f"{os.getcwd()}/template_load_regulation.xlsx"
    dst = f"{waveforms_folder}/LoadRegulation_CR_{special_test_condition}_{date}.xlsx"
    shutil.copyfile(src, dst)

    wb = load_workbook(dst)
    # df_to_excel(wb, "Load Regulation", df_header, 'A2') # bug
    df_to_excel(wb, "Load Regulation", df, 'B2')

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