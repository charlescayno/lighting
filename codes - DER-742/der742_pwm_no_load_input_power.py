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

from powi.equipment import ACSource, PowerMeter, ElectronicLoad, Oscilloscope, LEDControl, Keithley_DC_2230G, Tektronix_SigGen_AFG31000
from powi.equipment import headers, create_folder, footers, waveform_counter, soak, convert_argv_to_int_list, tts, prompt
from powi.equipment import excel_to_df, df_to_excel, image_to_excel, col_row_extractor, get_anchor
from powi.equipment import create_header_list, export_to_excel, export_screenshot_to_excel
from powi.equipment import path_maker, remove_file, start_timer, end_timer

import getpass
username = getpass.getuser().lower()

from datetime import datetime
now = datetime.now()
date = now.strftime('%m%d')
##################################################################################

"""COMMS"""
ac_source_address = 5
source_power_meter_address = 30 
load_power_meter_address_1 = 10 
load_power_meter_address = 2
sig_gen_address = 11
eload_address = 8
scope_address = "10.125.11.101"
dc_source_address = '4'

##################################################################################
condition = input("Enter condition: ")
vin_list = [230]
vout = 42
iout = 1
soak_time = 30 # s
integration_time = 30 # s
frequency_list = [300]
duty_list = [0, 10, 20, 30, 40, 50, 60, 70, 80, 100]

project = "Standby Efficiency at Open Load (DER-742)"
test = "Input Power at Different Dim"
waveforms_folder = f'C:/Users/{username}/Desktop/{project}/{date}/{test}/'
path = path_maker(f'{waveforms_folder}')
excel_name = f"{test}"


"""EQUIPMENT INITIALIZE"""
ac = ACSource(ac_source_address)
pms = PowerMeter(source_power_meter_address)
pml = PowerMeter(load_power_meter_address)
eload = ElectronicLoad(eload_address)
sig_gen = Tektronix_SigGen_AFG31000(sig_gen_address)


"""DISCHARGE FUNCTION"""
def discharge_output(times=1):
    ac.turn_off()
    for i in range(times):
        for i in range(1,9):
            eload.channel[i].cc = 1
            eload.channel[i].turn_on()
            eload.channel[i].short_on()
        sleep(2)

        for i in range(1,9):
            eload.channel[i].turn_off()
            eload.channel[i].short_off()
        sleep(2)

def collect_data(vin, frequency, duty):


    ################ SUBJECT TO CHANGE ################
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
    try: ireg1 = float(f"{100*(iout1-iout)/iout1:.6f}")
    except: ireg1 = 0
    eff = float(f"{100*(float(po1))/float(pin):.6f}")

   
    output_list = [frequency, duty, vin, freq, vac_actual, iin, pin, pf, thd, vo1, io1, po1, vreg1, ireg1, eff]

    return output_list



def set_pwm(frequency, duty):
    if duty == 100: duty = 99
    if duty == 0: duty = 0.005
    sig_gen.set_load_impedance(channel = 1, impedance = 'INF') 
    sig_gen.out_cont_pulse(channel = 1, frequency = frequency, phase = '0 DEG', low = '0V', high = '10V', units = 'VPP', duty = duty, width = 'ABC')
    sig_gen.channel_state(channel = 1, state ='ON')
    soak(2)


def operation():


    ################### COPY PASTE THIS CODE IN HEADER OF MAIN ###################
    df_header_list = ['Frequency (Hz)', 'Duty (%)', 'Vin', 'Freq (Hz)', 'Vac (VAC)', 'Iin (mA)', 'Pin (W)', 'PF', 'THD (%)', 'Vo (V)', 'Io1 (mA)', 'Po (W)', 'Vreg (%)', 'Ireg (%)', 'Eff (%)'] 
    df = create_header_list(df_header_list)
    ##############################################################################

    estimated_time = (len(frequency_list)*len(duty_list)*len(vin_list)*(soak_time+integration_time+5))/60
    print(f"Estimated time: {estimated_time} mins.")


    for frequency in frequency_list:
        for duty in duty_list:

            set_pwm(frequency, duty)

            for vin in vin_list:

                ac.voltage = vin
                ac.turn_on()
            
                soak(soak_time)
                pms.integrate(integration_time)
                soak(integration_time+5)

                output_list = collect_data(vin, frequency, duty)
                export_to_excel(df, waveforms_folder, output_list, excel_name, sheet_name=f"{vin}Vac, f={frequency}Hz, {condition}", anchor="B2")
                print(df)

    print(f"\n\nFinal Data: ")
    print(df)

    discharge_output(2)
    pms.reset()
    pml.reset()

    
def main():
    global waveform_counter
    discharge_output(1)
    operation()

        
if __name__ == "__main__":
    headers(test)
    main()
    footers(waveform_counter)
