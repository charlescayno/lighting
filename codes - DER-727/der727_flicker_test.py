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
dimming_power_meter_address = 10
source_power_meter_address = 30
load_power_meter_address = 2
eload_address = 8
scope_address = "10.125.11.10"
dc_source_address = '4'

"""USER INPUT"""
vin_list = [90,115,230,277]
# vin_list = [230]
led_list = [24]

# vin_list = [230]
vout = 48
iout = 1.59

# led = '36V'

# led = input("Enter LED voltage (48V / 36V / 30V / 24V): ")
# test_condition = input("Special test condition: ")


start = 0
mid = 1
end = 10

step1 = 0.1 # ideally 0.1 V
step2 = 1 # ideally 0.5 V
soak_time_1 = 10 # ideally 20 s
soak_time_2 = 10 # ideally 20 s


test = f"Flicker Test"
waveforms_folder = f'C:/Users/ccayno/Desktop/DER/DER-727/Test Data/Final Data/Dimming/{test}'
path = path_maker(f'{waveforms_folder}')


"""DO NOT EDIT BELOW THIS LINE"""
##################################################################################

"""EQUIPMENT INITIALIZE"""
ac = ACSource(ac_source_address)
pms = PowerMeter(source_power_meter_address)
pml = PowerMeter(load_power_meter_address)
pmdc = PowerMeter(dimming_power_meter_address)
eload = ElectronicLoad(eload_address)
# scope = Oscilloscope(scope_address)
dc = Keithley_DC_2230G(dc_source_address)
led_ctrl = LEDControl()

from AUTOGUI_FLICKER import *
ate_gui = AutoguiCalibrate()

def discharge_output(times):
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

# def screenshot(filename, path):
#     global waveform_counter

#     scope.get_screenshot(filename, path)
#     print(filename)
#     waveform_counter += 1





def set_dim_voltage(dim):
    dc.set_volt_curr(channel ='CH1', voltage = dim, current = 1.0)
    dc.channel_state(channel ='CH1', state = 'ON')
    soak(1)

def dc_source_off():
    dc.channel_state(channel ='CH1', state = 'OFF')


def get_flicker(led, vin, dim):

    # prompt("Getting flicker")
    print(f"{led}V, {vin}Vac, Dim = {dim}V")
    
    # sleep(1)
    # ate_gui.click('flicker_button')
    # sleep(1)
    # ate_gui.click('flicker_button')
    # sleep(1)


    soak(30)
        
    image_file = f"{led}V_{vin}Vac_{dim}V.png"
    pyautogui.screenshot(image_file)
    path = path_maker(f'{waveforms_folder}/{led}V/')
    source = f'{os.getcwd()}/{image_file}'
    destination = f'{path}/{image_file}'
    remove_file(destination)
    shutil.move(source, destination)


    ate_gui.click('select_file_tab')
    soak(1)
    ate_gui.click('select_save_as')
    soak(1)
    pyautogui.write(f"{led}V_{vin}Vac_{dim}V")
    soak(1)
    pyautogui.press('enter')
    soak(1)


    



def operation():

    ate_gui.recalibrate()

    ate_gui.click('select_flicker_app')

    for led in led_list:

        led_ctrl.voltage(led)

        set_dim_voltage(0)

        for vin in vin_list:

            for dim in np.arange(start, mid+step1, step1):

                set_dim_voltage(dim)

                ac.voltage = vin
                ac.turn_on()

                soak(5)

                get_flicker(led, vin, dim)


            for dim in np.arange(mid, end+step2, step2):

                set_dim_voltage(dim)

                ac.voltage = vin
                ac.turn_on()

                soak(5)

                get_flicker(led, vin, dim)


            discharge_output(5)
            set_dim_voltage(0)

            soak(5)

        set_dim_voltage(0)
        dc_source_off()

        soak(3)


def main():
    global waveform_counter
    discharge_output(1)
    operation()
    discharge_output(1)
        
if __name__ == "__main__":
    headers(test)
    main()
    footers(waveform_counter)