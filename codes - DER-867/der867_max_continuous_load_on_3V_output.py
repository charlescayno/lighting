"""COMMS"""
ac_source_address = 5
source_power_meter_address = 1 
load_power_meter_address = 2    
eload_address = 8
scope_address = "10.125.11.0"

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

"""INITIALIZE EQUIPMENT"""
ac = ACSource(ac_source_address)
pms = PowerMeter(source_power_meter_address)
pml = PowerMeter(load_power_meter_address)
eload = ElectronicLoad(eload_address)
# scope = Oscilloscope(scope_address)

"""USER'S INPUT"""
vin_list = [230]
soak_time = 10
test = f"DER-867 (Line Regulation)"
waveforms_folder = f'C:/Users/ccayno/Desktop/DER/DER-867/Final Test Data/{test}'
path = path_maker(f'{waveforms_folder}')



def discharge_output():
    ac.turn_off()
    for i in range(1,9):
        eload.channel[i].cc = 1
        eload.channel[i].turn_on()
        eload.channel[i].short_on()
    sleep(2)
    for i in range(1,9):
        eload.channel[i].turn_off()
        eload.channel[i].short_off()
    sleep(1)


def test_line_regulation(vin_list, soak_time):
    global waveform_counter
    
    ac.voltage = 230
    ac.turn_on()

    input("Turn on relay")

    print("rload_W, vac, freq, vin, iin, pin, pf, thd, vo1, io1, po1 (mW), eff")

    for i in vin_list:

        for j in range(400,0,-10):

            CR = j
            eload.channel[1].cr = CR
            eload.channel[1].turn_on()
            sleep(1)

            ac.voltage = i
            ac.turn_on()

            soak(3)

            soak(soak_time)

            # create output list
            rload_w = str(j)
            vac = str(i)
            freq = str(ac.frequency)
            vin = f"{pms.voltage:.2f}"
            iin = f"{pms.current*1000:.2f}"
            pin = f"{pms.power:.3f}"
            pf = f"{pms.pf:.4f}"
            thd = f"{pms.thd:.2f}"
            vo1 = f"{pml.voltage:.3f}"
            io1 = f"{pml.current*1000:.2f}"
            po1 = f"{pml.power*1000:.3f}"
            eff = f"{100*(float(po1))/float(pin):.4f}"

            output_list = [rload_w, vac, freq, vin, iin, pin, pf, thd, vo1, io1, po1, eff]

            print(','.join(output_list))

            soak(3)

    
    footers(waveform_counter)


if __name__ == "__main__":
    headers(test)
    discharge_output()
    test_line_regulation(vin_list, soak_time)
    discharge_output()
    footers(waveform_counter)


