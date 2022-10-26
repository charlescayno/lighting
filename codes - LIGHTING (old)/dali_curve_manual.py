
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
import winsound as ws
from playsound import playsound
waveform_counter = 0

"""INITIALIZE EQUIPMENT"""
ac = ACSource(ac_source_address)
pms = PowerMeter(source_power_meter_address)
pml = PowerMeter(load_power_meter_address)
eload = ElectronicLoad(eload_address)
scope = Oscilloscope(scope_address)

"""USER'S INPUT"""
vin_list = [230]
iout_reg = 1050
soak_time = 30
test = f"TCI (DALI Curve)"
waveforms_folder = f'waveforms/{test}'

def discharge_output():
    ac.turn_off()
    for i in range(1,9):
        eload.channel[i].cc = 1
        eload.channel[i].turn_on()
        eload.channel[i].short_on()
    sleep(3)
    for i in range(1,9):
        eload.channel[i].turn_off()
        eload.channel[i].short_off()
    sleep(1)


def test_line_regulation():
    global waveform_counter
    
    ac.voltage = vin
    ac.frequency = ac.set_freq(vin)
    ac.turn_on()

    for i in range(255):
        
        input(f"Change cmd bit to {i} in the DALI program. Press ENTER to record Iout")

        # create output list
        vac = str(vin)
        freq = str(ac.set_freq(vin))
        vin = f"{pms.voltage:.2f}"
        iin = f"{pms.current*1000:.2f}"
        pin = f"{pms.power:.3f}"
        pf = f"{pms.pf:.4f}"
        thd = f"{pms.thd:.2f}"
        vo1 = f"{pml.voltage:.3f}"
        io1 = f"{pml.current*1000:.2f}"
        po1 = f"{pml.power:.3f}"
        ireg1 = f"{100*(float(io1)-iout_reg)/iout_reg:.4f}"
        eff = f"{100*(float(po1))/float(pin):.4f}"

        output_list = [vac, freq, vin, iin, pin, pf, thd, vo1, io1, po1, ireg1, eff]

        print(','.join(output_list))

        soak(10)

    
    footers(waveform_counter)


if __name__ == "__main__":
    headers(test)
    discharge_output()
    test_line_regulation()
    discharge_output()
    footers(waveform_counter)


