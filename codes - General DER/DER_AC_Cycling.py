# November 11, 2021

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
date = now.strftime('%Y_%m_%d')	

##################################################################################

"""COMMS"""
ac_source_address = 5
source_power_meter_address = 30 
load_power_meter_address = 2
eload_address = 8
scope_address = "10.125.11.10"
sig_gen_address = '11'
dimming_power_meter_address = 21
"""USER INPUT"""
vin_list = [90,115,230,277]
iout = 1.59
iout_list = [iout]
led_load = int(input("Enter led load (48 / 36 / 24): "))
vout = led_load
pulse_list = [1]
led_list = [led_load]
special_test_condition = input("Enter special test condition: ")

test = "AC Cycling"
waveforms_folder = f'C:/Users/ccayno/Desktop/DER/DER-727/Test Data/{test}/{special_test_condition}'
path = path_maker(f'{waveforms_folder}')




"""DO NOT EDIT BELOW THIS LINE"""
##################################################################################

"""EQUIPMENT INITIALIZE"""
ac = ACSource(ac_source_address)
pms = PowerMeter(source_power_meter_address)
pml = PowerMeter(load_power_meter_address)
eload = ElectronicLoad(eload_address)
scope = Oscilloscope(scope_address)

"""GENERIC FUNCTIONS"""

def discharge_output():
    ac.turn_off()
    for i in range(1):
        for i in range(1,9):
            eload.channel[i].cc = 1
            eload.channel[i].turn_on()
            eload.channel[i].short_on()
        sleep(2)

        for i in range(1,9):
            eload.channel[i].turn_off()
            eload.channel[i].short_off()
        sleep(2)



def scope_settings():
    global condition
    global time_division
    
    scope.channel_settings(state='OFF', channel=1, scale=0.4, position=-4, label="Vfb",
                            color='YELLOW', rel_x_position=20, bandwidth=20, coupling='DCLimit', offset=0)
    
    
    scope.channel_settings(state='ON', channel=2, scale=0.2, position=-4, label="Output Current",
                            color='LIGHT_BLUE', rel_x_position=50, bandwidth=20, coupling='DCLimit', offset=0)
    
    
    scope.channel_settings(state='OFF', channel=3, scale=2, position=-4, label="Vaux",
                            color='ORANGE', rel_x_position=60, bandwidth=20, coupling='DCLimit', offset=0)
    
    
    scope.channel_settings(state='ON', channel=4, scale=100, position=0, label="Input Voltage",
                            color='YELLOW', rel_x_position=50, bandwidth=20, coupling='AC', offset=0)
    
    
    
    # scope.measure(1, "MAX,RMS")
    scope.measure(2, "MAX,MIN,RMS,MEAN,PDELta")
    # scope.measure(3, "MAX,RMS")
    scope.measure(4, "MAX,MIN,RMS,MEAN,PDELta")

    scope.time_position(30)
    
    scope.record_length(50E6)
    
    scope.time_scale(time_division)

    # scope.remove_zoom()
    # scope.add_zoom(rel_pos=21.727, rel_scale=1)
    
    trigger_channel = 2
    trigger_level = 0.8
    trigger_edge = 'NEG'
    scope.edge_trigger(trigger_channel, trigger_level, trigger_edge)

    scope.stop()
    
def ac_cycle_4_pulse(vin, start_soak, off_time, on_time, end_soak):
    
    freq = ac.set_freq(vin)
    ac.write(f"TRIG:TRAN:SOUR BUS")
    ac.write(f"LIST:DWEL {start_soak}, {off_time}, {on_time},{off_time}, {on_time}, {off_time}, {on_time}, {off_time}, {on_time}, {off_time}, {end_soak}")
    ac.write(f"VOLT:MODE LIST")
    ac.write(f"LIST:VOLT      {vin}, 0, {vin}, 0, {vin}, 0, {vin}, 0, {vin}, 0, {vin}")
    ac.write(f"VOLT:SLEW:MODE LIST")
    ac.write(f"LIST:VOLT:SLEW 9.9e+037, 9.9e+037, 9.9e+037, 9.9e+037, 9.9e+037, 9.9e+037, 9.9e+037, 9.9e+037, 9.9e+037, 9.9e+037, 9.9e+037")
    ac.write(f"FREQ:MODE LIST")
    ac.write(f"LIST:FREQ       {freq}, {freq}, {freq}, {freq}, {freq}, {freq}, {freq}, {freq}, {freq}, {freq}, {freq}")
    ac.write(f"FREQ:SLEW:MODE LIST")
    ac.write(f"LIST:FREQ:SLEW 9.9e+037, 9.9e+037, 9.9e+037, 9.9e+037, 9.9e+037, 9.9e+037, 9.9e+037, 9.9e+037, 9.9e+037, 9.9e+037, 9.9e+037")
    ac.write(f"VOLT:OFFS:MODE FIX")
    ac.write(f"VOLT:OFFS:SLEW:MODE FIX")
    ac.write(f"PHAS:MODE LIST")
    ac.write(f"LIST:PHAS      270, 270, 270, 270, 270, 270, 270, 270, 270, 270, 270")
    ac.write(f"CURR:PEAK:MODE LIST")
    ac.write(f"LIST:CURR     40.4, 40.4, 40.4, 40.4, 40.4, 40.4, 40.4, 40.4, 40.4, 40.4, 40.4")
    ac.write(f"FUNC:MODE FIX")
    ac.write(f"LIST:TTLT ON,OFF,OFF,OFF,OFF,OFF,OFF,OFF,OFF,OFF,OFF")
    ac.write(f"LIST:STEP AUTO")
    ac.write(f"OUTP:TTLT:STAT ON")
    ac.write(f"OUTP:TTLT:SOUR LIST")
    ac.write(f"TRIG:SYNC:SOUR PHASE")
    ac.write(f"TRIG:SYNC:PHAS 0.0")
    ac.write(f"TRIG:TRAN:DEL 0")
    ac.write(f"Sens:Swe:Offs:Poin 0")
    ac.write(f"TRIG:ACQ:SOUR TTLT")
    ac.write(f"INIT:IMM:SEQ3")
    ac.write(f"LIST:COUN 1")
    ac.write(f"INIT:IMM:SEQ1")
    ac.write(f"TRIG:TRAN:SOUR BUS")
    ac.write(f"TRIG:IMM")

    sleep(13*time_division)



def screenshot(filename, path):
    global waveform_counter

    scope.get_screenshot(filename, path)
    print(filename)
    waveform_counter += 1




def operation():
    global waveform_counter
    # setup time scale for the oscilloscope
    test_time = 2*(end_voltage-start_voltage)*slew_rate+time_fixvoltage
    scope_time = roundup(test_time)
    delay = scope_time-test_time
    time_scale = scope_time/10
    scope.time_scale(time_scale)
    scope.run()
    soak(int(delay/2))

    # start of test
    browning(start_voltage, end_voltage, slew_rate, frequency)
    fixvoltage(end_voltage,time_fixvoltage)
    browning(end_voltage, start_voltage, slew_rate, frequency)

    discharge_output()


    soak(int(delay/2))
    scope.stop()

    filename = f"{start_voltage}-{end_voltage}-{start_voltage} Vac.png"
    scope.get_screenshot(filename, path)
    waveform_counter += 1
    print(filename)


def main():
    global waveform_counter
    discharge_output()
    scope_settings()
    operation()
        
if __name__ == "__main__":
    headers(test)
    main()
    footers(waveform_counter)
















