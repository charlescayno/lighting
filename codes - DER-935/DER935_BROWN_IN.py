batch = input("Batch: ")
unit = input("Unit: ")
der = 935
print("Charles Cayno | 03-17-2022")
print()
print("Set the following:")
print("Channel 1 - Output Current")
print("Channel 4 - Input Voltage")
input("Press ENTER to continue...")
print()

test = f"Brown-in Brown-out"

# comms
ac_source_address = 5
source_power_meter_address = 1 
load_power_meter_address = 2
eload_address = 8
scope_address = "10.125.10.184"


## USER INPUT ###

start_voltage = 0
end_voltage = 120
slew_rate = 1
frequency = 60
time_fixvoltage = 80


#########################################################################
from time import time, sleep
import numpy as np
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

waveform_counter = 0
waveforms_folder = f'C:/Users/ccayno/Desktop/DER/DER-{der}/2 Marketing Sample Units/Batch {batch}/Unit {unit}/{test}/'
path = path_maker(f'{waveforms_folder}')



# # initialize equipment
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
    sleep(2)
    for i in range(1,9):
        eload.channel[i].turn_off()
        eload.channel[i].short_off()
    sleep(1)



def scope_settings():
    """CHANNEL SETTINGS"""
    scope.channel_settings(state='ON', channel=1, scale=0.2, position=-4, label="OUTPUT CURRENT", color='GREEN', rel_x_position=50, bandwidth=20, coupling='DCLimit', offset=0)
    # scope.channel_settings(state='OFF', channel=2, scale=2, position=0, label="", color='', rel_x_position=0, bandwidth=20, coupling='DCLimit', offset=0)
    # scope.channel_settings(state='OFF', channel=3, scale=1, position=0, label="", color='', rel_x_position=0, bandwidth=20, coupling='DCLimit', offset=0)
    scope.channel_settings(state='ON', channel=4, scale=100, position=0, label="INPUT VOLTAGE", color='YELLOW', rel_x_position=50, bandwidth=20, coupling='AC', offset=0)
    
    """MEASURE SETTINGS"""
    scope.measure(1, "MAX,MIN")
    # scope.measure(2, "MAX,MIN")
    # scope.measure(3, "MAX,MIN")
    scope.measure(4, "MAX,MIN")

    """HORIZONTAL SETTINGS"""
    # scope.time_position(10)
    scope.record_length(50E6)
    # scope.time_scale(1)

    """ZOOM SETTINGS"""
    # scope.remove_zoom()
    # scope.add_zoom(rel_pos=21.727, rel_scale=1)
    
    """TRIGGER SETTINGS"""
    trigger_channel = 4
    trigger_level = 0
    trigger_edge = 'POS'
    scope.edge_trigger(trigger_channel, trigger_level, trigger_edge)




def browning(start, end, slew, frequency):

    if start > end:
        print(f"brownout: {start} -> {end} Vac")
        for voltage in np.arange(start, end+1, -slew):
            ac.voltage = voltage
            ac.frequency = frequency
            ac.turn_on()
            sleep(1)
    if start < end:
        print(f"brownin: {start} -> {end} Vac")
        for voltage in np.arange(start, end+1, slew):
            ac.voltage = voltage
            ac.frequency = frequency
            ac.turn_on()
            sleep(1)

def fixvoltage(voltage, soak):
    ac.voltage = voltage
    ac.frequency = ac.set_freq(voltage)
    ac.turn_on()
    sleep(soak)

def roundup(x):
    return int(math.ceil(x / 100.0)) * 100

def soak(soak_time=30):
    for seconds in np.arange(soak_time, 0, -1):
        sleep(1)
        print(f"{seconds:5d}s", end="\r")
    print("       ", end="\r")


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