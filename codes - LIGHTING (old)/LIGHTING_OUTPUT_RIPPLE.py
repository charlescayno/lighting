"""LIBRARIES"""
from time import time, sleep
import sys
import os
from powi.equipment import ACSource, PowerMeter, ElectronicLoad, Oscilloscope
from powi.equipment import headers, create_folder, footers, waveform_counter, soak, convert_argv_to_int_list
waveform_counter = 0

"""COMMS"""
ac_source_address = 5
source_power_meter_address = 1 
load_power_meter_address = 2
eload_address = 8
scope_address = "10.125.10.148"

"""EQUIPMENT INITIALIZE"""
ac = ACSource(ac_source_address)
pms = PowerMeter(source_power_meter_address)
pml = PowerMeter(load_power_meter_address)
eload = ElectronicLoad(eload_address)
scope = Oscilloscope(scope_address)
# control = MultiprotocolControl()

"""USER INPUT"""
vin_list = [90, 115, 230, 277, 300]

vout = 46
iout = 1.05

trigger_channel = 1
trigger_delta = 0.003 # 3mV sensitivity - how reactive trigger automation

test = "Output Voltage Ripple"
waveforms_folder = f'waveforms/{test}'

"""DO NOT EDIT BELOW THIS LINE"""
CR_FL = float(f'{vout/iout:.2f}')
cr_list = [0.10*CR_FL, 0.25*CR_FL, 0.5*CR_FL, 0.75*CR_FL, CR_FL]
iout_name = [10, 25, 50, 75, 100]




"""GENERIC FUNCTIONS"""
def discharge_output():
    ac.turn_off()
    eload.channel[1].cr = 100
    eload.channel[2].cr = 100
    eload.channel[1].turn_on()
    eload.channel[2].turn_on()
    sleep(3)
    eload.channel[1].turn_off()
    eload.channel[2].turn_off()


def scope_settings_100ms():

    scope.add_zoom(rel_pos = 0, rel_scale = 0.5) # pa adjust na lang nito

    scope.position_scale(time_position = 0, time_scale = 0.1)
    scope.record_length(record_length = 20E6) # 20 MSa

    scope.channel_coupling(channel = trigger_channel, channel_coupling = 'AC')
    scope.channel_settings(channel = trigger_channel, scale = 0.5, position = 0)
    scope.channel_BW(channel = trigger_channel, channel_BW = 20) # 20 MHz

    scope.remove_zoom()


def scope_settings_5ms():

    scope.remove_zoom()

    scope.position_scale(time_position = 0, time_scale = 0.005)
    scope.record_length(record_length = 20E6) # 20 MSa

    scope.channel_coupling(channel = trigger_channel, channel_coupling = 'AC')
    scope.channel_settings(channel = trigger_channel, scale = 0.5, position = 0)
    scope.channel_BW(channel = trigger_channel, channel_BW = 20) # 20 MHz


def conv_str_to_float(string):
    temp = float(string)
    return float(f'{temp:.4f}')

def check_trigger_status():
    scope.run_single()
    sleep(3)
    return scope.trigger_status()

def find_trigger():

    scope.run_single()
    sleep(5)

    labels, values = scope.get_measure()
    """
    values[0] = max
    values[1] = min
    values[2] = peak to peak
    values[3] = mean
    values[4] = rms
    values[5] = frequency
    """

    max_value = conv_str_to_float(values[2])

    trigger_level = max_value
    scope.trigger_level(trigger_channel, trigger_level) # set max value as initial trigger level

    while (check_trigger_status() == 1): # increase trigger level until it reaches max trig level
        trigger_level += trigger_delta
        scope.trigger_level(trigger_channel, trigger_level)
    
    trigger_level -= 2*trigger_delta # decrease trigger level to get max trigger possible
    scope.trigger_level(trigger_channel, trigger_level) # final trigger level found!


def main():
    global waveform_counter
    iout_index = 0

    scope.edge_trigger(trigger_channel, 0.010, 'POS') # initial trigger

    for vin in vin_list:

        ac.voltage = vin
        ac.frequency = ac.set_freq(vin)
        ac.turn_on()

        for cr in cr_list:
            eload.channel[1].cr = cr
            eload.channel[1].turn_on()

            """100ms/div"""

            scope_settings_100ms()
            find_trigger()

            scope.run_single()
            sleep(6)

            filename = f'{vin}VAC, {iout_name[iout_index]}LoadCR, 100ms_per_div.png'
            scope.get_screenshot(filename, waveforms_folder)
            print(filename)
            waveform_counter += 1
    
            sleep(1)

            """5ms/div"""

            scope_settings_5ms()

            scope.run_single()
            sleep(6)

            filename = f'{vin}VAC, {iout_name[iout_index]}LoadCR, 5ms_per_div.png'
            scope.get_screenshot(filename, waveforms_folder)
            print(filename)
            waveform_counter += 1
            
            
            iout_index += 1
            scope.edge_trigger(trigger_channel, 0.010, 'POS') # reset trigger level
        
        iout_index = 0
        discharge_output()
        print()

if __name__ == "__main__":
    headers(test)
    main()
    discharge_output()
    footers(waveform_counter)