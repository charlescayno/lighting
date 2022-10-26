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

##################################################################################

"""COMMS"""
ac_source_address = 5
source_power_meter_address = 1 
load_power_meter_address = 2
eload_address = 8
scope_address = "10.125.10.115"

"""USER INPUT"""
led_list = [46,36,24]
vin_list = [100,115,230,265,277,300]


test = "Input Voltage and Input Current"
waveforms_folder = f'C:/Users/ccayno/Desktop/DER-945/Test Data/{test}'

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
    for i in range(1,9):
        eload.channel[i].cc = 1
        eload.channel[i].turn_on()
        eload.channel[i].short_on()
    sleep(2)
    for i in range(1,9):
        eload.channel[i].turn_off()
        eload.channel[i].short_off()
    sleep(1)


def message():
    print(">> Set CH1 to Input Current")
    print(">> Set CH4 to Input Voltage")
    print("ETC: ~ 9.5 mins.")
    input("Press ENTER to continue")


def scope_settings():
    global ch1, ch2, ch3, ch4
    global channel_trigger, channel_trigger_delta

    scope.stop()
    scope.remove_zoom()

    ch1 = scope.channel_settings(state='OFF')
    ch2 = scope.channel_settings(state='OFF')
    ch3 = scope.channel_settings(state='OFF')
    ch4 = scope.channel_settings(state='OFF')

    scope.position_scale(time_position = 50, time_scale = 0.01)
    scope.edge_trigger(4, 120, 'POS')

    ch1 = scope.channel_settings(state='ON', channel=1, scale=1, position=2, label="Input Current", color='ORANGE', rel_x_position=30, bandwidth=500, coupling='DCLimit', offset=0)
    ch4 = scope.channel_settings(state='ON', channel=4, scale=200, position=-2, label="Input Voltage", color='PINK', rel_x_position=70, bandwidth=500, coupling='AC', offset=0)


def find_trigger(channel, trigger_delta):

    # finding trigger level
    scope.run_single()
    soak(5)

    # get initial peak-to-peak measurement value
    labels, values = scope.get_measure(channel)
    max_value = float(values[0])
    max_value = float(f"{max_value:.4f}")

    # set max_value as initial trigger level
    trigger_level = max_value
    scope.edge_trigger(channel, trigger_level, 'POS')

    # check if it triggered within 5 seconds
    scope.run_single()
    soak(3)
    trigger_status = scope.trigger_status()

    # increase trigger level until it reaches the maximum trigger level
    while (trigger_status == 1):
        trigger_level += trigger_delta
        scope.edge_trigger(channel, trigger_level, 'POS')

        # check trigger status
        scope.run_single()
        soak(3)
        trigger_status = scope.trigger_status()

    # decrease trigger level below to get the maximum trigger possible
    trigger_level -= 2*trigger_delta
    scope.edge_trigger(channel, trigger_level, 'POS')
    # print(f'Maximum trigger level found at: {trigger_level}')

def operation():

    global waveform_counter
    global channel_trigger, channel_trigger_delta

    scope_settings()
    
    for LED in led_list:

        led.voltage(LED)
        
        for vin in vin_list:

                ac.voltage = vin
                ac.frequency = ac.set_freq(vin)
                ac.turn_on()

                sleep(3)

                scope.edge_trigger(4, 0, 'POS')
                find_trigger(channel=4, trigger_delta=2)
                scope.run_single()
                sleep(6)

                filename = f'Input Current and Input Voltage - {LED}V, {vin}Vac.png'
                path = path_maker(f'{waveforms_folder}/{LED}V')
                scope.get_screenshot(filename, path)
                print(filename)
                waveform_counter += 1
                
                discharge_output()


def main():
    global waveform_counter
    operation()
        
if __name__ == "__main__":
    message()
    headers(test)
    discharge_output()
    led = LEDControl()
    main()
    discharge_output()
    footers(waveform_counter)