# November 13, 2021

"""IMPORT DEPENDENCIES"""
from time import time, sleep
import sys
import os
import math
from powi.equipment import ACSource, PowerMeter, ElectronicLoad, Oscilloscope, LEDControl
from powi.equipment import headers, create_folder, footers, waveform_counter, soak, convert_argv_to_int_list, tts, prompt, sfx
from filemanager import path_maker, remove_file, move_file
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

# 1 - Startup
# 2 - Normal
# 3 - Startup Short
# 4 - Running Short
test_list =  convert_argv_to_int_list(sys.argv[1])

test = "Drain Voltage and Drain Current"
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
    print(">> Set CH2 to Drain Voltage")
    print(">> Set CH3 to Drain Current")
    
def scope_settings():
    global ch1, ch2, ch3, ch4

    scope.stop()

    scope.remove_zoom()

    scope.position_scale(time_position = 50, time_scale = 0.02)
    scope.record_length(50E6)
    
    scope.edge_trigger(2, 10, 'POS')

    

    ch2 = scope.channel_settings(state='ON', channel=2, scale=200, position=0, label="Drain Voltage", color='YELLOW', rel_x_position=30, bandwidth=500, coupling='DCLimit', offset=0)
    ch3 = scope.channel_settings(state='ON', channel=3, scale=2, position=-4, label="Drain Current", color='LIGHT_BLUE', rel_x_position=70, bandwidth=500, coupling='DCLimit', offset=0)



def find_trigger(channel, trigger_delta):

    scope.edge_trigger(2, 0, 'POS')

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



def operation(test_type):

    global waveform_counter

    scope_settings()
    
    for LED in led_list:

        led.voltage(LED)

        for vin in vin_list:

            if test_type == 'Startup':
                scope.position_scale(time_position = 30, time_scale = 0.02)
                scope.add_zoom(35, 20)
                scope.edge_trigger(3, 2.8, 'POS')
                sleep(1)
                scope.run_single()
                sleep(3)
                ac.voltage = vin
                ac.frequency = ac.set_freq(vin)
                ac.turn_on()
                sleep(5)
                discharge_output()

                filename = f'Vds Ids {test_type} - {LED}V, {vin}Vac.png'
                path = path_maker(f'{waveforms_folder}/{test_type}')
                scope.get_screenshot(filename, path)
                print(filename)
                waveform_counter += 1
            
            if test_type == 'Normal':
                scope.position_scale(time_position = 50, time_scale = 0.01)
                scope.add_zoom(50, 0.1)
                ac.voltage = vin
                ac.frequency = ac.set_freq(vin)
                ac.turn_on()
                sleep(5)
                find_trigger(channel=2, trigger_delta=1)
                scope.run_single()
                sleep(3)
                discharge_output()
            
                filename = f'Vds Ids {test_type} - {LED}V, {vin}Vac.png'
                path = path_maker(f'{waveforms_folder}/{test_type}')
                scope.get_screenshot(filename, path)
                print(filename)
                waveform_counter += 1

            if test_type == 'Startup Short':
                timeperdiv = 2
                scope.position_scale(time_position = 10, time_scale = timeperdiv)
                scope.edge_trigger(3, 1, 'POS')
                ch2 = scope.channel_settings(state='ON', channel=2, scale=100, position=-4, label="Drain Voltage", color='YELLOW', rel_x_position=50, bandwidth=500, coupling='DCLimit', offset=0)
                ch3 = scope.channel_settings(state='ON', channel=3, scale=2, position=-4, label="Drain Current", color='LIGHT_BLUE', rel_x_position=50, bandwidth=500, coupling='DCLimit', offset=0)
                sleep(1)
                
                sfx()
                input(">> Short Output. Press ENTER to continue...")
                scope.run_single()
                sleep(3)
                ac.voltage = vin
                ac.frequency = ac.set_freq(vin)
                ac.turn_on()
                sleep(1)
                sleep(10*timeperdiv)
                discharge_output()

                filename = f'Vds Ids at Output Startup Short - {LED}V, {vin}Vac.png'
                path = path_maker(f'{waveforms_folder}/{test_type}')
                scope.get_screenshot(filename, path)
                print(filename)
                waveform_counter += 1

                scope.position_scale(time_position = 10, time_scale = 0.1)
                sfx()
                input(">> Remove Output Short. Press ENTER to continue...")
                scope.run_single()
                sleep(4)
                ac.voltage = vin
                ac.frequency = ac.set_freq(vin)
                ac.turn_on()
                sleep(4)
                discharge_output()

                filename = f'Vds Ids After Output Startup Short - {LED}V, {vin}Vac.png'
                path = path_maker(f'{waveforms_folder}/{test_type}')
                scope.get_screenshot(filename, path)
                print(filename)
                waveform_counter += 1
            
            if test_type == 'Running Short':
                timeperdiv = 2
                scope.position_scale(time_position = 20, time_scale = timeperdiv)
                
                scope.edge_trigger(3, 1, 'NEG')

                scope.timeout_trigger(3, timeout_range='LOW', timeout_time=2E-3)

                ch2 = scope.channel_settings(state='ON', channel=2, scale=200, position=0, label="Drain Voltage", color='YELLOW', rel_x_position=30, bandwidth=500, coupling='DCLimit', offset=0)
                ch3 = scope.channel_settings(state='ON', channel=3, scale=2, position=-4, label="Drain Current", color='LIGHT_BLUE', rel_x_position=70, bandwidth=500, coupling='DCLimit', offset=0)
                
                sleep(1)
                
                ac.voltage = vin
                ac.frequency = ac.set_freq(vin)
                ac.turn_on()
                sleep(2)

                scope.run_single()

                sleep(2*timeperdiv)

                print(">> Short Output.", end="\r")

                sleep(5*timeperdiv)

                print(">> Remove Output Short.", end="\r")

                sleep(5*timeperdiv)

                filename = f'Running Output Short - {LED}V, {vin}Vac.png'
                path = path_maker(f'{waveforms_folder}/{test_type}')
                scope.get_screenshot(filename, path)
                print(filename)
                waveform_counter += 1

                discharge_output()

def main():
    global waveform_counter

    test = 1 # startup
    if test in test_list: operation(test_type='Startup')

    test = 2 # normal
    if test in test_list: operation(test_type='Normal')

    test = 3 # startup short
    if test in test_list:
        input("Change saveset.")
        operation(test_type='Startup Short')
    
    test = 4 # running short
    if test in test_list:
        operation(test_type='Running Short')

        
if __name__ == "__main__":
    headers(test)
    discharge_output()
    led = LEDControl()
    main()
    discharge_output()
    footers(waveform_counter)