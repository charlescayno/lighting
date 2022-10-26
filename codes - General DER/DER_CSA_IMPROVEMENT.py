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
led_list = [46]
vin_list = [277,300]
test_list = convert_argv_to_int_list(sys.argv[1]) # 1 - Startup, 2 - Normal, 3 - Normal-Zoomed 
# test_list = [1,2]

component = "VDS"
# component = "BoostFET"
# component = "Boost Diode"
# component = "PassFET"
# component = "Output Rectifier Diode"

# condition = "Charles reworks"
condition = "RT-ZP10"
print(condition)

test = "CSA - Improvements"
waveforms_folder = f'C:/Users/ccayno/Desktop/DER-945/waveforms/{test}'

"""DO NOT EDIT BELOW THIS LINE"""
##################################################################################


"""EQUIPMENT INITIALIZE"""
ac = ACSource(ac_source_address)
pms = PowerMeter(source_power_meter_address)
pml = PowerMeter(load_power_meter_address)
eload = ElectronicLoad(eload_address)
scope = Oscilloscope(scope_address)
led = LEDControl()

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



def scope_settings(test_type):
    global ch1, ch2, ch3, ch4
    global channel_trigger, channel_trigger_delta


    print(f'Test Type: {test_type}')
    
    scope.stop()
    scope.remove_zoom()

    ch1 = scope.channel_settings(state='OFF')
    ch2 = scope.channel_settings(state='OFF')
    ch3 = scope.channel_settings(state='OFF')
    ch4 = scope.channel_settings(state='OFF')


    if component == "VDS":
        if test_type == 'Startup':
            scope.position_scale(time_position = 10, time_scale = 0.2)
            scope.edge_trigger(2, 100, 'POS')
            # scope.position_scale(time_position = 10, time_scale = 200E-9)
            # scope.edge_trigger(2, 700, 'POS')
        if test_type == 'Normal':
            scope.position_scale(time_position = 50, time_scale = 500E-6)
            scope.edge_trigger(2, 500, 'POS')
            # scope.position_scale(time_position = 50, time_scale = 0.001)
            # scope.edge_trigger(2, 100, 'POS')
            scope.add_zoom(rel_pos = 50, rel_scale = 1.2)
            
            channel_trigger = 2
            channel_trigger_delta = 1

        if test_type == 'Normal_Zoomed':
            scope.position_scale(time_position = 50, time_scale = 200E-9)
            scope.edge_trigger(2, 500, 'POS')
            
            channel_trigger = 2
            channel_trigger_delta = 1

        ch2 = scope.channel_settings(state='ON', channel=2, scale=100, position=-4, label=component, color='YELLOW', bandwidth=500, coupling='DCLimit', offset=0)
        # ch3 = scope.channel_settings(state='ON', channel=3, scale=100, position=-4, label='VCSN', color='PINK', bandwidth=500, coupling='DCLimit', offset=0)

    if component == 'BoostFET':
        if test_type == 'Startup':
            scope.position_scale(time_position = 10, time_scale = 0.2)
            scope.edge_trigger(2, 10, 'POS')
        if test_type == 'Normal':
            scope.position_scale(time_position = 50, time_scale = 0.010)
            scope.edge_trigger(2, 50, 'POS')
            scope.add_zoom(rel_pos = 50, rel_scale = 0.1)

            channel_trigger = 2
            channel_trigger_delta = 0.5
        
        ch2 = scope.channel_settings(state='ON', channel=2, scale=20, position=-4, label=component, color='LIGHT_BLUE', bandwidth=500, coupling='DCLimit', offset=0)
    
    if component == 'Boost Diode':
        if test_type == 'Startup':
            scope.position_scale(time_position = 10, time_scale = 0.2)
            scope.edge_trigger(2, 10, 'POS')
        if test_type == 'Normal':
            scope.position_scale(time_position = 50, time_scale = 0.010)
            scope.edge_trigger(2, 62.5, 'POS')
            scope.add_zoom(rel_pos = 50, rel_scale = 0.1)

            channel_trigger = 2
            channel_trigger_delta = 0.5
        
        ch2 = scope.channel_settings(state='ON', channel=2, scale=20, position=-4, label=component, color='GREEN', bandwidth=500, coupling='DCLimit', offset=0)
    
    if component == 'PassFET':
        if test_type == 'Startup':
            scope.position_scale(time_position = 10, time_scale = 0.2)
            scope.edge_trigger(2, 10, 'POS')
        if test_type == 'Normal':
            scope.position_scale(time_position = 50, time_scale = 0.01)
            scope.edge_trigger(2, 30, 'POS')
            scope.add_zoom(rel_pos = 50, rel_scale = 0.2)

            channel_trigger = 2
            channel_trigger_delta = 0.5
        
        ch2 = scope.channel_settings(state='ON', channel=2, scale=20, position=-4, label=component, color='BLUE', bandwidth=500, coupling='DCLimit', offset=0)

    if component == 'Output Rectifier Diode':
        if test_type == 'Startup':
            scope.position_scale(time_position = 10, time_scale = 0.2)
            scope.edge_trigger(2, 100, 'POS')
        if test_type == 'Normal':
            scope.position_scale(time_position = 50, time_scale = 0.010)
            scope.edge_trigger(2, 50, 'POS')
            scope.add_zoom(rel_pos = 50, rel_scale = 0.1)

            channel_trigger = 2
            channel_trigger_delta = 0.5
        
        ch2 = scope.channel_settings(state='ON', channel=2, scale=100, position=-4, label=component, color='PINK', bandwidth=500, coupling='DCLimit', offset=0)
        input("Invert channel.")
    
    sleep(2) 


def get_value(channel):
    labels, values = scope.get_measure(channel)
    max_value = float(f"{values[0]:.2f}")
    return max_value

def export_scope_measurement_label(component, LED, test_type, time_scale):
    global ch1, ch2, ch3, ch4

    output_list = ['LED', 'VIN']

    if ch1 == 'ON': output_list.append("CH1")
    if ch2 == 'ON': output_list.append("CH2")
    if ch3 == 'ON': output_list.append("CH3")
    if ch4 == 'ON': output_list.append("CH4")

    with open(f'{component} - {LED}V at {test_type} - {time_scale}s perdiv.txt', 'a+') as f:
        f.write(','.join(output_list))
        f.write('\n')

def export_scope_measurement(component, LED, vin, test_type, time_scale):
    global ch1, ch2, ch3, ch4

    load = str(LED)
    input = str(vin)
    output_list = [load, input]

    if ch1 == 'ON':
        max = f"{get_value(1)}"
        output_list.append(max)
    if ch2 == 'ON':
        max = f"{get_value(2)}"
        output_list.append(max)
    if ch3 == 'ON':
        max = f"{get_value(3)}"
        output_list.append(max)
    if ch4 == 'ON':
        max = f"{get_value(4)}"
        output_list.append(max)

    with open(f'{component} - {LED}V at {test_type} - {time_scale}s perdiv.txt', 'a+') as f:
        f.write(','.join(output_list))
        f.write('\n')

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


def operation(test_type):

    global waveform_counter
    global channel_trigger, channel_trigger_delta

    scope_settings(test_type)
    
    for LED in led_list:

        led.voltage(LED)

        time_scale = scope.get_horizontal()['scale']
        file = f'{component} - {LED}V at {test_type} - {time_scale}s perdiv.txt'
        remove_file(file)
        export_scope_measurement_label(component, LED, test_type, time_scale)
        
        for vin in vin_list:

            if vin == 300 or vin == 277: count = 1
            else: count = 1

            for i in range(count):

                if test_type == 'Startup':
                    scope.run_single()
                    sleep(3)
                    ac.voltage = vin
                    ac.frequency = ac.set_freq(vin)
                    ac.turn_on()
                    sleep(5)
                    discharge_output()
                
                if test_type == 'Normal' or test_type == 'Normal_Zoomed':
                    ac.voltage = vin
                    ac.frequency = ac.set_freq(vin)
                    ac.turn_on()
                    sleep(5)
                    find_trigger(channel=channel_trigger, trigger_delta=channel_trigger_delta)
                    scope.run_single()
                    sleep(3)
                    discharge_output()

                filename = f'{component} {LED}V, {vin}Vac, {test_type}_{i}, {time_scale}s perdiv.png'
                path = path_maker(f'{waveforms_folder}/{component}/{condition}/{test_type}/{LED}V')
                scope.get_screenshot(filename, path)
                print(filename)
                waveform_counter += 1

                export_scope_measurement(component, LED, vin, test_type, time_scale)
        
        # transporting txt file from codes folder to waveforms folder
        import shutil
        source = f'{os.getcwd()}/{file}'
        destination = f'{path}/{file}'
        remove_file(destination)
        shutil.move(source, destination)



def main():
    global waveform_counter

    test = 1 # startup
    if test in test_list: operation(test_type='Startup')

    test = 2 # normal
    if test in test_list: operation(test_type='Normal')

    test = 3 # normal_zoomed
    if test in test_list: operation(test_type='Normal_Zoomed')
        
if __name__ == "__main__":
    headers(test)
    discharge_output()
    main()
    discharge_output()
    footers(waveform_counter)