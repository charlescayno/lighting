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
scope_address = "10.125.11.0"

"""USER INPUT"""
led_list = [46]
vin_list = [100,115,120,230,265,277,300]
# test_list = convert_argv_to_int_list(sys.argv[1]) # 1 - Startup, 2 - Normal
test_list = [2]

test = "Iout Ripple"
waveforms_folder = f'C:/Users/ccayno/Desktop/DER-945/waveforms/{test}'

"""DO NOT EDIT BELOW THIS LINE"""
##################################################################################


"""EQUIPMENT INITIALIZE"""
ac = ACSource(ac_source_address)
pms = PowerMeter(source_power_meter_address)
pml = PowerMeter(load_power_meter_address)
eload = ElectronicLoad(eload_address)
scope = Oscilloscope(scope_address)
# led = LEDControl()

"""GENERIC FUNCTIONS"""

def discharge_output():
    ac.turn_off()
    for i in range(10):
        eload.channel[i].cc = 1
        eload.channel[i].turn_on()
        eload.channel[i].short_on()
    sleep(2)
    for i in range(10):
        eload.channel[i].turn_off()
        eload.channel[i].short_off()
    sleep(1)

def scope_settings(test_type):
    global ch1, ch2, ch3, ch4

    print(f'Test Type: {test_type}')
    
    scope.stop()
    scope.remove_zoom()
    ch1 = scope.channel_settings(state='OFF')
    ch2 = scope.channel_settings(state='OFF')
    ch3 = scope.channel_settings(state='OFF')
    ch4 = scope.channel_settings(state='OFF')

    scope.position_scale(time_position = 50, time_scale = 0.004)
    scope.edge_trigger(2, 1.6, 'POS')

    ch2 = scope.channel_settings(state='ON', channel=2, scale=0.4, position=-4, label='IOUT', color='GREEN', bandwidth=500, coupling='DC', offset=0)

    sleep(2)







# def get_value(channel):
#     labels, values = scope.get_measure(channel)
#     max_value = float(f"{values[0]:.2f}")
#     return max_value

# def export_scope_measurement_label(test, LED, test_type):
#     global ch1, ch2, ch3, ch4

#     output_list = ['LED', 'VIN']

#     if ch1 == 'ON': output_list.append("CH1")
#     if ch2 == 'ON': output_list.append("CH2")
#     if ch3 == 'ON': output_list.append("CH3")
#     if ch4 == 'ON': output_list.append("CH4")

#     with open(f'{test} - {LED}V at {test_type}.txt', 'a+') as f:
#         f.write(','.join(output_list))
#         f.write('\n')

# def export_scope_measurement(test, LED, vin, test_type):
#     global ch1, ch2, ch3, ch4

#     load = str(LED)
#     input = str(vin)
#     output_list = [load, input]

#     if ch1 == 'ON':
#         max = f"{get_value(1)}"
#         output_list.append(max)
#     if ch2 == 'ON':
#         max = f"{get_value(2)}"
#         output_list.append(max)
#     if ch3 == 'ON':
#         max = f"{get_value(3)}"
#         output_list.append(max)
#     if ch4 == 'ON':
#         max = f"{get_value(4)}"
#         output_list.append(max)

#     with open(f'{test} - {LED}V at {test_type}.txt', 'a+') as f:
#         f.write(','.join(output_list))
#         f.write('\n')












def operation(test_type):

    global waveform_counter

    scope_settings(test_type)
    
    for LED in led_list:
        # led.voltage(LED)
        prompt(f"Change LED load to {LED} Volts")
        
        # file = f'{test} - {LED}V at {test_type}.txt'
        # remove_file(file)
        # export_scope_measurement_label(test, LED, test_type)

        for vin in vin_list:

            if test_type == 'Startup':
                pass
            
            if test_type == 'Normal':
                ac.voltage = vin
                ac.frequency = ac.set_freq(vin)
                ac.turn_on()
                sleep(5)
                scope.run()
                input("Adjust cursor then press enter to continue.")
                scope.stop()
                sleep(1)
                discharge_output()

            filename = f'{test} {LED}V, {vin}Vac.png'
            path = path_maker(f'{waveforms_folder}/{LED}V')
            scope.get_screenshot(filename, path)
            print(filename)
            waveform_counter += 1

        #     export_scope_measurement(test, LED, vin, test_type)

        # # transporting txt file from codes folder to waveforms folder
        # import shutil
        # source = f'{os.getcwd()}/{file}'
        # destination = f'{path}/{file}'
        # remove_file(destination)
        # shutil.move(source, destination)

def main():
    global waveform_counter

    test = 1 # startup
    if test in test_list: operation(test_type='Startup')

    test = 2 # normal
    if test in test_list: operation(test_type='Normal')
        
if __name__ == "__main__":
    headers(test)
    discharge_output()
    main()
    discharge_output()
    footers(waveform_counter)