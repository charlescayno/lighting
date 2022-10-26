##################################################################################
"""COMMS"""
ac_source_address = 5
source_power_meter_address = 1 
load_power_meter_address = 2    
eload_address = 8
scope_address = "10.125.11.0"

"""USER INPUT"""
vout = 46
iout = 1.660
cr_load = vout/iout
cr_load = f'{cr_load:.4f}'

vin_list = [100,115,230,265,277,300]

cr_list = [30,60,90]
led_list = [46,36,24]

von = 0
component = "IBOOST"

test = "Component Stress Analysis (Final) - Iboost"
waveforms_folder = f'C:/Users/ccayno/Desktop/DER-945/waveforms/{test}'

"""DO NOT EDIT BELOW THIS LINE"""
##################################################################################

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

"""GENERIC FUNCTIONS"""
if not os.path.exists(waveforms_folder): os.mkdir(waveforms_folder)

def discharge_output():
    ac.turn_off()
    eload.channel[1].cc = 1
    eload.channel[1].turn_on()
    eload.channel[2].cc = 1
    eload.channel[2].turn_on()
    eload.channel[3].cc = 1
    eload.channel[3].turn_on()
    eload.channel[1].short_on()
    sleep(2)
    eload.channel[1].turn_off()
    eload.channel[2].turn_off()
    eload.channel[3].turn_off()
    eload.channel[1].short_off()
    sleep(1)


def scope_settings_for_startup_operation():

    print("Startup Operation")

    scope.stop()
    scope.position_scale(time_position = 10, time_scale = 1)
    scope.edge_trigger(1, 10, 'POS')
    
    if component == "IBOOST":
        scope.channel_settings(1, 5, -3, 'Iboost')
        scope.channel_state(2, 'OFF')
        scope.channel_state(3, 'OFF')
        scope.channel_state(4, 'OFF')
    
    sleep(2) 



def get_value(channel):
    labels, values = scope.get_measure(channel)
    max_value = float(values[0])
    max_value = float(f"{max_value:.2f}")
    return max_value


def export_scope_measurement(component, LED, vin, test_type):

    load = str(LED)
    input = str(vin)
    ch1_max = f"{get_value(1)}"
    # ch2_max = f"{get_value(2)}"
    # ch3_max = f"{get_value(3)}"
    # ch4_max = f"{get_value(4)}"
    
    output_list = [load, input, ch1_max]
    # print(','.join(output_list))

    with open(f'{component} - {LED}V at {test_type}.txt', 'a+') as f:
        f.write(','.join(output_list))
        f.write('\n')






def main():
    global waveform_counter
    test_type = "Startup"

    scope_settings_for_startup_operation()

    eload.channel[1].von = von


    for LED in led_list:

        prompt(f"Change LED load to {LED} Volts")
        if os.path.exists(f"{component} - {LED}V at {test_type}.txt"): os.remove(f"{component} - {LED}V at {test_type}.txt")

        for vin in vin_list:
            
            scope.run_single()

            sleep(5)

            ac.voltage = vin
            ac.frequency = ac.set_freq(vin)
            ac.turn_on()

            sleep(10)

            # filename = f"{vin}Vac, {cr}Ohms Eload CRH (VON={von}V) - Startup.png"
            filename = f"{vin}Vac, {LED}V - Startup.png"
            path = waveforms_folder + f'/LED Load'
            if not os.path.exists(path): os.mkdir(path)
            scope.get_screenshot(filename, path)
            waveform_counter += 1
            print(filename)
            export_scope_measurement(component, LED, vin, test_type)

            discharge_output()


if __name__ == "__main__":
    headers(test)
    discharge_output()
    main()
    discharge_output()
    footers(waveform_counter)
