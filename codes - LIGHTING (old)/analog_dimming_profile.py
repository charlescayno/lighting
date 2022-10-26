
# COMMS
ac_source_address = 5
source_power_meter_address = 1
load_power_meter_address = 2
eload_address = 8
scope_address = "10.125.11.0"

"""IMPORT DEPENDENCIES"""
import sys
import pyautogui
from time import sleep, time
from powi.equipment import ACSource, PowerMeter, ElectronicLoad, Oscilloscope
from powi.equipment import headers, create_folder, footers, waveform_counter, soak, convert_argv_to_int_list, prompt
import os
from powi.equipment import LEDControl
waveform_counter = 0

"""INITIALIZE EQUIPMENT"""
ac = ACSource(ac_source_address)
# pms = PowerMeter(source_power_meter_address)
# pml = PowerMeter(load_power_meter_address)
eload = ElectronicLoad(eload_address)
# scope = Oscilloscope(scope_address)
led = LEDControl()

def discharge_output():
    ac.turn_off()
    eload.channel[1].cr = 100
    eload.channel[1].turn_on()
    eload.channel[2].cr = 100
    eload.channel[2].turn_on()
    eload.channel[3].cr = 100
    eload.channel[3].turn_on()
    
    sleep(1)
    
    eload.channel[1].turn_off()
    eload.channel[2].turn_off()
    eload.channel[3].turn_off()

    eload.channel[1].cc = 1
    eload.channel[1].turn_on()
    eload.channel[2].cc = 1
    eload.channel[2].turn_on()
    eload.channel[3].cc = 1
    eload.channel[3].turn_on()
    
    sleep(1)

    eload.channel[1].turn_off()
    eload.channel[2].turn_off()
    eload.channel[3].turn_off()

    sleep(1)





from AUTOGUI_CONFIG import *
from time import sleep, time
import os.path
from os import path
import math





### USER INPUT #######################################
led_list = [46,36,24]
led_list = convert_argv_to_int_list(sys.argv[1])

vin_list = [120,230,277]
vin_list = convert_argv_to_int_list(sys.argv[2])

test_list = [1,2] # 1 - min to max, 2 - max to min
test_list = convert_argv_to_int_list(sys.argv[3])

test = f"Analog Dimming"
waveforms_folder = f'waveforms/{test}'
# vin_list = [230]
######################################################



ate_gui = AutoguiCalibrate()


def local_calibration():
    vin_dict = {}
    for LED in led_list:
        for vin in vin_list:
            vin_dict[f"{LED}V {vin}Vac"] = ate_gui.get_coordinates_manual(f"{LED}V {vin}Vac")

    with open('local_coordinates.txt', 'w') as f:
        f.write(str(vin_dict))

    return vin_dict

def load_local_calibration():
    with open('local_coordinates.txt', 'r') as f:
        str_dict = f.read()
        vin_dict = eval(str_dict)

    return vin_dict

def analog_dimming(LED, vin, start, end, step, soak, delay_per_line, delay_per_load):

    global waveform_counter

    # scope.stop()
    sleep(1)
    # scope.run()
    discharge_output()

    timer = int( abs((end-start)/step) * 13)
    time_scale = math.ceil(timer)/10
    # scope.time_scale(time_scale)

    ate_gui.initialize_analog_dimming(start)
    ate_gui.ANALOG_DIMMING(start=start, end=end, step=step, soak=soak, delay_per_line=delay_per_line, delay_per_load = delay_per_load)
    ate_gui.ac_turn_on(vin)
    ate_gui.run_test()

    sleep(timer)
    sleep(30)
    # prompt("Press ENTER to continue")
    # ate_gui.alt_tab()
    # sleep(2)
    # scope.stop()


    filename = f'{LED}V, {vin}Vac, {start}-{end}V, {step}V increment.png'
    path = waveforms_folder + f'/Profiles'
    if not os.path.exists(path): os.mkdir(path)
    # scope.get_screenshot(filename, path)
    print(filename)
    waveform_counter += 1

    ate_gui.test_complete()
    discharge_output()



def main():
    
    # loading coordinates for the autogui
    if os.path.isfile('local_coordinates.txt'): vin_dict = load_local_calibration()
    else: vin_dict = local_calibration()

    # asking the user if calibration is needed
    recalibrate_status = 0
    while recalibrate_status != 'y':
        recalibrate_status = input("Recalibrate? (y/n): ")
        if recalibrate_status == 'y': vin_dict = local_calibration()
        elif recalibrate_status == 'n': break

    ate_gui.alt_tab()
    ate_gui.esc()

    test = 1
    if test in test_list:
        if len(test_list) > 1: input("Change ATE template in use.")
        for LED in led_list:
            led.voltage(LED)
            for vin in vin_list:    
                """MIN to MAX"""
                ate_gui.click(vin_dict[f"{LED}V {vin}Vac"])
                # analog_dimming(LED = LED, vin = vin, start = 0, end = 10, step = 0.1, soak = 10, delay_per_line = 10, delay_per_load = 5)
                analog_dimming(LED = LED, vin = vin, start = 0, end = 10, step = 0.1, soak = 10, delay_per_line = 10, delay_per_load = 5)

    test = 2
    if test in test_list:
        if len(test_list) > 1: input("Change ATE template in use.")
        for LED in led_list:
            led.voltage(LED)
            for vin in vin_list:    
                """MAX to MIN"""
                ate_gui.click(vin_dict[f"{LED}V {vin}Vac"])
                # analog_dimming(LED = LED, vin = vin, start = 10, end = 0, step = 0.1, soak = 10, delay_per_line = 10, delay_per_load = 5)
                analog_dimming(LED = LED, vin = vin, start = 10, end = 0, step = 0.1, soak = 10, delay_per_line = 10, delay_per_load = 5)
    
    ate_gui.alt_tab()
    





if __name__ == "__main__":
    headers(test)
    main()
    footers(waveform_counter)