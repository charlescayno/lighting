batch = input("Batch: ")
unit = input("Unit: ")
print("Charles Cayno | 26-Oct-2021")

"""COMMS ADDRESS"""
ac_source_address = 5
source_power_meter_address = 2 
load_power_meter_address = 1
eload_address = 8
scope_address = "10.125.10.184"

"""IMPORT DEPENDENCIES"""
import sys
import pyautogui
from time import sleep, time
from powi.equipment import ACSource, PowerMeter, ElectronicLoad, Oscilloscope
from powi.equipment import headers, create_folder, footers, waveform_counter, soak, convert_argv_to_int_list
from powi.equipment import *
from powi.equipment import LEDControl
from time import sleep, time
import os
import winsound as ws
waveform_counter = 0

"""INITIALIZE EQUIPMENT"""
ac = ACSource(ac_source_address)
pms = PowerMeter(source_power_meter_address)
pml = PowerMeter(load_power_meter_address)
eload = ElectronicLoad(eload_address)
scope = Oscilloscope(scope_address)
# led = LEDControl()

"""USER INPUT"""

# LED = sys.argv[1] # 46, 36, 24V
LED = 48
# vin = convert_argv_to_int_list(sys.argv[2]) # [90, 115, 230, 265, 277, 300] Vac
vin = [100,115,230,277,300]
# vin = [100]
test_list = [1,2]

timeperdiv = 2

# Folder Names
if LED == 'NL': test = f"Output Short at NL"
else: test = f"Output Short LED={LED}V (Batch {batch} - Unit {unit})"
waveforms_folder = f'waveforms/{test}'

"""DEFAULT FUNCTIONS"""

def discharge_output():
    ac.turn_off()
    eload.channel[1].cc = 1
    eload.channel[1].turn_on()
    eload.channel[2].cc = 1
    eload.channel[2].turn_on()
    eload.channel[3].cc = 1
    eload.channel[3].turn_on()
    sleep(2)
    eload.channel[1].turn_off()
    eload.channel[2].turn_off()
    eload.channel[3].turn_off()

"""CUSTOM FUNCTIONS FOR THIS TEST"""

"""MAIN"""
headers(test)

# led.voltage(LED)
scope.time_scale(timeperdiv)
scope.time_position(10)

# scope.init_trigger(2, 0.2, 'POS')

sleep(2)

"""Startup Short"""
test = 1
if test in test_list:
    print("Startup Short")
    scope.stop()
    ids_trigger = 0.05
    # scope.edge_trigger(3, 50, 'POS')
    discharge_output()
    print()

    for voltage in vin:
        if voltage == 230: frequency = 50
        elif voltage == 265: frequency = 50
        else: frequency = 60
        
        input(">> Short Output. Press ENTER to continue...")

        scope.run_single()
        sleep(4)
        ac.voltage = voltage
        ac.frequency = frequency
        ac.turn_on()
        soak(9*timeperdiv+timeperdiv) 
        discharge_output()
        # input("Press ENTER to capture waveform.")
        if LED != 'NL': filename = f"{voltage}Vac {frequency}Hz, {LED}V, Startup Output Short.png"
        else: filename = f"{voltage}Vac {frequency}Hz, NL, Startup Output Short.png"
        scope.get_screenshot(filename, waveforms_folder)
        waveform_counter+=1
        print(filename)


        input(">> Remove Output Short. Press ENTER to continue...")

        scope.run_single()
        sleep(4)
        ac.voltage = voltage
        ac.frequency = frequency
        ac.turn_on()
        soak(9*timeperdiv+timeperdiv) 
        discharge_output()
        # input("Press ENTER to capture waveform.")
        if LED != 'NL': filename = f"{voltage}Vac {frequency}Hz, {LED}V, Startup Output After Short.png"
        else: filename = f"{voltage}Vac {frequency}Hz, NL, Startup Output After Short.png"
        scope.get_screenshot(filename, waveforms_folder)
        waveform_counter+=1
        print(filename)


        print()


"""Running Short"""
test = 2
if test in test_list:
    input("Press ENTER to continue test to: Running Short")
    scope.stop()
    discharge_output()
    print()

    for voltage in vin:
        if voltage == 230: frequency = 50
        elif voltage == 265: frequency = 50
        else: frequency = 60
        
        
        ac.voltage = voltage
        ac.frequency = frequency
        ac.turn_on()
        
        scope.run()
        sleep(4)
        
        soak(2*timeperdiv) 
        
        print(">> Short Output...")
        soak(4*timeperdiv)
        print(">> Remove Short...")
        soak(4*timeperdiv)
        
        scope.stop()



        # input("Press ENTER to capture waveform.")
        if LED != 'NL': filename = f"{voltage}Vac {frequency}Hz, {LED}V, Running Output Short.png"
        else: filename = f"{voltage}Vac {frequency}Hz, NL, Running Output Short.png"
        scope.get_screenshot(filename, waveforms_folder)
        waveform_counter+=1
        print(filename)

        discharge_output()
        print()

footers(waveform_counter)


ws.PlaySound("dingding.wav", ws.SND_ASYNC)
sleep(2)