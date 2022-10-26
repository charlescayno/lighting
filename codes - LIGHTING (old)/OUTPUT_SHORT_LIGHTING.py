print("Charles Cayno | 22-Jun-2021")

"""COMMS ADDRESS"""
ac_source_address = 5
source_power_meter_address = 2 
load_power_meter_address = 1
eload_address = 8
scope_address = "10.125.10.170"

"""IMPORT DEPENDENCIES"""
import sys
import pyautogui
from time import sleep, time
from powi.equipment import ACSource, PowerMeter, ElectronicLoad, Oscilloscope
from powi.equipment import headers, create_folder, footers, waveform_counter, soak, convert_argv_to_int_list
from powi.equipment import *
from time import sleep, time
import os
import winsound as ws

"""INITIALIZE EQUIPMENT"""
ac = ACSource(ac_source_address)
pms = PowerMeter(source_power_meter_address)
pml = PowerMeter(load_power_meter_address)
eload = ElectronicLoad(eload_address)
scope = Oscilloscope(scope_address)

"""USER INPUT"""

LED = sys.argv[1] # 46, 36, 24V
vin = convert_argv_to_int_list(sys.argv[2]) # [90, 115, 230, 265, 277, 300] Vac
test_list = convert_argv_to_int_list(sys.argv[3]) # 1, 2
timeperdiv = 3

# Probes
print("CH1: IDS")
print("CH2: VR")
print("CH3: VOUT")
print("CH4: IOUT")
vout_channel = 3
vout_trigger = 0
iout_channel = 4
iout_trigger = 0

# Folder Names
if LED == 'NL': test = f"Output Short at NL"
else: test = f"Output Short LED={LED}V"
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

scope.time_scale(timeperdiv)
scope.time_position(10)


ids_channel = 1
ids_trigger = 0.5
vr_channel = 2
vr_trigger = 0

scope.channel_scale(ids_channel, 2)
scope.channel_scale(vr_channel, 10)
scope.channel_scale(vout_channel, 10)
scope.channel_scale(iout_channel, 3)

scope.channel_position(ids_channel, -3)
scope.channel_position(vr_channel, -2)
scope.channel_position(vout_channel, -1)
scope.channel_position(iout_channel, -4)

scope.init_trigger(3, 40, 'NEG')

sleep(2)

"""Startup Short"""
test = 1
if test in test_list:
    print("Startup Short")
    scope.stop()

    ids_trigger = 0.05
    scope.edge_trigger(ids_channel, ids_trigger, 'POS')

    discharge_output()

    for voltage in vin:
        if voltage == 230: frequency = 50
        elif voltage == 265: frequency = 50
        else: frequency = 60
        
        scope.run_single()

        sleep(4)

        ac.voltage = voltage
        ac.frequency = frequency
        ac.turn_on()

        soak(9*timeperdiv+timeperdiv) 

        discharge_output()
        
        input("Press ENTER to capture waveform.")

        if LED != 'NL': filename = f"{voltage}Vac {frequency}Hz, {LED}V, Startup Output Short.png"
        else: filename = f"{voltage}Vac {frequency}Hz, NL, Startup Output Short.png"
        scope.get_screenshot(filename, waveforms_folder)
        print(filename)

        print()


"""Running Short"""
test = 2 
if test in test_list: 

    # input("Remove Output Short")

    scope.time_scale(3)
    trigger_source = 3
    scope.channel_scale(3,10)
    scope.channel_scale(iout_channel, 20)

    if LED == '46': trigger_level = 40
    elif LED == '36': trigger_level = 30
    elif LED == '24': trigger_level = 22
    else:
        trigger_level = 20
        trigger_source = 2
        scope.channel_scale(3,20)
        
    trigger_edge = 'NEG'
    scope.edge_trigger(trigger_source,trigger_level,trigger_edge)
    
    discharge_output()
    
    for voltage in vin:

        if voltage == 230: frequency = 50
        elif voltage == 265: frequency = 50
        else: frequency = 60

        ac.voltage = voltage
        ac.frequency = frequency
        ac.turn_on()

        scope.run_single()

        if LED == 'NL': soak(15)
        else: soak(5)

        # relay_ON()
        input("Short Output...")
        
        # if LED == 'NL': sleep(3)
        # else: sleep(6)

        # relay_OFF()
        # sleep(5)
        # print("Un-Short Output...")
        
        
        input("Press ENTER to capture waveform")
        discharge_output()

        if LED != 'NL': filename = f"{voltage}Vac {frequency}Hz, {LED}V, Running Output Short.png"
        else: filename = f"{voltage}Vac {frequency}Hz, NL, Running Output Short.png"
        scope.get_screenshot(filename, waveforms_folder)
        print(filename)
        print()
        # relay_OFF()
        discharge_output()


"Running Short (Onset of Short)"

test = 3 
if test in test_list: 

    scope.time_scale(0.001)
    trigger_source = 3
    scope.channel_scale(3,10)

    scope.channel_scale(ids_channel, 0.5)
    scope.channel_scale(vr_channel, 10)
    scope.channel_scale(vout_channel, 10)
    scope.channel_scale(iout_channel, 20)

    scope.time_position(20)
    scope.resolution(50E6)

    scope.channel_scale(3,20)

    if LED == '46': pass
    elif LED == '36': scope.channel_scale(ids_channel, 2)
    elif LED == '24': scope.channel_scale(ids_channel, 1)
    else: scope.channel_scale(ids_channel, 0.5)

    trigger_level = 10
    scope.edge_trigger(iout_channel,trigger_level,'POS')
    
    discharge_output()
    
    for voltage in vin:

        if voltage == 230: frequency = 50
        elif voltage == 265: frequency = 50
        else: frequency = 60

        ac.voltage = voltage
        ac.frequency = frequency
        ac.turn_on()

        if LED == 'NL': sleep(5)

        scope.run_single()

        input("Short Output...")
        
        input("Press ENTER to capture waveform")
        discharge_output()

        if LED != 'NL': filename = f"{voltage}Vac {frequency}Hz, {LED}V, Running Output Short (Onset of Short).png"
        else: filename = f"{voltage}Vac {frequency}Hz, NL, Running Output Short (Onset of Short).png"
        scope.get_screenshot(filename, waveforms_folder)
        print(filename)
        print()
        discharge_output()




test = 4 
if test in test_list: 

    # input("short output")

    scope.channel_scale(ids_channel, 2)
    scope.channel_scale(vr_channel, 10)
    scope.channel_scale(vout_channel, 10)
    scope.channel_scale(iout_channel, 20)

    scope.time_scale(0.001)
    scope.time_position(50)
    scope.resolution(50E6)

    trigger_source = 3
    trigger_level = 2        
    trigger_edge = 'POS'
    
    scope.edge_trigger(trigger_source,trigger_level,trigger_edge)
    
    discharge_output()
    
    for voltage in vin:

        if voltage == 230: frequency = 50
        elif voltage == 265: frequency = 50
        else: frequency = 60

        ac.voltage = voltage
        ac.frequency = frequency
        ac.turn_on()

        scope.run_single()

        sleep(2)

        input("remove short")      
        
        input("Press ENTER to capture waveform")

        if LED != 'NL': filename = f"{voltage}Vac {frequency}Hz, {LED}V, Running Output Short (Onset of Recovery).png"
        else: filename = f"{voltage}Vac {frequency}Hz, NL, Running Output Short (Onset of Recovery).png"
        scope.get_screenshot(filename, waveforms_folder)
        print(filename)
        print()
        
        
        discharge_output()
















footers(waveform_counter)


ws.PlaySound("dingding.wav", ws.SND_ASYNC)
sleep(2)