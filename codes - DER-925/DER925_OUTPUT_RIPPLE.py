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
from powi.equipment import LEDControl
from powi.equipment import *
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
led_list = [48,24]
# vin = convert_argv_to_int_list(sys.argv[2]) # [90, 115, 230, 265, 277, 300] Vac
vin_list = [100,115,230,265,277,300]
timeperdiv = 0.004 #s

# Folder Names
test = f"Output Ripple (Batch {batch} - Unit {unit})"
waveforms_folder = f'waveforms/{test}'

"""DEFAULT FUNCTIONS"""

def discharge_output():
    ac.turn_off()
    for i in range(1,9):
        eload.channel[i].cc = 1
        eload.channel[i].turn_on()
    sleep(2)
    for i in range(1,9):
        eload.channel[i].turn_off()

"""CUSTOM FUNCTIONS FOR THIS TEST"""

"""MAIN"""
headers(test)

scope.time_scale(timeperdiv)
scope.time_position(50)


sleep(2)

for LED in led_list:
    # led.voltage(LED)
    input(f"Change LED load to {LED}V")
    for voltage in vin_list:
        ac.voltage = voltage
        ac.frequency = ac.set_freq(voltage)
        ac.turn_on()

        scope.run()

        input("Press ENTER to capture waveform..")

        scope.stop()

        sleep(1)
        
        filename = f"{voltage}Vac, {LED}V, After Output Short and Brownin.png"
        scope.get_screenshot(filename, waveforms_folder)
        waveform_counter+=1
        print(filename)

        discharge_output()


footers(waveform_counter)


ws.PlaySound("dingding.wav", ws.SND_ASYNC)
sleep(2)