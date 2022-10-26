#### DISCLAIMER: Largely dependent on pyautogui to change analog dimming voltage
#### Newly improved code 

test = "Analog Dimming"

# comms
ac_source_address = 5
source_power_meter_address = 1 
load_power_meter_address = 2
eload_address = 8
scope_address = "10.125.10.166"


############################################################################
import sys
import pyautogui
from time import sleep, time
from powi.equipment import ACSource, PowerMeter, ElectronicLoad, Oscilloscope
from powi.equipment import headers, create_folder, footers, waveform_counter
from time import sleep, time
import os
import math
waveforms_folder = f'waveforms/{test}'
############################################################################

########################## USER INPUT ######################################


LED = sys.argv[1] # 46, 36, 24V
vin = sys.argv[2] # 120, 230, 277Vac
if vin == '230':
    freq = 50
else:
    freq = 60

dim_trigger_source = 3
# ############################################################################

"""For Calibration"""
# currentMouseX, currentMouseY = pyautogui.position()
# print(currentMouseX,currentMouseY)
# input()



########################## FUNCTIONS #######################################
def set_analog_dim(dim=0):
    delay=1

    pyautogui.moveTo(543, 78)
    pyautogui.click()
    sleep(delay)

    # click settings
    pyautogui.moveTo(543, 78)
    pyautogui.click()
    sleep(delay)
    pyautogui.moveTo(541, 175)
    pyautogui.click()
    sleep(delay)

    # click 0-10V tab
    pyautogui.moveTo(151, 268)
    pyautogui.click()
    sleep(0.5)

    # manual 0-10V control
    pyautogui.moveTo(185, 359)
    pyautogui.doubleClick()
    sleep(0.5)
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    sleep(0.5)

    # set voltage
    pyautogui.write(f"{dim}")
    pyautogui.moveTo(360, 359)
    pyautogui.doubleClick()
    sleep(0.5)

def discharge_output():
    # ac.turn_off()
    eload.channel[1].cc = 1
    eload.channel[1].turn_on()
    eload.channel[2].cc = 1
    eload.channel[2].turn_on()
    sleep(3)
    eload.channel[1].turn_off()
    eload.channel[2].turn_off()

def analog_dimming(start=0, end=10, voltage=vin, frequency=freq, LED=LED,
                    condition="First Power Up Discharged Cout"):
    # turn on AC
    ac.voltage = voltage
    ac.frequency = frequency
    ac.turn_on()

    # set dimming
    set_analog_dim(start)
    sleep(1)

    # sleep(1)
    scope.run()
    # soak(2)
    sleep(1)
    set_analog_dim(end)
    sleep(6)
    scope.stop()
    sleep(1)
    if condition=="NA":
        filename = f'{LED}V, {voltage}Vac {frequency}Hz, {start}-{end}V Analog.png'
    else:
        filename = f'{LED}V, {voltage}Vac {frequency}Hz, {start}-{end}V Analog, {condition}.png'
    
    scope.get_screenshot(filename, waveforms_folder)
    print(filename)
    # print()

def main():
    
    soak(5)
    scope.time_position(10)
    dim_trigger_level = end - 0.5
    scope.edge_trigger(trigger_source=dim_trigger_source, trigger_level=dim_trigger_level,trigger_slope="POS")    
    scope.time_scale(0.5)

    set_analog_dim(0)
    for i in start_list:
        analog_dimming(i, end, vin, freq, LED, "First Power Up (Discharged Cout)")
        analog_dimming(i, end, vin, freq, LED, "No AC reset (Not Discharged Cout)")
        ac.turn_off()
        discharge_output()
        set_analog_dim(i+1)
        # print()

    scope.time_position(10)
    scope.edge_trigger(trigger_source=dim_trigger_source, trigger_level=dim_trigger_level,trigger_slope="NEG")

    for i in start_list:
        analog_dimming(end, i, vin, freq, LED, "NA")

############################################################################

############################################################################
# initialize equipment
ac = ACSource(ac_source_address)
# pms = PowerMeter(source_power_meter_address)
# pml = PowerMeter(load_power_meter_address)
eload = ElectronicLoad(eload_address)
scope = Oscilloscope(scope_address)

def soak(soak_time=30):
    for seconds in range(soak_time, 0, -1):
        sleep(1)
        print(f"{seconds:5d}s", end="\r")
    print("       ", end="\r")
############################################################################

########################## MAIN ##############################################
headers(test)

start_list = [0,1,2,3,4,5,6,7,8,9]
end = 10
main()

# start_list = [0,1,2,3,4,5,6]
# end = 7
# main()

# start_list = [0,1,2,3,4]
# end = 5
# main()

# start_list = [0,1,2,3]
# end = 4
# main()

# start_list = [0,1,2]
# end = 3
# main()

# start_list = [0,1]
# end = 2
# main()

# start_list = [0]
# end = 1
# main()

footers(waveform_counter)
############################################################################