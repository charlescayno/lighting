print("-"*80)
developer = "Charles Cayno"
lastmod = "05-Jun-2021"
codestatus = "Code Complete"
print(f"{developer} | {lastmod} | {codestatus}")
print("-"*80)

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
############################################################################

########################## USER INPUT ######################################

LED = sys.argv[1] # 46, 36, 24V
vin = sys.argv[2] # 120, 230, 277Vac
pwm_frequency = sys.argv[3] # 2000Hz, 400Hz
start_pwm = int(sys.argv[4])
end_pwm = int(sys.argv[5])
# dim_trigger_source = sys.argv[6]
# dim_trigger_level = sys.argv[7]

dim_trigger_source = 3

if start_pwm < end_pwm:
    if start_pwm == 0: dim_trigger_level = input("Enter dim voltage")
    if start_pwm == 10: dim_trigger_level = input("Enter dim voltage")
    if start_pwm == 20: dim_trigger_level = 2.6
    if start_pwm == 30: dim_trigger_level = 3.7
    if start_pwm == 40: dim_trigger_level = 4.8
    if start_pwm == 50: dim_trigger_level = 4.6
    if start_pwm == 60: dim_trigger_level = 7
    if start_pwm == 70: dim_trigger_level = input("Enter dim voltage")
    if start_pwm == 80: dim_trigger_level = input("Enter dim voltage")
    if start_pwm == 90: dim_trigger_level = input("Enter dim voltage")
    if start_pwm == 100: dim_trigger_level = input("Enter dim voltage")
    # else: input("Enter correct dim level.")
else:
    if start_pwm == 0: dim_trigger_level = input("Enter dim voltage")
    if start_pwm == 10: dim_trigger_level = input("Enter dim voltage")
    if start_pwm == 20: dim_trigger_level = input("Enter dim voltage")
    if start_pwm == 30: dim_trigger_level = input("Enter dim voltage")
    if start_pwm == 40: dim_trigger_level = input("Enter dim voltage")
    if start_pwm == 50: dim_trigger_level = 4.6
    if start_pwm == 60: dim_trigger_level = 5.6
    if start_pwm == 70: dim_trigger_level = 6.6
    if start_pwm == 80: dim_trigger_level = input("Enter dim voltage")
    if start_pwm == 90: dim_trigger_level = input("Enter dim voltage")
    if start_pwm == 100: dim_trigger_level = input("Enter dim voltage")
    # else: input("Enter correct dim level.")







if start_pwm > end_pwm:
    trigger_slope = 'NEG'
    test_type = 2
    start = end_pwm
    end = start_pwm
else:
    trigger_slope = 'POS'
    test_type = 1
    start = start_pwm
    end = end_pwm


dim_trigger_source
dim_trigger_level


iout_oscilloscope_channel = 4

if vin == '230':
    freq = 50
else:
    freq = 60

test = f"PWM Dimming (f={pwm_frequency}Hz) {LED}V {vin}Vac"
waveforms_folder = f'waveforms/{test}'

# ############################################################################




############################################################################

"""For pyautogui Calibration"""
# currentMouseX, currentMouseY = pyautogui.position()
# print(currentMouseX,currentMouseY)

# input()
########################## FUNCTIONS #######################################

def discharge_output():
    # ac.turn_off()
    eload.channel[1].cc = 1
    eload.channel[1].turn_on()
    eload.channel[2].cc = 1
    eload.channel[2].turn_on()
    sleep(2)
    eload.channel[1].turn_off()
    eload.channel[2].turn_off()

def set_pwm_dim(dim=0, pwm_frequency=2000):

    # click ATE file
    pyautogui.moveTo(543, 78)
    pyautogui.click()
    sleep(0.25)

    pyautogui.press('esc')
    sleep(0.25)

    # click test selection
    pyautogui.moveTo(543, 78)
    pyautogui.click()
    sleep(1)

    # click settings
    pyautogui.moveTo(541, 175)
    pyautogui.click()
    sleep(1)

    # click PWM tab
    pyautogui.moveTo(116, 264)
    pyautogui.click()
    # sleep(0.2)

    # click Initialize Com Port
    pyautogui.moveTo(101, 681)
    pyautogui.click()
    # sleep(0.2)

    # set PWM freq
    pyautogui.moveTo(109, 355)
    pyautogui.doubleClick()
    # sleep(0.2)
    pyautogui.press('backspace', presses=3)
    # sleep(0.2)
    pyautogui.write(f"{pwm_frequency}")

    # set PWM duty
    pyautogui.moveTo(301, 359)
    pyautogui.doubleClick()
    # sleep(0.2)
    pyautogui.press('backspace', presses=3)
    # sleep(0.2)
    pyautogui.write(f"{dim}")

    # set voltage
    pyautogui.moveTo(500, 359)
    pyautogui.doubleClick()
    # sleep(0.2)

def pwm_dimming(start=0, end=10, voltage=vin, frequency=freq, LED=LED,
                    condition="First Power Up Discharged Cout",
                    pwm_frequency=2000):
    
    set_pwm_dim(start, pwm_frequency)
    
    sleep(2)

    ac.voltage = voltage
    ac.frequency = frequency
    ac.turn_on()

    scope.run_single()
    sleep(1)

    set_pwm_dim(end, pwm_frequency)

    sleep(6)

    scope.stop()

    sleep(0.5)

    if condition=="NA":
        filename = f'{LED}V, {voltage}Vac {frequency}Hz, {start}-{end}% PWM Dimming, f={pwm_frequency}Hz.png'
    else:
        filename = f'{LED}V, {voltage}Vac {frequency}Hz, {start}-{end}% PWM Dimming, f={pwm_frequency}Hz, {condition}.png'
     
    scope.get_screenshot(filename, waveforms_folder)
    print(filename)



def main():
    
    ac.turn_off()
    discharge_output()

    # soak(5)
    scope.time_position(10)
    # dim_trigger_level = (end - 5)*0.01*10
    # trigger_slope = "POS"
    scope.edge_trigger(trigger_source=dim_trigger_source, trigger_level=dim_trigger_level, trigger_slope=trigger_slope)    
    scope.time_scale(0.5)

    # set_pwm_dim(0, pwm_frequency) 

    for i in start_list:

        # Iout Channel Scale Adjuster
        if end <= 30:
            iout_channel_scale = 0.04
            scope.channel_scale(iout_oscilloscope_channel, iout_channel_scale)
        else:
            iout_channel_scale = 0.2
            scope.channel_scale(iout_oscilloscope_channel, iout_channel_scale)

        if test_type == 1:
            pwm_dimming(i, end, vin, freq, LED, "First Power Up (Discharged Cout)", pwm_frequency)
            pwm_dimming(i, end, vin, freq, LED, "No AC reset (Not Discharged Cout)", pwm_frequency)

        ac.turn_off()
        discharge_output()
        
        # set_pwm_dim(i+10, pwm_frequency) # +10% duty TODO: think of way to dynamically set to next item in list

    scope.time_position(10)
    # trigger_slope = "POS"
    scope.edge_trigger(trigger_source=dim_trigger_source, trigger_level=dim_trigger_level, trigger_slope=trigger_slope)

    if test_type == 2:
        for i in start_list:
            pwm_dimming(end, i, vin, freq, LED, "NA", pwm_frequency)

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

start_list = [start]
end = end
main()

footers(waveform_counter)
############################################################################