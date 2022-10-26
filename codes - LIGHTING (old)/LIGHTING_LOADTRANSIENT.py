print("CMC | 22-Jun-2021")

"""IMPORT DEPENDENCIES"""
import sys
import pyautogui
from time import sleep, time
from powi.equipment import ACSource, PowerMeter, ElectronicLoad, Oscilloscope
from powi.equipment import headers, create_folder, footers, waveform_counter, soak, convert_argv_to_int_list, convert_argv_to_str_list
from powi.equipment import *
from time import sleep, time
import os
import winsound as ws

# USER INPUT STARTS HERE
########################################################################################################
"""COMMS ADDRESS"""
ac_source_address = 9
source_power_meter_address = 2 
load_power_meter_address = 10
eload_address = 12
scope_address = "10.125.10.229"

"""TEST TITLE"""
test = f"Output Load Transient"
waveforms_folder = f'waveforms/{test}'


"""USER INPUT"""
LED = sys.argv[1] # 46, 36, 24V
vin = convert_argv_to_int_list(sys.argv[2]) # [120, 230, 277] Vac
transient_frequency = float(sys.argv[3]) # 500 Hz
timeperdiv = float(sys.argv[4]) # 1s, 2s

eload_channel = 1

# initial trigger settings
trigger_source = 2
trigger_level = 50
trigger_slope = 'POS'

vout = int(LED)
iout = 1.667
rload = vout/iout

iout_list = [iout, iout*0.5]
rload_list = [rload, rload*2]

ton = 1/(2*transient_frequency)
toff = ton




# input()
########################################################################################################
########################################################################################################
# USER INPUT ENDS HERE

"""INITIALIZE EQUIPMENT"""
ac = ACSource(ac_source_address)
pms = PowerMeter(source_power_meter_address)
pml = PowerMeter(load_power_meter_address)
eload = ElectronicLoad(eload_address)
scope = Oscilloscope(scope_address)

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

#### special functions #####

def loadtransient_cc(x=1, case="0-100", eload_channel=1, ton=0.05, toff=0.05):
    global waveform_counter

    # parse values from case
    temp = case.split("-")
    low = x*(float(temp[0])/100)
    if low == 0: low = 0
    high = x*(float(temp[1])/100)
    if high == 0: high = 0


    # if low < high: trigger_slope = 'POS'
    # else: trigger_slope = 'NEG'
    # # adjust trigger level
    # trigger_level = (low+high)/2
    # scope.edge_trigger(trigger_source=trigger_source, trigger_level=trigger_level, trigger_slope=trigger_slope)



    eload.channel[eload_channel].dynamic(low, high, ton, toff)
    eload.channel[eload_channel].turn_on()


    # get screenshot
    scope.run_single()
    sleep(timeperdiv)
    
    ac.voltage = voltage
    ac.frequency = frequency
    ac.turn_on()

    sleep(9*timeperdiv)

    filename = f'{LED}V {voltage}Vac {frequency}Hz {case}LoadCC {transient_frequency}Hz.png'
    scope.get_screenshot(filename, waveforms_folder)
    print(filename)
    waveform_counter += 1

    discharge_output()

def main():

    scope.init_trigger(trigger_source, trigger_level, trigger_slope)

    global voltage
    global frequency

    for voltage in vin:
        if voltage == 230: frequency = 50
        elif voltage == 265: frequency = 50
        else: frequency = 60

        ac.voltage = voltage
        ac.frequency = frequency
        ac.turn_on()

        print(f"[{voltage}Vac {frequency}Hz]")

        
        loadtransient_cc(iout, "0-100", eload_channel, ton, toff)
        loadtransient_cc(iout, "0-50", eload_channel, ton, toff)
        loadtransient_cc(iout, "50-100", eload_channel, ton, toff)

        sleep(2)
        discharge_output()
        print()


def roll_main():

    scope.init_trigger(trigger_source, trigger_level, trigger_slope)
    scope.time_scale(timeperdiv)
    scope.time_position(10)

    global voltage
    global frequency

    for voltage in vin:
        if voltage == 230: frequency = 50
        elif voltage == 265: frequency = 50
        else: frequency = 60

        print(f"[{voltage}Vac {frequency}Hz]")

        loadtransient_cc(iout, "0-100", eload_channel, ton, toff)
        loadtransient_cc(iout, "0-50", eload_channel, ton, toff)
        loadtransient_cc(iout, "50-100", eload_channel, ton, toff)

        discharge_output()
        print()

## main code ##
headers(test)
roll_main()
footers(waveform_counter)
