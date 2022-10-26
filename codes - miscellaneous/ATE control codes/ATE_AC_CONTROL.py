# COMMS
ac_source_address = 5
source_power_meter_address = 1
load_power_meter_address = 2
eload_address = 8
scope_address = "10.125.10.148"
scope2_address = "10.125.10.227"

"""IMPORT DEPENDENCIES"""
import sys
import pyautogui
from time import sleep, time
from powi.equipment import ACSource, PowerMeter, ElectronicLoad, Oscilloscope
from powi.equipment import headers, create_folder, footers, waveform_counter, soak, convert_argv_to_int_list
from time import sleep, time
import os
waveform_counter = 0

"""INITIALIZE EQUIPMENT"""
ac = ACSource(ac_source_address)
pms = PowerMeter(source_power_meter_address)
pml = PowerMeter(load_power_meter_address)
eload = ElectronicLoad(eload_address)
scope = Oscilloscope(scope_address)
# scope2 = Oscilloscope(scope2_address)



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



voltage = int(sys.argv[1])

if voltage == 230: frequency = 50
elif voltage == 265: frequency = 50
else: frequency = 60

# scope.run_single()
# sleep(1)

ac.voltage = voltage
ac.frequency = frequency
ac.turn_on()

input("Press ENTER to turn off ac.")


discharge_output()
