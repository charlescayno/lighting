print("Discharging..")

"""LIBRARIES"""
from time import time, sleep
import sys
import os
from powi.equipment import ACSource, PowerMeter, ElectronicLoad, Oscilloscope
from powi.equipment import headers, create_folder, footers, waveform_counter, soak, convert_argv_to_int_list
waveform_counter = 0

"""COMMS"""
ac_source_address = 5
source_power_meter_address = 1 
load_power_meter_address = 2
eload_address = 8
scope_address = "10.125.10.148"

"""EQUIPMENT INITIALIZE"""
ac = ACSource(ac_source_address)
# pms = PowerMeter(source_power_meter_address)
# pml = PowerMeter(load_power_meter_address)
eload = ElectronicLoad(eload_address)
# scope = Oscilloscope(scope_address)
# control = MultiprotocolControl()


"""GENERIC FUNCTIONS"""
def discharge_output():
    ac.turn_off()
    eload.channel[1].cr = 100
    eload.channel[2].cr = 100
    eload.channel[1].turn_on()
    eload.channel[2].turn_on()
    sleep(2)
    eload.channel[1].turn_off()
    eload.channel[2].turn_off()


if __name__ == "__main__":
    discharge_output()