print("Charles Cayno | 02-Jul-2021")

"""COMMS"""
ac_source_address = 5
source_power_meter_address = 1 
load_power_meter_address = 2    
eload_address = 8
scope_address = "10.125.11.0"

"""IMPORT DEPENDENCIES"""
from time import time, sleep
import sys
import os
import math
from powi.equipment import ACSource, PowerMeter, ElectronicLoad, Oscilloscope, LEDControl
from powi.equipment import headers, create_folder, footers, waveform_counter, soak, convert_argv_to_int_list
import winsound as ws
waveform_counter = 0


"""INITIALIZE EQUIPMENT"""
# ac = ACSource(ac_source_address)
# pms = PowerMeter(source_power_meter_address)
# pml = PowerMeter(load_power_meter_address)
# eload = ElectronicLoad(eload_address)
scope = Oscilloscope(scope_address)

scope.run()
sleep(3)
scope.stop()

a = scope.save_channel_data(4)
print(a)


