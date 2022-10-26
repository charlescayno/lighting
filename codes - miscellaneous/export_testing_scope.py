"""LIBRARIES"""
from time import time, sleep
import sys
import os
import math
from powi.equipment import ACSource, PowerMeter, ElectronicLoad, Oscilloscope, LEDControl
from powi.equipment import headers, create_folder, footers, waveform_counter, soak, convert_argv_to_int_list, tts
import winsound as ws
waveform_counter = 0

"""COMMS"""
ac_source_address = 5
source_power_meter_address = 1 
load_power_meter_address = 2    
eload_address = 8
scope_address = "10.125.11.0"

"""EQUIPMENT INITIALIZE"""
ac = ACSource(ac_source_address)
pms = PowerMeter(source_power_meter_address)
pml = PowerMeter(load_power_meter_address)
eload = ElectronicLoad(eload_address)
scope = Oscilloscope(scope_address)
# led = LEDControl()

# scope.remove_zoom()
# scope.add_zoom(rel_pos=50, rel_scale=10)
# # scope.add_zoom_1(rel_pos=20, zoom_scale=0.5)
# test = "Test"
# headers(test)
# tts("Hellow")
# footers(waveform_counter)


# scope.write("STOP;*OPC?")
scope.write("EXPort:WAVeform:FASTexport ON")
scope.write("CHANnel1:WAVeform1:STATe 1")
# scope.write("LAYout:ZOOM:ADD 'Diagram1',HORIZONTAL,OFF,-0.00012,-5e-06,0.308,-0.092,'ExportAreaZoom'")
scope.write("EXPort:WAVeform:SOURce C1W1")
scope.write("EXPort:WAVeform:SCOPe ZOOM")
scope.write("EXPort:WAVeform:ZOOM 'Diagram1', 'ExportAreaZoom'")
scope.write("EXPort:WAVeform:NAME 'C:\Data\DataExportWfm_analog.csv'")
scope.write("EXPort:WAVeform:RAW OFF")
scope.write("EXPort:WAVeform:INCXvalues ON")
scope.write("EXPort:WAVeform:DLOGging OFF")
scope.write("MMEM:DEL 'C:\Data\DataExportWfm_analog.*'")
scope.write("EXPort:WAVeform:SAVE")
# scope.write("MMEM:DATA? 'C:\Data\DataExportWfm_analog.csv'")
# scope.write("MMEM:DATA? 'C:\Data\DataExportWfm_analog.wfm.csv'")








