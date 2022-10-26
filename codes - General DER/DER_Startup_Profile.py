"""IMPORT DEPENDENCIES"""
from time import time, sleep
import sys
import os
import math
from powi.equipment import ACSource, PowerMeter, ElectronicLoad, Oscilloscope, LEDControl
from powi.equipment import headers, create_folder, footers, waveform_counter, soak, convert_argv_to_int_list, tts, prompt
from filemanager import path_maker, remove_file
import winsound as ws
from playsound import playsound
waveform_counter = 0

##################################################################################

"""COMMS"""
ac_source_address = 5
source_power_meter_address = 1 
load_power_meter_address = 2
eload_address = 8
scope_address = "10.125.10.115"

"""USER INPUT"""
led_list = [46,36,24]
# led_list = [36,24]
vin_list = [100,115,230,265,277,300]
iout = 1.66

test = "Startup Profile"
waveforms_folder = f'C:/Users/ccayno/Desktop/DER-945/Test Data/{test}'

"""DO NOT EDIT BELOW THIS LINE"""
##################################################################################

"""EQUIPMENT INITIALIZE"""
ac = ACSource(ac_source_address)
pms = PowerMeter(source_power_meter_address)
pml = PowerMeter(load_power_meter_address)
eload = ElectronicLoad(eload_address)
scope = Oscilloscope(scope_address)


"""GENERIC FUNCTIONS"""

def discharge_output():
    ac.turn_off()
    for i in range(1,9):
        eload.channel[i].cc = 1
        eload.channel[i].turn_on()
        eload.channel[i].short_on()
    sleep(2)
    for i in range(1,9):
        eload.channel[i].turn_off()
        eload.channel[i].short_off()
    sleep(1)


def message():
    print(">> Set CH1 to Output Current")
    print(">> Set CH4 to Input Voltage")
    input("Press ENTER to continue")


def scope_settings():
    global ch1, ch2, ch3, ch4
    global ton_time


    scope.stop()
    scope.remove_zoom()
    scope.cursor(channel=2, cursor_set=1, X1=0, X2=0, Y1=0, Y2=0, type='HOR')
    scope.cursor(channel=2, cursor_set=2, X1=0, X2=0, Y1=0, Y2=0, type='HOR')

    ch1 = scope.channel_settings(state='OFF')
    ch2 = scope.channel_settings(state='OFF')
    ch3 = scope.channel_settings(state='OFF')
    ch4 = scope.channel_settings(state='OFF')

    scope.position_scale(time_position = 10, time_scale = 0.100)
    scope.edge_trigger(4, 10, 'POS')

    ch1 = scope.channel_settings(state='ON', channel=1, scale=0.5, position=1, label="Output Current", color='ORANGE', rel_x_position=30, bandwidth=500, coupling='DCLimit', offset=0)
    ch4 = scope.channel_settings(state='ON', channel=4, scale=200, position=-2, label="Input Voltage", color='PINK', rel_x_position=70, bandwidth=500, coupling='AC', offset=0)
    scope.cursor(channel=1, cursor_set=1, X1=0, X2=0.4, Y1=0, Y2=iout, type='HOR')
    
    ton_time = 0.4 # initial setting
    scope.cursor(channel=1, cursor_set=2, X1=0, X2=ton_time, Y1=0, Y2=1, type='PAIR')


def startup_at_90deg(vin, soak):
    ac.voltage = 0
    ac.frequency = ac.set_freq(0)
    ac.turn_on()
    ac.write("TRIG:TRAN:SOUR BUS")
    ac.write(f"LIST:DWEL        1,       {soak}")
    ac.write("VOLT:MODE LIST")
    ac.write(f"LIST:VOLT        0,      {vin}")
    ac.write("VOLT:SLEW:MODE LIST")
    ac.write("LIST:VOLT:SLEW 9.9e+037, 9.9e+037")
    ac.write("FREQ:MODE LIST")
    ac.write(f"LIST:FREQ       {ac.set_freq(0)},       {ac.set_freq(vin)}")
    ac.write("FREQ:SLEW:MODE LIST")
    ac.write("LIST:FREQ:SLEW 9.9e+037, 9.9e+037")
    ac.write("VOLT:OFFS:MODE FIX")
    ac.write("VOLT:OFFS:SLEW:MODE FIX")
    ac.write("PHAS:MODE LIST")
    ac.write("LIST:PHAS       90,       90")
    ac.write("CURR:PEAK:MODE LIST")
    ac.write("LIST:CURR     40.4,     40.4")
    ac.write("FUNC:MODE FIX")
    ac.write("LIST:TTLT ON,OFF")
    ac.write("LIST:STEP AUTO")
    ac.write("OUTP:TTLT:STAT ON")
    ac.write("OUTP:TTLT:SOUR LIST")
    ac.write("TRIG:SYNC:SOUR PHASE")
    ac.write("TRIG:SYNC:PHAS 0.0")
    ac.write("TRIG:TRAN:DEL 0")
    ac.write("Sens:Swe:Offs:Poin 0")
    ac.write("TRIG:ACQ:SOUR TTLT")
    ac.write("INIT:IMM:SEQ3")
    ac.write("LIST:COUN 1")
    ac.write("INIT:IMM:SEQ1")
    ac.write("TRIG:TRAN:SOUR BUS")
    ac.write("TRIG:IMM")


def operation():

    global waveform_counter
    global ton_time

    scope_settings()
    
    for LED in led_list:

        led.voltage(LED)
        
        for vin in vin_list:

                scope.run_single()

                sleep(3)

                
                startup_at_90deg(vin, soak=3)
                
                sleep(3)

                input("Set to Iout_DC")

                iout_dc = float(scope.get_cursor(cursor=1)['y2 position'])
                iout_dc_90percent = 0.9*iout_dc
                scope.cursor(channel=1, cursor_set=1, X1=0, X2=0.4, Y1=0, Y2=iout_dc, type='HOR')
                scope.cursor(channel=1, cursor_set=2, X1=0, X2=ton_time, Y1=0, Y2=iout_dc_90percent, type='PAIR')

                input("Adjust Turn-on Time")
                ton_time = float(scope.get_cursor(cursor=2)['x2 position'])


                filename = f'Startup Profile - {LED}V, {vin}Vac.png'
                path = path_maker(f'{waveforms_folder}/{LED}V')
                scope.get_screenshot(filename, path)
                print(filename)
                waveform_counter += 1
                
                discharge_output()


def main():
    global waveform_counter
    operation()
        
if __name__ == "__main__":
    message()
    headers(test)
    discharge_output()
    led = LEDControl()
    main()
    discharge_output()
    footers(waveform_counter)