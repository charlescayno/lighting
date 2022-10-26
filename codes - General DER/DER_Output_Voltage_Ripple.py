# November 19, 2021

"""IMPORT DEPENDENCIES"""
from time import time, sleep
import sys
import os
import math
from powi.equipment import ACSource, PowerMeter, ElectronicLoad, Oscilloscope, LEDControl
from powi.equipment import headers, create_folder, footers, waveform_counter, soak, convert_argv_to_int_list, tts, prompt, sfx
from filemanager import path_maker, remove_file, move_file
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
vin_list = [100,115,230,265,277,300]
iout = 1.63
vout = 46
rload = vout/iout
rload_list = [rload, 2*rload]

eload_channel = 3

test = "Output Voltage Ripple - Retest - 1630 mA"
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
    print(">> Set CH3 to Output Voltage")
    print(">> Set ELOAD to CR mode")
    # input("Press ENTER to continue")
    
def scope_settings():
    global ch1, ch2, ch3, ch4
    global ton_time

    scope.stop()

    scope.position_scale(time_position = 50, time_scale = 0.4)
    
    scope.edge_trigger(3, 0, 'POS')


    scope.add_zoom(50, 5)

    ch1 = scope.channel_settings(state='ON', channel=3, scale=0.4, position=0, label="Output Voltage Ripple", color='GREEN', rel_x_position=50, bandwidth=20, coupling='AC', offset=0)
    scope.cursor(channel=3, cursor_set=1, X1=0, X2=0, Y1=-1, Y2=1, type='HOR')


def find_trigger(channel, trigger_delta):

    scope.edge_trigger(3, 0, 'POS')

    # finding trigger level
    scope.run_single()
    soak(5)

    # get initial peak-to-peak measurement value
    labels, values = scope.get_measure(channel)
    max_value = float(values[0])
    max_value = float(f"{max_value:.4f}")

    # set max_value as initial trigger level
    trigger_level = max_value
    scope.edge_trigger(channel, trigger_level, 'POS')

    # check if it triggered within 3 seconds
    scope.run_single()
    soak(3)
    trigger_status = scope.trigger_status()

    # increase trigger level until it reaches the maximum trigger level
    while (trigger_status == 1):
        trigger_level += trigger_delta
        scope.edge_trigger(channel, trigger_level, 'POS')

        # check trigger status
        scope.run_single()
        soak(3)
        trigger_status = scope.trigger_status()

    # decrease trigger level below to get the maximum trigger possible
    trigger_level -= 2*trigger_delta
    scope.edge_trigger(channel, trigger_level, 'POS')
    # print(f'Maximum trigger level found at: {trigger_level}')



def operation():

    global waveform_counter
    global ton_time
    
    scope_settings()
    
    for r in rload_list:

        # header for output text
        output_list = ['LED', 'VIN', 'Vout (V)', 'Iout (mA)', 'High Freq pk-pk (V)', '%pk-pk High Freq (%)', 'Low Freq pk-pk (V)', '%pk-pk Low Freq (%)']
        file = f"{test} at {r:.4f} CR Load.txt"
        with open(file, 'a+') as f:
            f.write(','.join(output_list))
            f.write('\n')

        for vin in vin_list:

            # turning on the unit
            ac.voltage = vin
            ac.turn_on()
            eload.channel[eload_channel].cr = r
            eload.channel[eload_channel].turn_on()
            sleep(3)

            # finding the maximm trigger
            find_trigger(channel=3, trigger_delta=0.002)
            scope.run_single()
            input("\nAdjust cursor. Press ENTER to capture waveform.")
            scope.stop()

            # saving waveform
            filename = f'Output Voltage Ripple - {r:.4f} CR Load, {vin}Vac.png'
            path = path_maker(f'{waveforms_folder}/{r:.4f} CR Load')
            scope.get_screenshot(filename, path)
            print(filename)
            waveform_counter += 1

            # output flicker and peak to peak to a text file
            labels, values = scope.get_measure(channel=3)
            high_freq_peak_to_peak = float(f"{values[2]:.4f}")
            low_freq_peak_to_peak = float(scope.get_cursor(cursor=1)['delta y'])
            vo1 = float(f"{pml.voltage:.3f}")
            io1 = float(f"{pml.current*1000:.2f}")
            percent_peak_to_peak_high_frequency = 100*(high_freq_peak_to_peak/vo1)
            percent_peak_to_peak_low_frequency = 100*(low_freq_peak_to_peak/vo1)

            # printing values for the user to see
            print(f'vo1:                                 {vo1} V')
            print(f'io1:                                 {io1} mA')
            print(f'high_freq_peak_to_peak:              {high_freq_peak_to_peak} V')
            print(f'percent_peak_to_peak_high_frequency: {percent_peak_to_peak_high_frequency} %')
            print(f'low_freq_peak_to_peak:               {low_freq_peak_to_peak} V')
            print(f'percent_peak_to_peak_low_frequency:  {percent_peak_to_peak_low_frequency} %')

            # compiling to all data to a list
            load = f'{r:.4f}'
            input_voltage = str(vin)
            vo1 = str(vo1)
            io1 = str(io1)
            output_list = [load, input_voltage, vo1, io1]
            high_freq_peak_to_peak = f"{high_freq_peak_to_peak:.4f}"
            output_list.append(high_freq_peak_to_peak)
            percent_peak_to_peak_high_frequency = f"{percent_peak_to_peak_high_frequency:.4f}"
            output_list.append(percent_peak_to_peak_high_frequency)
            low_freq_peak_to_peak = f"{low_freq_peak_to_peak:.4f}"
            output_list.append(low_freq_peak_to_peak)
            percent_peak_to_peak_low_frequency = f"{percent_peak_to_peak_low_frequency:.4f}"
            output_list.append(percent_peak_to_peak_low_frequency)

            # writing all data from the list to a text file
            file = f"{test} at {r:.4f} CR Load.txt"
            with open(file, 'a+') as f:
                f.write(','.join(output_list))
                f.write('\n')
            
            print()

        # moving text file to the file path
        source = f'{os.getcwd()}/{file}'
        destination = f'{path}/{file}'
        move_file(source, destination)


def main():
    global waveform_counter
    operation()
        
if __name__ == "__main__":
    message()
    headers(test)
    discharge_output()
    main()
    discharge_output()
    footers(waveform_counter)