batch = input("Batch: ")
unit = input("Unit: ")
print("Charles Cayno | 25-Jan-2022")

test = f"Brown-In (Batch {batch}) - Unit {unit}"

# comms
ac_source_address = 5
source_power_meter_address = 1 
load_power_meter_address = 2
eload_address = 8
scope_address = "10.125.10.184"


## USER INPUT ###

start_voltage = 0
end_voltage = 120
slew_rate = 1
frequency = 60
time_fixvoltage = 80


#########################################################################
from powi.equipment import ACSource, PowerMeter, ElectronicLoad, Oscilloscope
from powi.equipment import headers, create_folder, footers, waveform_counter
from time import sleep, time
import os
import math
waveform_counter = 0
waveforms_folder = f'waveforms/{test}'

# # initialize equipment
ac = ACSource(ac_source_address)
pms = PowerMeter(source_power_meter_address)
pml = PowerMeter(load_power_meter_address)
eload = ElectronicLoad(eload_address)
scope = Oscilloscope(scope_address)

def reset():
    ac.turn_off()
    eload.channel[1].cc = 1
    eload.channel[1].turn_on()
    eload.channel[2].cc = 1
    eload.channel[2].turn_on()
    sleep(3)
    eload.channel[1].turn_off()
    eload.channel[2].turn_off()


def browning(start, end, slew, frequency):

    if start > end:
        print(f"brownout: {start} -> {end} Vac")
        for voltage in range(start, end+1, -slew):
            ac.voltage = voltage
            ac.frequency = frequency
            ac.turn_on()
            sleep(1)
    if start < end:
        print(f"brownin: {start} -> {end} Vac")
        for voltage in range(start, end+1, slew):
            ac.voltage = voltage
            ac.frequency = frequency
            ac.turn_on()
            sleep(1)

def fixvoltage(voltage, soak):
    ac.voltage = voltage
    ac.frequency = ac.set_freq(voltage)
    ac.turn_on()
    sleep(soak)

def roundup(x):
    return int(math.ceil(x / 100.0)) * 100

def soak(soak_time=30):
    for seconds in range(soak_time, 0, -1):
        sleep(1)
        print(f"{seconds:5d}s", end="\r")
    print("       ", end="\r")

## main code ##
reset()

headers(test)

# setup time scale for the oscilloscope
test_time = 2*(end_voltage-start_voltage)*slew_rate+time_fixvoltage
scope_time = roundup(test_time)
delay = scope_time-test_time
time_scale = scope_time/10
scope.time_scale(time_scale)
scope.run()
soak(int(delay/2))

# start of test
browning(start_voltage, end_voltage, slew_rate, frequency)
fixvoltage(end_voltage,time_fixvoltage)
browning(end_voltage, start_voltage, slew_rate, frequency)

reset()


soak(int(delay/2))
scope.stop()

filename = f"{start_voltage}-{end_voltage}-{start_voltage} Vac.png"
scope.get_screenshot(filename, waveforms_folder)
waveform_counter += 1
print(filename)


footers(waveform_counter)
