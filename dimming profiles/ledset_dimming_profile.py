# COMMS
ac_source_address = 5
source_power_meter_address = 1
load_power_meter_address = 2
eload_address = 8
scope_address = "10.125.10.148"

"""IMPORT DEPENDENCIES"""
import sys
import pyautogui
from time import sleep, time
from powi.equipment import ACSource, PowerMeter, ElectronicLoad, Oscilloscope
from powi.equipment import headers, create_folder, footers, waveform_counter, soak, convert_argv_to_int_list
import os
from powi.equipment import LEDControl
waveform_counter = 0

"""INITIALIZE EQUIPMENT"""
ac = ACSource(ac_source_address)
pms = PowerMeter(source_power_meter_address)
pml = PowerMeter(load_power_meter_address)
eload = ElectronicLoad(eload_address)
# scope = Oscilloscope(scope_address)
led = LEDControl()

def discharge_output():
    ac.turn_off()

    eload.channel[1].cr = 100
    eload.channel[1].turn_on()
    eload.channel[2].cr = 100
    eload.channel[2].turn_on()
    eload.channel[3].cr = 100
    eload.channel[3].turn_on()
    eload.channel[4].cr = 100
    eload.channel[4].turn_on()
    eload.channel[5].cr = 100
    eload.channel[5].turn_on()

    # input("check eload")
    sleep(1)
    
    eload.channel[1].turn_off()
    eload.channel[2].turn_off()
    eload.channel[3].turn_off()
    eload.channel[4].turn_off()
    eload.channel[5].turn_off()

    eload.channel[1].cc = 1
    eload.channel[1].turn_on()
    eload.channel[2].cc = 1
    eload.channel[2].turn_on()
    eload.channel[3].cc = 1
    eload.channel[3].turn_on()
    eload.channel[4].cc = 1
    eload.channel[4].turn_on()
    eload.channel[5].cc = 1
    eload.channel[5].turn_on()
    
    sleep(1)

    eload.channel[1].turn_off()
    eload.channel[2].turn_off()
    eload.channel[3].turn_off()
    eload.channel[1].turn_off()
    eload.channel[2].turn_off()
    eload.channel[3].turn_off()
    eload.channel[4].turn_off()
    eload.channel[5].turn_off()

    sleep(1)

### USER INPUT #######################################
led_list = [46,36,24]
led_list = convert_argv_to_int_list(sys.argv[1])

vin_list = [120,230,277]
vin_list = convert_argv_to_int_list(sys.argv[2])

rset_list = [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20] 

IOUT_REG = 1050

test = f"LEDSET Dimming"
waveforms_folder = f'waveforms/{test}'
######################################################

def ledset_line_regulation(soak_time=10):

    for rset in rset_list:
        
        input(f"RSET: {rset}kOhms")

        for LED in led_list:

            led.voltage(LED)          

            for vin in vin_list:

                print(f"{LED}V, {vrms}Vac, RSET Dimming")

                sleep(2)

                ac.voltage = vin
                freq = ac.set_freq(vin)
                ac.frequency = freq
                ac.turn_on()

                soak(soak_time)

                # create output list
                vac = str(vin)
                freq = str(freq)
                vin = f"{pms.voltage:.2f}"
                iin = f"{pms.current*1000:.2f}"
                pin = f"{pms.power:.3f}"
                pf = f"{pms.pf:.4f}"
                thd = f"{pms.thd:.2f}"
                vo1 = f"{pml.voltage:.3f}"
                io1 = f"{pml.current*1000:.2f}"
                po1 = f"{pml.power:.3f}"
                ireg1 = f"{100*(float(io1)-IOUT_REG)/12:.4f}"
                eff = f"{100*(float(po1))/float(pin):.4f}"

                output_list = [str(rset), vac, freq, vin, iin, pin, pf, thd, vo1, io1, po1, ireg1, eff]
                print(','.join(output_list))

                with open(f'{LED}V, {vin}Vac, RSET Dimming.txt', 'a+') as f:
                    f.write(','.join(output_list))
                    f.write('\n')

                discharge_output()

def main():
    ledset_line_regulation(soak_time=10)
    
if __name__ == "__main__":
    headers(test)
    main()
    footers(waveform_counter)
    