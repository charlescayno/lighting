print("Charles Cayno | 10-Jun-2021")
print("Autogui Configured on Personal Laptop")

"""COMMS ADDRESS"""
ac_source_address = 5
source_power_meter_address = 1 
load_power_meter_address = 2
eload_address = 8
scope_address = "10.125.10.166"

"""IMPORT DEPENDENCIES"""
import sys
import pyautogui
from time import sleep, time
from powi.equipment import ACSource, PowerMeter, ElectronicLoad, Oscilloscope
from powi.equipment import headers, create_folder, footers, waveform_counter, soak, convert_argv_to_int_list
from time import sleep, time
import os

"""USER INPUT"""


LED = sys.argv[1] # 46, 36, 24V
vin_list = convert_argv_to_int_list(sys.argv[2]) # [120, 230, 277] Vac
pwm_frequency_list = convert_argv_to_int_list(sys.argv[3]) # [2000, 400] Hz
end_list = convert_argv_to_int_list(sys.argv[4]) # [100, 70, 50, 30, 20, 10]
test_list = convert_argv_to_int_list(sys.argv[5])   # 1 - First Power Up and No AC Reset Test
                                                    # 2 - Transient Down
# raise Exception('python LIGHTING_PWM_DIMMING.py 46 [120,230,277] [2000,400] [100,70,50,40,30,20,10] [1,2]')
# print("Reload Saveset pwm_dimming.")
print("CH1: VR | CH2: VOUT | CH3: DIM+ | CH4: IOUT")

dim_trigger_source = 3
iout_oscilloscope_channel = 4


"""PYAUTOGUI CALIBRATION"""
# x, y = pyautogui.position()
# print(x,y)
# input()

"""DEFAULT FUNCTIONS"""

def discharge_output():
    ac.turn_off()
    eload.channel[1].cc = 1
    eload.channel[1].turn_on()
    eload.channel[2].cc = 1
    eload.channel[2].turn_on()
    sleep(2)
    eload.channel[1].turn_off()
    eload.channel[2].turn_off()
    pass

def soak(soak_time=30):
    for seconds in range(soak_time, 0, -1):
        sleep(1)
        print(f"{seconds:5d}s", end="\r")
    print("       ", end="\r")



"""CUSTOM FUNCTIONS FOR THIS TEST"""

def initial_dim_autogui():

    # click ATE file
    pyautogui.moveTo(483, 18)
    pyautogui.click()
    sleep(0.25)

    # click 'esc' for occassional gui errors
    pyautogui.press('esc')
    sleep(0.25)

    # click test selection
    pyautogui.moveTo(479, 93)
    pyautogui.click()
    sleep(1)

    # click settings
    pyautogui.moveTo(474, 206)
    pyautogui.click()
    sleep(1)

    # click PWM tab
    pyautogui.moveTo(151, 199)
    pyautogui.click()

def set_pwm_dim(dim=0, pwm_frequency=2000):

    # click Initialize Com Port
    pyautogui.moveTo(133, 717)
    pyautogui.click()

    # set PWM freq
    pyautogui.moveTo(143, 315)
    pyautogui.doubleClick()
    pyautogui.press('backspace', presses=3)
    pyautogui.write(f"{pwm_frequency}")

    # set PWM duty
    pyautogui.moveTo(379, 315)
    pyautogui.doubleClick()
    pyautogui.press('backspace', presses=3)
    pyautogui.write(f"{dim}")

    # set pwm duty and frequency
    pyautogui.moveTo(614, 315)
    pyautogui.doubleClick()

def pwm_dimming(start=0, end=10, voltage=120, frequency=60, LED=LED,
                    condition="First Power Up Discharged Cout",
                    pwm_frequency=2000):
    global waveform_counter

    set_pwm_dim(start, pwm_frequency)
    
    ac.voltage = voltage
    ac.frequency = frequency
    ac.turn_on()

    sleep(3)

    scope.run_single()

    sleep(1)

    set_pwm_dim(end, pwm_frequency)

    sleep(6)

    scope.stop()

    sleep(1)

    if condition=="NA":
        filename = f'{LED}V, {voltage}Vac {frequency}Hz, {start}-{end}% PWM Dimming, f={pwm_frequency}Hz.png'
    else:
        filename = f'{LED}V, {voltage}Vac {frequency}Hz, {start}-{end}% PWM Dimming, f={pwm_frequency}Hz, {condition}.png'
     
    scope.get_screenshot(filename, waveforms_folder)
    print(filename)
    waveform_counter += 1

def pwm_width(f=2000, pwm_start=0, pwm_end=100):
    
    T = 1/int(f)

    Ton1 = T*(pwm_start/100)
    Ton2 = T*(pwm_end/100)

    if Ton1 < Ton2:
        width = Ton1
        width_range = 'OUTS'
    else:
        width = Ton2
        width_range = 'WITH'

    return width, width_range

def main():
    
    discharge_output()
    soak(3)

    # initial_dim_autogui() # to minimize testing time

    """Adjust Initial Scope Settings"""
    scope.time_position(10)
    scope.edge_trigger(trigger_source=dim_trigger_source, trigger_level=5, trigger_slope="POS")    
    scope.time_scale(0.5)


    """First Power Up and No AC Reset Test"""
    for i in test_list:

        if i == 1:        

            for i in start_list:

                """Adjust Trigger settings"""
                if end == 100:
                    scope.timeout_trigger(trigger_source=dim_trigger_source, timeout_range='HIGH', timeout_time=2.8E-3)
                else:
                    width, width_range = pwm_width(pwm_frequency, i, end)
                    delta = width*0.03
                    scope.width_trigger(trigger_source=dim_trigger_source, width_polarity='POS', width_range=width_range, width=width, delta=delta)

                """Iout Channel Scale Adjuster"""
                if end <= 30:
                    iout_channel_scale = 0.04
                    scope.channel_scale(iout_oscilloscope_channel, iout_channel_scale)
                else:
                    iout_channel_scale = 0.2
                    scope.channel_scale(iout_oscilloscope_channel, iout_channel_scale)

                pwm_dimming(i, end, vin, freq, LED, "First Power Up (Discharged Cout)", pwm_frequency)
                # pwm_dimming(i, end, vin, freq, LED, "No AC reset (Not Discharged Cout)", pwm_frequency)

                discharge_output()
        
            print()

    """Transient Down"""

    for i in test_list:

        if i == 2:

            for i in start_list:

                """adjust trigger"""
                if i == 0:
                    scope.timeout_trigger(trigger_source=dim_trigger_source, timeout_range='LOW', timeout_time=2.8E-3)
                else:
                    width, width_range = pwm_width(pwm_frequency, end, i)
                    delta = width*0.03
                    scope.width_trigger(trigger_source=dim_trigger_source, width_polarity='POS', width_range=width_range, width=width, delta=delta)

                pwm_dimming(end, i, vin, freq, LED, "NA", pwm_frequency)

                discharge_output()
            
            print()
    
    waveform_counter = 0

"""INITIALIZE EQUIPMENT"""
ac = ACSource(ac_source_address)
eload = ElectronicLoad(eload_address)
scope = Oscilloscope(scope_address)
# pms = PowerMeter(source_power_meter_address)
# pml = PowerMeter(load_power_meter_address)
"""MAIN"""


# initial_dim_autogui() # to minimize testing time

for vin in vin_list:

    for pwm_frequency in pwm_frequency_list:

        if vin == '230': freq = 50
        else: freq = 60

        test = f"PWM Dimming (f={pwm_frequency}Hz) {LED}V {vin}Vac"
        waveforms_folder = f'waveforms/{test}'

        headers(test)

        # for end in end_list:
        #     start_list = range(0, end, 10)
        #     main()

        """Special Case: Recapturing all Startup from 0-end"""
        for end in end_list:
            start_list = range(0,end,end)
            main()   


        footers(waveform_counter)
        waveform_counter = 0