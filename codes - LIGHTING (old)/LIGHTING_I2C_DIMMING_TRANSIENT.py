print("TEST 1: Set the proper sleep delay for one test item. To minimize testing time. Test with only 1s delay")
print("TEST 2: Test for 1 set of data to ensure correct waveforms. ")







from datetime import datetime

start_timestamp = datetime.now()

total_waveform_counter = 0

for i in range(2): print()

print("Charles Cayno | 25-Jun-2021")

for i in range(2): print()

print("CODE USAGE:")
print("python LIGHTING_I2C_DIMMING_TRANSIENT.py 46 120 0 1")
print("<LED load> <Input Voltage> <CP Setting> <Test List>")

print()
print("PROBES USED:")
print("CH1: VR | CH2: VOUT | CH3: DIM+ | CH4: IOUT")

for i in range(3): print()





"""COMMS ADDRESS"""
ac_source_address = 5
source_power_meter_address = 1 
load_power_meter_address = 2
eload_address = 8
scope_address = "10.125.10.224"

"""IMPORT DEPENDENCIES"""
import sys
import pyautogui
from time import sleep, time
from powi.equipment import ACSource, PowerMeter, ElectronicLoad, Oscilloscope
from powi.equipment import headers, create_folder, footers, waveform_counter, soak, convert_argv_to_int_list
from powi.equipment import *
from time import sleep, time
import os
import cv2
waveform_counter = 0

"""INITIALIZE EQUIPMENT"""


print("Intialize equipment.")
# ac = ACSource(ac_source_address)
# eload = ElectronicLoad(eload_address)
# scope = Oscilloscope(scope_address)
# # pms = PowerMeter(source_power_meter_address)
# # pml = PowerMeter(load_power_meter_address)

"""USER INPUT"""
LED = convert_argv_to_int_list(sys.argv[1]) # 46, 36, 24V
vin_list = convert_argv_to_int_list(sys.argv[2]) # [120, 230, 277] Vac
# cp = int(sys.argv[3])
cp = 0
# I2C_COMMAND_LIST = convert_argv_to_int_list(sys.argv[3])
I2C_COMMAND_LIST = [0, 50, 100, 150, 200, 250, 254]
test_list = convert_argv_to_int_list(sys.argv[3]) # 1 - First Power Up and No AC Reset Test, # 2 - Transient Down
dim_trigger_source = 3
iout_oscilloscope_channel = 4

print()
print(f"LED: {LED} V")
print(f"VIN LIST: {vin_list}Vac")
print(f"CP: {cp}")
print(f"TEST LIST: {test_list}")
print()

"""INITIAL SCOPE SETTINGS"""
def initial_scope_settings():
    # scope.record_length(50E6)   # 50 MSa
    # scope.time_scale(0.5)       # 0.5 s/div
    # scope.time_position(10)     # 10%

    print("time_scale: 0.5 s/div")
    print("time_position: 10%")

    print("Initial trigger")
    print()
    # scope.init_trigger(trigger_source=dim_trigger_source, trigger_level=1.5, trigger_slope="NEG")




"""RELAY OPTIONS"""

print("Initialize RELAYS.")
print()

# from powi.equipment import *

# board = pyfirmata.Arduino('COM6')
# iterator = util.Iterator(board)
# iterator.start()

# RELAY1 = board.get_pin('d:10:o')
# RELAY2 = board.get_pin('d:9:o')
# RELAY3 = board.get_pin('d:8:o')

# def led_46V():
#     # print("46V LED")
#     RELAY1.write(1)
#     RELAY2.write(0)
#     RELAY3.write(0)

# def led_36V():
#     # print("36V LED")
#     RELAY1.write(0)
#     RELAY2.write(1)
#     RELAY3.write(0)

# def led_24V():
#     # print("24V LED")
#     RELAY1.write(0)
#     RELAY2.write(0)
#     RELAY3.write(1)

# def NL():
#     # print("NL")
#     RELAY1.write(0)
#     RELAY2.write(0)
#     RELAY3.write(0)


"""DEFAULT FUNCTIONS"""

def discharge_output():
    # ac.turn_off()
    # eload.channel[1].cc = 1
    # eload.channel[1].turn_on()
    # eload.channel[2].cc = 1
    # eload.channel[2].turn_on()
    sleep(2)
    # eload.channel[1].turn_off()
    # eload.channel[2].turn_off()
    # print("Discharge output for 2s.")

"""GENERAL AUTOGUI FUNCTIONS"""
def give_loc(filename='binno_tab'):
    return pyautogui.locateOnScreen(f'image_autogui_library\{filename}.png', confidence=0.8)

def click(coords):
    pyautogui.moveTo(coords)
    pyautogui.click()

def change_value(coords, val):
    click(coords)
    pyautogui.press('backspace', presses=3)
    pyautogui.write(f"{val}")

def give_loc_rel(filename='binno_tab', relative_width=0.75):
    x,y,w,h = give_loc(filename)
    x = x + relative_width*w
    y = y + h/2
    return x,y

def add_relative_y(anchor, add_relative_y=0.1):
    x,y = anchor
    y += add_relative_y
    return x,y

"""I2C Autogui Variables"""

# print("Initialize I2C Autogui Variables.")
print()

binno_tab = give_loc('binno_tab')

i2c_send_textbox = give_loc_rel('i2c_send', 0.5)
i2c_send_write = give_loc_rel('i2c_send', 0.75)

i2c_cp_dropdown = give_loc_rel('i2c_cp_setting', 0.56)
i2c_cp_0 = add_relative_y(i2c_cp_dropdown, add_relative_y=24)
i2c_cp_1 = add_relative_y(i2c_cp_0, add_relative_y=18)
i2c_cp_2 = add_relative_y(i2c_cp_1, add_relative_y=18)
i2c_cp_3 = add_relative_y(i2c_cp_2, add_relative_y=18)
i2c_cp_write = give_loc_rel('i2c_cp_setting', 0.75)

initialize_com_port = give_loc('initialize_com_port')

"""I2C Autogui Functions"""
def change_cp(cp):
    click(i2c_cp_dropdown)
    if cp == 0: click(i2c_cp_0)
    elif cp == 1: click(i2c_cp_1)
    elif cp == 2: click(i2c_cp_2)
    elif cp == 3: click(i2c_cp_3)
    else: input("Invalid CP!")
    click(i2c_cp_write)

def set_i2c_bit(cmdbit):
    change_value(i2c_send_textbox, cmdbit)
    click(i2c_send_write)

def i2c_dimming(start=0, end=254, voltage=120, frequency=60, LED=46, cp=0,
                    condition="First Power Up Discharged Cout"):
    
    global waveform_counter
    global total_waveform_counter

    # print("Turn on AC.")
    # ac.voltage = voltage
    # ac.frequency = frequency
    # ac.turn_on()

    # print("Set CP.")
    change_cp(cp)

    # print(f"Set START I2C bit: {start}")
    set_i2c_bit(start)

    # print("Sleep for 2s.")
    sleep(2)

    # print("Set CP.")
    change_cp(cp)
    
    # print("Run Single.")
    # scope.run_single()

    # print("Sleep for 1.5s.")
    sleep(2)

    # print(f"Set END I2C bit: {end}")
    # set_i2c_bit(end)

    # print("Sleep 10s.")
    sleep(10)

    # print("Scope stop.")
    # scope.stop()

    # print("Sleep 1s.")
    sleep(1)

    filename = f'{load}V, {voltage}Vac {frequency}Hz, {start}-{end} bit, I2C Dimming, {condition}.png'
    
    # print("Get screenshot.")
    # print("Incremet waveform counter.")
    # scope.get_screenshot(filename, waveforms_folder)
    print(filename)
    print()
    waveform_counter += 1
    total_waveform_counter += 1
    

def i2c_main():
    
    discharge_output()

    # initial_dim_autogui() # to minimize testing time

    initial_scope_settings()

    """First Power Up and No AC Reset Test"""

    

    test = 1
    if test in test_list:
        
        print("Test 1: First Power Up and No AC Reset Test")

        """TRANSIENT UP"""
        I2C_COMMAND_LIST.sort()
        print(I2C_COMMAND_LIST)
        print()
        for i in range(len(I2C_COMMAND_LIST)):
            for cmdbit in I2C_COMMAND_LIST:
                if I2C_COMMAND_LIST[i] < cmdbit:

                    # print(f"Transient Up: {I2C_COMMAND_LIST[i]}-{cmdbit}")
                    start = I2C_COMMAND_LIST[i]
                    end = cmdbit
                    
                    """Iout Channel Scale Adjuster"""
                    
                    # print("Adjust Iout Channel")
                    if end <= 150:
                        iout_channel_scale = 0.04
                        # print("Iout scale: 40 mA/div ")
                        # scope.channel_scale(iout_oscilloscope_channel, iout_channel_scale)
                    else:
                        # print("Iout scale: 200 mA/div ")
                        iout_channel_scale = 0.2
                        # scope.channel_scale(iout_oscilloscope_channel, iout_channel_scale)

                    i2c_dimming(start, end, vin, freq, load, cp, "First Power Up (Discharged Cout)")
                    # print()
                    i2c_dimming(start, end, vin, freq, load, cp, "No AC reset (Not Discharged Cout)")
                    print()

                    discharge_output()
                else: pass

    print()
    test = 2
    if test in test_list:

        print("Test 2: Transient Down")

        """TRANSIENT DOWN"""
        I2C_COMMAND_LIST.sort(reverse=True)
        print(I2C_COMMAND_LIST)
        print()
        for i in range(len(I2C_COMMAND_LIST)):
            for cmdbit in I2C_COMMAND_LIST:
                if I2C_COMMAND_LIST[i] > cmdbit:
                    # print(f"Transient Down: {I2C_COMMAND_LIST[i]}-{cmdbit}")
                    start = I2C_COMMAND_LIST[i]
                    end = cmdbit

                    """Iout Channel Scale Adjuster"""
                    
                    # print("Adjust Iout Channel")
                    if start <= 150:
                        iout_channel_scale = 0.04
                        # print("Iout scale: 40 mA/div ")
                        # scope.channel_scale(iout_oscilloscope_channel, iout_channel_scale)
                    else:
                        # print("Iout scale: 200 mA/div ")
                        iout_channel_scale = 0.2
                        # scope.channel_scale(iout_oscilloscope_channel, iout_channel_scale)

                    i2c_dimming(start, end, vin, freq, load, cp, "Transient Down")

                    # print()

                    discharge_output()

                else: pass

"""MAIN"""
discharge_output()
print()

print("START OF TEST...")
click(initialize_com_port)

for load in LED:

    for vin in vin_list:
        
        if vin >= 180 and vin < 277: freq = 50
        else: freq = 60

        print(f"Set Vin = {vin} Vac, {freq} Hz")

        test = f"I2C Dimming {load}V {vin}Vac"
        waveforms_folder = f'waveforms/{test}'

        print(f"Change LED load to {load}V.")
        # if LED == '46':
        #     led_46V()
        #     print("46V")
        # elif LED == '36': led_36V()
        # elif LED == '24': led_24V()
        # elif LED == '0': NL()
        # else: print("Invalid LED.")

        headers(test)
        i2c_main()
        footers(waveform_counter)
        waveform_counter = 0

print("END OF TEST...")

print()

end_timestamp = datetime.now()

print(f"START: {start_timestamp}")
print(f"END: {end_timestamp}")

print(f"TEST TIME: {end_timestamp-start_timestamp}")
print(f"TOTAL WAVEFORMS CAPTURED: {total_waveform_counter}")