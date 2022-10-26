
print("COMPATIBLE ONLY WITH CHARLES' PERSONAL LAPTOP")
print("CMC | 23-Jun-2021")

from powi.equipment import convert_argv_to_int_list
from tkinter.constants import END
import pyautogui
from time import time, sleep
import sys

start = int(sys.argv[1])
end = int(sys.argv[2])
step = int(sys.argv[3])
cp = int(sys.argv[4])
test_list = convert_argv_to_int_list(sys.argv[5])

"""RELAY OPTIONS"""

from powi.equipment import *

board = pyfirmata.Arduino('COM8')
iterator = util.Iterator(board)
iterator.start()

RELAY1 = board.get_pin('d:10:o')
RELAY2 = board.get_pin('d:9:o')
RELAY3 = board.get_pin('d:8:o')

def led_46V():
    # print("46V LED")
    RELAY1.write(1)
    RELAY2.write(0)
    RELAY3.write(0)

def led_36V():
    # print("36V LED")
    RELAY1.write(0)
    RELAY2.write(1)
    RELAY3.write(0)

def led_24V():
    # print("24V LED")
    RELAY1.write(0)
    RELAY2.write(0)
    RELAY3.write(1)

def NL():
    # print("NL")
    RELAY1.write(0)
    RELAY2.write(0)
    RELAY3.write(0)

# led_46V()
# input()
# led_36V()
# input()
# led_24V()
# input()
# NL()
# input()

"""PYAUTOGUI CALIBRATION"""
# x, y = pyautogui.position()
# print(x,y)
# input()

# x, y = pyautogui.position()
# print(x,y)
# input()

# x, y = pyautogui.position()
# print(x,y)
# input()

# x, y = pyautogui.position()
# print(x,y)
# input()

# x, y = pyautogui.position()
# print(x,y)
# input()

# x, y = pyautogui.position()
# print(x,y)
# input()

# x, y = pyautogui.position()
# print(x,y)
# input()

# x, y = pyautogui.position()
# print(x,y)
# input()

"""GENERAL AUTOGUI FUNCTIONS"""

def move_click(x, y):
    pyautogui.moveTo(x, y)
    pyautogui.click()

def alt_tab():
    pyautogui.keyDown('alt')
    sleep(0.2)
    pyautogui.press('tab')
    sleep(0.2)
    pyautogui.keyUp('alt')

def ctrl_a():
    pyautogui.keyDown('ctrl')
    sleep(0.2)
    pyautogui.press('a')
    sleep(0.2)
    pyautogui.keyUp('ctrl')

def ctrl_s():
    pyautogui.keyDown('ctrl')
    sleep(0.2)
    pyautogui.press('s')
    sleep(0.2)
    pyautogui.keyUp('ctrl')

def change_value(x, y, val):
    move_click(x,y)
    pyautogui.press('backspace', presses=3)
    pyautogui.write(f"{val}")

"""SPECIFIC ATE AUTOGUI FUNCTIONS"""

def change_select_test_to_digital_control():
    move_click(922, 84)
    sleep(0.5)
    move_click(837, 353)
    sleep(0.5)

def change_setting_to_binno():
    #click setting
    move_click(641, 97)
    sleep(0.5)
    #click Binno tab
    move_click(387, 196)
    #initialize COM port
    move_click(447, 723)

def load_settings_for_ate():
    
    move_click(863, 727)    # activate load settings for ATE
    sleep(0.5)
    move_click(1222, 463)   # close pop-up
    move_click(1049, 792)   # drop-down arrow
    move_click(1049, 820)   # choose I2C setting
    move_click(1554, 141)   # close setting menu
    sleep(0.8)
    # move_click(655, 448)    # enable load 1
    move_click(1312, 233)   # close load setting menu

def change_i2c_dimming_curve_settings(start, end, step_size, cp):
    
    change_setting_to_binno()
    
    change_value(647, 436, start) # start bit
    change_value(647, 486, end)   # end bit
    change_value(647, 536, step_size)   # step size

    # change initial start bit
    change_value(581, 346, start)
    move_click(654, 346) # write start bit
    sleep(0.25)

    #change cp setting
    move_click(883, 278) # drop down

    if cp == 0: move_click(883, 299)
    elif cp == 1: move_click(883, 318)
    elif cp == 2: move_click(883, 332)
    elif cp == 3: move_click(883, 348)
    else: print("Invalid CP!!!")

    move_click(936, 277) # write CP mode
    sleep(0.25)

    load_settings_for_ate()



def change_ac(vin):
    if vin >= 180 and vin < 277: frequency = 50
    else: frequency = 60

    change_value(350, 86, vin)
    change_value(350, 113, frequency)
    move_click(206, 98)
    sleep(1)

def ac_off():
    move_click(148, 105)
    sleep(1)

def change_soak_settings():
    change_value(1101, 83, 10)
    change_value(1101, 113, 1)
    change_value(1101, 140, 1)
    change_value(1278, 83, 1)

def start_test():
    move_click(1613, 97)

def led_load(led_load):
    global led
    led = led_load
    if led == 46: led_46V()
    elif led == 36: led_36V()
    else: led_24V()

def delete_page():
    move_click(500, 500)
    ctrl_a()
    ctrl_a()
    pyautogui.press('delete')

def click_test_complete():
    move_click(979, 610)
    sleep(0.2)

alt_tab()
sleep(2)

print()
print("Test Done:")
print()

ac_off()
change_select_test_to_digital_control()
change_soak_settings()
led_46V()
sleep(2)


for i in test_list:

    a = 395 + (i-1)*34
    move_click(a, 985)
    sleep(0.5)

    if i in [1, 4, 7]: vin = 120
    elif i in [2, 5, 8]: vin = 230
    else: vin = 277

    if i in [1,2,3]: led_load(46)
    elif i in [4,5,6]: led_load(36)
    else: led_load(24)

    sleep(2)

    print(f"{led}V, {vin}Vac, {start}-to-{end}, CP={cp}")

    # delete_page()

    change_ac(vin)    
    change_i2c_dimming_curve_settings(start=start, end=end, step_size=step, cp=cp)
    start_test()
    sleep(1810) # 30mins
    click_test_complete()
    ac_off()

print()
print("Test Complete.")
ctrl_s()
alt_tab()
