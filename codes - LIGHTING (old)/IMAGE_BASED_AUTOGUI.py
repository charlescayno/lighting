print("TO SLOW")


from sys import int_info
import pyautogui
from time import time, sleep
import cv2
sleep(2)
start=time()

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







click(initialize_com_port)

click(binno_tab)
set_i2c_bit(123)
change_cp(1)
sleep(1)
change_cp(2)
sleep(1)
change_cp(3)
sleep(1)
change_cp(0)





end=time()
total_time = end-start
print(f'test time: {total_time:.2f} secs. / {total_time/60:.2f} mins.')