import pyautogui
from time import time, sleep

def alt_tab():
    pyautogui.keyDown('alt')
    sleep(0.2)
    pyautogui.press('tab')
    sleep(0.2)
    pyautogui.keyUp('alt')

def right_click(): 
    pyautogui.click(button='right')


def main():

    alt_tab()

    for i in range(334,529):

        pyautogui.moveTo(943,558)
        right_click()

        pyautogui.moveTo(1000,573)
        pyautogui.click()

        sleep(1)

        pyautogui.write(f"{i}")
        pyautogui.press('enter')

        pyautogui.moveTo(1205,303)
        pyautogui.click()

        sleep(5)


main()



