import pyautogui
from time import sleep, time
import os
import shutil
        

class AutoguiCalibrate():

    def __init__(self):
        import os.path
        from os import path
        
        if os.path.isfile('coordinates.txt'):
            self.calibration_status = True
        else:
            self.calibration_status = False
            print("This is an ATE autogui configurator.")
            print("Design to determine the exact location of specific points to be automated.\n\n")
            print("Calibration begins now..\n\n")   
            with open('coordinates.txt', 'w') as f: pass
            self.save_coordinates()

    def initialize_dictionary(self):
        self.dict = {}

    def get_coordinates(self, target):
        input(f">> {target}. Press ENTER to get coordinates.")
        x,y = pyautogui.position()
        self.dict[target] = x,y

    def get_coordinates_manual(self, target):
        input(f">> {target}. Press ENTER to get coordinates.")
        x,y = pyautogui.position()
        return x,y

    def ac_control(self):
        self.ac_on = self.get_coordinates('AC_ON')
        self.ac_off = self.get_coordinates('AC_OFF')
        self.vrms_gui = self.get_coordinates('VRMS')
        self.freq_gui = self.get_coordinates('FREQ')

    def soak_settings(self):
        self.soak_time = self.get_coordinates('SOAK_TIME')
        self.delay_per_line = self.get_coordinates('DELAY_PER_LINE')
        self.delay_per_load = self.get_coordinates('DELAY_PER_LOAD')
        self.integration_time = self.get_coordinates('INTEGRATION')

        self.run_test = self.get_coordinates('RUN_TEST')
        self.abort = self.get_coordinates('ABORT')

    def test_selection(self):
        # self.select_test = self.get_coordinates('SELECT_TEST')
        self.select_test_option = self.get_coordinates('SELECT_TEST_OPTION')
        self.select_test_digital_control = self.get_coordinates('SELECT_TEST_DIGITAL_CONTROL')

    def digital_control(self):
        # initialize digital control settings
        self.test_selection()
        self.settings = self.get_coordinates("SETTINGS")
        self.initialize_com_port = self.get_coordinates('INITIALIZE_COM_PORT')
        self.activate_load_settings_for_ate = self.get_coordinates('ACTIVATE_LOAD_SETTINGS_FOR_ATE')
        self.activate_load_settings_for_ate_ok_button = self.get_coordinates('ACTIVATE_LOAD_SETTINGS_FOR_ATE_OK_BUTTON')
        self.activate_load_settings_for_ate_test_type = self.get_coordinates('ACTIVATE_LOAD_SETTINGS_FOR_ATE_TEST_TYPE')
        self.multiprotocol_control_exit_button = self.get_coordinates('MULTIPROTOCOL_CONTROL_EXIT_BUTTON')
        self.load_settings_exit_button = self.get_coordinates('LOAD_SETTINGS_EXIT_BUTTON')
        # self.multiprotocol_control('binno')
        # self.multiprotocol_control('rheostat')
        # self.multiprotocol_control('analog')
        self.multiprotocol_control('pwm')
        # self.test_complete_ok_button = self.get_coordinates('TEST_COMPLETE_OK_BUTTON')

    def multiprotocol_control(self, type='analog'):
        
        if type == "analog":
            self.analog_tab = self.get_coordinates('ANALOG_TAB')
            self.analog_manual_change = self.get_coordinates('ANALOG_MANUAL_CHANGE')
            self.analog_set_button = self.get_coordinates('ANALOG_SET_BUTTON')
            self.analog_settings_start = self.get_coordinates('ANALOG_SETTINGS_START')
            self.analog_settings_end = self.get_coordinates('ANALOG_SETTINGS_END')
            self.analog_settings_step_size = self.get_coordinates('ANALOG_SETTINGS_STEP_SIZE')

        if type == 'binno':
            self.binno_tab = self.get_coordinates('BINNO_TAB')
            
            self.i2c_send_write_bit = self.get_coordinates('I2C_SEND_WRITE_BIT')
            self.i2c_send_write_button = self.get_coordinates('I2C_SEND_WRITE_BUTTON')
            
            self.i2c_linear_dimming_enable = self.get_coordinates('I2C_LINEAR_DIMMING_ENABLE')
            self.i2c_linear_dimming_disable = self.get_coordinates('I2C_LINEAR_DIMMING_DISABLE')
            self.i2c_linear_dimming_type_write = self.get_coordinates('I2C_LINEAR_DIMMING_TYPE_WRITE')
            
            self.i2c_constant_power_setting_textbox = self.get_coordinates('I2C_CONSTANT_POWER_SETTING_TEXTBOX')
            self.i2c_constant_power_setting_write = self.get_coordinates('I2C_CONSTANT_POWER_SETTING_WRITE')

            self.i2c_dimming_curve_settings_start_bit = self.get_coordinates('I2C_DIMMING_CURVE_SETTINGS_START_BIT')
            self.i2c_dimming_curve_settings_end_bit = self.get_coordinates('I2C_DIMMING_CURVE_SETTINGS_END_BIT')
            self.i2c_dimming_curve_settings_step_size = self.get_coordinates('I2C_DIMMING_CURVE_SETTINGS_STEP_SIZE')

        if type == 'rheostat':
            self.rheostat_tab = self.get_coordinates('RHEOSTAT_TAB')
            self.rheostat_manual_change = self.get_coordinates('RHEOSTAT_MANUAL_CHANGE')
            self.rheostat_set_resistance_button = self.get_coordinates('RHEOSTAT_SET_RESISTANCE_BUTTON')
            self.rheostat_settings_start = self.get_coordinates('RHEOSTAT_SETTINGS_START')
            self.rheostat_settings_end = self.get_coordinates('RHEOSTAT_SETTINGS_END')
            self.rheostat_settings_step_size = self.get_coordinates('RHEOSTAT_SETTINGS_STEP_SIZE')

        if type == 'pwm':
            self.pwm_tab = self.get_coordinates('PWM_TAB')

            self.pwm_manual_frequency_textbox = self.get_coordinates('PWM_MANUAL_FREQUENCY_TEXTBOX')
            self.pwm_manual_duty_textbox = self.get_coordinates('PWM_MANUAL_DUTY_TEXTBOX')
            self.pwm_manual_duty_write_button = self.get_coordinates('PWM_MANUAL_SETTINGS_WRITE_BUTTON')

            self.pwm_settings_start = self.get_coordinates('PWM_SETTINGS_START')
            self.pwm_settings_end = self.get_coordinates('PWM_SETTINGS_END')
            self.pwm_settings_size = self.get_coordinates('PWM_SETTINGS_STEP')


    def ANALOG_DIMMING(self, start, end, step, soak, delay_per_line, delay_per_load):
        
        self.change_value("SOAK_TIME", soak)
        self.change_value("DELAY_PER_LINE", delay_per_line)
        self.change_value("DELAY_PER_LOAD", delay_per_load)
        
        self.click('SETTINGS')
        self.click('ANALOG_TAB')
        self.click('INITIALIZE_COM_PORT')

        self.change_value('ANALOG_MANUAL_CHANGE', start)
        self.click('ANALOG_SET_BUTTON')

        self.change_value('ANALOG_SETTINGS_START', start)
        self.change_value('ANALOG_SETTINGS_END', end)
        self.change_value('ANALOG_SETTINGS_STEP_SIZE', step)

        self.click('ACTIVATE_LOAD_SETTINGS_FOR_ATE')
        self.click('ACTIVATE_LOAD_SETTINGS_FOR_ATE_OK_BUTTON')
        self.change_value('ACTIVATE_LOAD_SETTINGS_FOR_ATE_TEST_TYPE', "0-10V")
        self.click('MULTIPROTOCOL_CONTROL_EXIT_BUTTON')
        self.click('LOAD_SETTINGS_EXIT_BUTTON')

    def initialize_analog_dimming(self, start):

        self.click('SETTINGS')
        self.click('ANALOG_TAB')
        self.click('INITIALIZE_COM_PORT')

        self.change_value('ANALOG_MANUAL_CHANGE', start)
        self.click('ANALOG_SET_BUTTON')

        self.click('MULTIPROTOCOL_CONTROL_EXIT_BUTTON')

    def I2C_DIMMING(self, start, end, step, cp, i2c_type, soak, delay_per_line, delay_per_load):
        
        # soak time
        self.change_value("SOAK_TIME", soak)
        self.change_value("DELAY_PER_LINE", delay_per_line)
        self.change_value("DELAY_PER_LOAD", delay_per_load)
        
        # initialization
        self.click('SETTINGS')
        self.click('BINNO_TAB')
        self.click('INITIALIZE_COM_PORT')

        # set i2c dimming settings
        self.change_value('I2C_DIMMING_CURVE_SETTINGS_START_BIT', start)
        self.change_value('I2C_DIMMING_CURVE_SETTINGS_END_BIT', end)
        self.change_value('I2C_DIMMING_CURVE_SETTINGS_STEP_SIZE', step)

        # initialize first bit I2C
        self.change_value('I2C_SEND_WRITE_BIT', start)
        self.click('I2C_SEND_WRITE_BUTTON')

        # change CP settings
        self.change_value('I2C_CONSTANT_POWER_SETTING_TEXTBOX', cp)
        self.click('I2C_CONSTANT_POWER_SETTING_WRITE')

        # change I2C type settings
        if i2c_type == 'logarithmic':
            self.click('I2C_LINEAR_DIMMING_DISABLE')       
        elif i2c_type == 'linear':
            self.click('I2C_LINEAR_DIMMING_ENABLE')
        else: input('Invalid I2C Type. Enter either linear or logarithmic.')
        self.click('I2C_LINEAR_DIMMING_TYPE_WRITE')

        # activate load settings
        self.click('ACTIVATE_LOAD_SETTINGS_FOR_ATE')
        self.click('ACTIVATE_LOAD_SETTINGS_FOR_ATE_OK_BUTTON')
        self.change_value('ACTIVATE_LOAD_SETTINGS_FOR_ATE_TEST_TYPE', "I2C")
        self.click('MULTIPROTOCOL_CONTROL_EXIT_BUTTON')
        self.click('LOAD_SETTINGS_EXIT_BUTTON')      

    def PWM_DIMMING(self, start, end, step, frequency, soak, delay_per_line, delay_per_load):
        
        # soak time
        self.change_value("SOAK_TIME", soak)
        self.change_value("DELAY_PER_LINE", delay_per_line)
        self.change_value("DELAY_PER_LOAD", delay_per_load)
        
        # initialization
        self.click('SETTINGS')
        self.click('PWM_TAB')
        self.click('INITIALIZE_COM_PORT')

        # set PWM dimming settings
        self.change_value('PWM_SETTINGS_START', start)
        self.change_value('PWM_SETTINGS_END', end)
        self.change_value('PWM_SETTINGS_STEP', step)

        # initialize PWM
        self.change_value('PWM_MANUAL_FREQUENCY_TEXTBOX', frequency)
        self.change_value('PWM_MANUAL_DUTY_TEXTBOX', start)
        self.click('PWM_MANUAL_SETTINGS_WRITE_BUTTON')

        # activate load settings
        self.click('ACTIVATE_LOAD_SETTINGS_FOR_ATE')
        self.click('ACTIVATE_LOAD_SETTINGS_FOR_ATE_OK_BUTTON')
        self.change_value('ACTIVATE_LOAD_SETTINGS_FOR_ATE_TEST_TYPE', "I2C")
        self.click('MULTIPROTOCOL_CONTROL_EXIT_BUTTON')
        self.click('LOAD_SETTINGS_EXIT_BUTTON') 


    def run_test(self):
        self.click('RUN_TEST')

    def ac_turn_on(self, vin):

        if vin == 230: freq = 50
        elif vin == 265: freq = 50
        else: freq = 60

        self.change_value('VRMS', vin)
        self.change_value('FREQ', freq)
        self.click('AC_ON')

    def ac_turn_off(self, vin, freq):
        self.click('AC_OFF')


    # recalibrate -> save_coordinates -> calibrate_autogui() -> initialization


    def calibrate_autogui(self):
        self.initialize_dictionary()
        self.ac_control()
        self.soak_settings()
        self.digital_control()
        

    def save_coordinates(self):
        if self.calibration_status == False:
            self.calibrate_autogui()
            with open('coordinates.txt', 'w') as f:
                f.write(str(self.dict))

    def recalibrate(self):
        print("Recalibrating...")
        self.calibration_status = False
        self.save_coordinates()
        print("Recalibration complete.")

    def load_coordinates(self):
        with open('coordinates.txt', 'r') as f:
            str_dict = f.read()
            self.dictionary = eval(str_dict)
            return self.dictionary

    def moveTo(self, target):
        if type(target) == type((0,0)):
            pyautogui.moveTo(target)
        else:
            self.dictionary = self.load_coordinates()
            x,y = self.dictionary[target]
            pyautogui.moveTo(x,y)

    def click(self, target):
        self.moveTo(target)
        pyautogui.click()

    def change_value(self, target, value):
        self.click(target)
        self.ctrl_a()
        pyautogui.press('backspace', presses=1)
        pyautogui.write(f"{value}")
    
    def change_rheostat_settings(self, start, end, step_size):
        self.dictionary = self.load_coordinates()
        self.change_value(self.dictionary['RHEOSTAT_SETTINGS_START'], start)
        self.change_value(self.dictionary['RHEOSTAT_SETTINGS_END'], end)
        self.change_value(self.dictionary['RHEOSTAT_SETTINGS_STEP_SIZE'], step_size)
    
    def change_rheostat(self, value):
        self.dictionary = self.load_coordinates()
        self.click('INITIALIZE_COM_PORT')
        self.change_value(self.dictionary['RHEOSTAT_MANUAL_CHANGE'], value)
        self.click('RHEOSTAT_SET_RESISTANCE_BUTTON')
    
    def ctrl_a(self):
        pyautogui.keyDown('ctrl')
        sleep(0.1)
        pyautogui.press('a')
        sleep(0.1)
        pyautogui.keyUp('ctrl')

    def alt_tab(self):
        pyautogui.keyDown('alt')
        sleep(0.2)
        pyautogui.press('tab')
        sleep(0.2)
        pyautogui.keyUp('alt')

    def esc(self):
        pyautogui.press('esc')

    def enter(self):
        pyautogui.press('enter')

    def test_complete(self):
        self.click('TEST_COMPLETE_OK_BUTTON')

    # def screenshot(self):
    #     im = pyautogui.screenshot('my_screenshot.png')
    #     shutil.move(f'{os.getcwd()}\my_screenshot.png', f'{os.getcwd()}\images\my_screenshot.png')


    # def click(self, filename):
    #     coords = pyautogui.locateOnScreen(f'images/{filename}.jpg', confidence=0.95)
    #     pyautogui.moveTo(coords)
    #     pyautogui.click()


    # def alert1(self, filename):
    #     self.coords = pyautogui.locateOnScreen(f'images/{filename}.png', grayscale=False, confidence=0.95)
    #     # pyautogui.moveTo(self.coords)
    #     if self.coords != None:
    #         ws.PlaySound("sfx/coins.wav", ws.SND_ASYNC)
    #         sleep(0.1)
    
    # def alert(self, filename):
    #     self.coords = pyautogui.locateOnScreen(f'images/{filename}.jpg', grayscale=False, confidence=0.95)
    #     # pyautogui.moveTo(self.coords)
    #     if self.coords != None:
    #         ws.PlaySound("sfx/coins.wav", ws.SND_ASYNC)
    #         sleep(0.1)