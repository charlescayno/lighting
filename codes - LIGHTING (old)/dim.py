from xml.sax.handler import feature_external_ges
from powi.equipment import Keithley_DC_2230G
import pyautogui
from time import sleep, time
import pandas as pd
import os
import shutil
import numpy as np
# C:\Users\ccayno\Anaconda3\Lib\site-packages\powi
# C:\Users\ccayno\Desktop\Audible Automation

def exists(filename):
    return os.path.isfile(filename)

def get_file_size(filename):
    if exists():        
        statinfo = os.stat(filename)
        return statinfo.st_size
    return 0

def file_to_df(filename, column_name):
    df = pd.read_excel(filename, skiprows=3)
    df.set_index('Hz', inplace=True)
    df.index.name = 'Frequency'
    df.columns = [column_name]
    return df

def append_df_to_file(filename, df):
    if exists(filename):
        base_df = pd.read_excel(filename, index_col=0)
        base_df = pd.concat([base_df, df], axis=1)
    else:
        base_df = df
    base_df.to_excel(filename)

def copy_file(src, dst):
    shutil.copyfile(src, dst)

def delete_file(filename):
    try:
        os.remove(filename)
    except:
        pass


dc = Keithley_DC_2230G('4')

def set_dim_voltage(dim):
    dc.set_volt_curr(channel='CH1', voltage=dim, current=1.0)
    dc.channel_state(channel='CH1', state='ON')
    sleep(5)

def operation(dim, raw_acoustic_file):
    
    dim = float(f"{dim:.1f}")
    print(f"dim: {dim} V")
    set_dim_voltage(dim)

    sleep(1)
    pyautogui.moveTo(button_position)
    sleep(1)
    pyautogui.click()

    while not exists(raw_acoustic_file):
        sleep(1)
    raw_df = file_to_df(filename=raw_acoustic_file, column_name=f'{dim}V')
    append_df_to_file(output_file, raw_df)
    delete_file(raw_acoustic_file)
    sleep(3)


                
start = time()

raw_acoustic_file = 'FrequencyResponse.xlsx'
button_position = (30, 135)



### ENTER USER INPUTS HERE
vin = float(input(">> Enter input voltage (VAC): "))
voltages = [vin]
if vin == 230:
    freq = 50
else:
    freq = 60
frequencies = [freq]
load = input(">> ENTER LED load (V): ")
led_list = [load]
### ENTER USER INPUTS HERE

for led in led_list:
    
    input(f"Set LED load to {led} V.")

    for voltage, frequency in zip(voltages, frequencies):
        
        input(f">> Turn on power supply to {voltage} Vac, {frequency} Hz. Press ENTER to continue...")
        sleep(5)

        output_file = f'{led}V {voltage}VAC.xlsx'

        delete_file(output_file)

        for dim in np.arange(0, 0.5, 0.1):
        
            operation(dim, raw_acoustic_file)
            
        for dim in np.arange(1, 10.5, 0.5):
            
            operation(dim, raw_acoustic_file)

        break

end = time()
print(f'Elapsed: {end-start} s')
