import pyautogui
from time import sleep, time
import pandas as pd
import os
import shutil
#from powi.equipment import ACSource
#from powi.equipment import ElectronicLoad
from powi.equipment import Keithley_DC_2230G

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

start = time()

raw_acoustic_file = 'FrequencyResponse.xlsx'
iterations = 300
button_position = (30, 135)

# ac = ACSource(5)
# eload = ElectronicLoad(8)

### ENTER USER INPUTS HERE
voltages = [230]
frequencies = [50]
max_current = 3
eload_channel = 1
### ENTER USER INPUTS HERE

for voltage, frequency in zip(voltages, frequencies):
    output_file = f'{voltage}VAC.xlsx'

    delete_file(output_file)

##    # special condition for DER_831
##    eload.channel[7].cc = 0.004
##    eload.channel[7].turn_on()
    
    #ac.voltage = voltage
    #ac.frequency = frequency
    #ac.turn_on()

    for i in range(1, iterations+1):
        current = max_current * (i / iterations)
        print('current', current)
        #eload.channel[eload_channel].cc = current
        #eload.channel[eload_channel].turn_on()
        
        sleep(1)
        pyautogui.moveTo(button_position)
        sleep(1)
        pyautogui.click()

        while not exists(raw_acoustic_file):
            sleep(1)
        raw_df = file_to_df(filename=raw_acoustic_file, column_name=f'{int(100*i/iterations)}%')
        append_df_to_file(output_file, raw_df)
        delete_file(raw_acoustic_file)

    break

#ac.turn_off()
#eload.channel[eload_channel].turn_off()

end = time()
print(f'Elapsed: {end-start})')
