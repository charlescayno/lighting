
from filemanager import path_maker
from openpyxl import Workbook, load_workbook
import pandas as pd
from powi.equipment import excel_to_df, df_to_excel, image_to_excel, col_row_extractor, get_anchor, create_chartsheet
import os

led_list = [48,36,24]
vin_list = [90,115,230,277]
dim_list = [0.0, 0.1, 0.2 ,0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10]

waveforms_folder = f'C:/Users/ccayno/Desktop/DER/DER-727/Test Data/Final Data/Dimming/Flicker Test/csv file/'
path = path_maker(f'{waveforms_folder}')

dir_list = os.listdir(path)
print(dir_list)
print()

dict = {}


import pandas as pd



for led in led_list:
    for vin in vin_list:
        for dim in dim_list:
            filename = f"{led}V_{vin}Vac_{dim}V.csv"
            dst = path + filename
            f_name, f_ext = os.path.splitext(dst)

            read_file = pd.read_csv (dst)
            read_file.to_excel (f'{path}{f_name}.xlsx', index = None, header=True)

            input()

            dict[f_name] = excel_to_df(dst, f_name, 'D15', 'F814')

            print(dict)
            input()

# dst = "C:/Users/ccayno/Desktop/DER/DER-727/Test Data/Final Data/Dimming/Resistor Dimming/Resistor Dimming_1.xlsx"



# wb = load_workbook(dst)
# df_48V_90Vac = excel_to_df(dst, '48', 'B2', 'P20')
# df = df_48V_90Vac
# create_chartsheet(wb, df, title="Efficiency (%)", style=1, legend=None, x_title='Input Voltage (VAC)', y_title='Efficiency',
#                     x_anchor="B1", y_anchor="N1",
#                     x_min_scale = 90, x_max_scale = 277,  x_major_unit = 20, x_minor_unit = 10,
#                     y_min_scale = 85, y_max_scale = 100, y_major_unit = 1, y_minor_unit = 1)
# wb.save(dst)




# create_chartsheet(wb, df, title="Power Factor", style=2, legend=None, x_title='Input Voltage (VAC)', y_title='Power Factor',
#                     x_anchor="B1", y_anchor="G1",
#                     x_min_scale = 90, x_max_scale = 277, x_major_unit = 20, x_minor_unit = 10,
#                     y_min_scale = 0.9, y_max_scale = 1, y_major_unit = 0.01, y_minor_unit = 0.01)

# create_chartsheet(wb, df, title="THD (%)", style=3, legend=None, x_title='Input Voltage (VAC)', y_title='THD (%)',
#                     x_anchor="B1", y_anchor="H1",
#                     x_min_scale = 90, x_max_scale = 277, x_major_unit = 20, x_minor_unit = 10,
#                     y_min_scale = 0, y_max_scale = 15, y_major_unit = 1, y_minor_unit = 1)

# create_chartsheet(wb, df, title="Output Current (mA)", style=4, legend=None, x_title='Input Voltage (VAC)', y_title='Output Current (mA)',
#                     x_anchor="B1", y_anchor="J1",
#                     x_min_scale = 90, x_max_scale = 277, x_major_unit = 20, x_minor_unit = 10,
#                     y_min_scale = iout*0.95*1000, y_max_scale = iout*1.05*1000, y_major_unit = iout*0.05*100, y_minor_unit = iout*0.05*100)

# create_chartsheet(wb, df, title="Output Voltage (V)", style=5, legend=None, x_title='Input Voltage (VAC)', y_title='Output Voltage (V)',
#                     x_anchor="B1", y_anchor="I1",
#                   /  x_min_scale = 90, x_max_scale = 277, x_major_unit = 20, x_minor_unit = 10,
#                     y_min_scale = vout*0.95, y_max_scale = vout*1.05, y_major_unit = vout*0.005, y_minor_unit = vout*0.005)

# create_chartsheet(wb, df, title="Input Power (W)", style=6, legend=None, x_title='Input Voltage (VAC)', y_title='Input Power (W)',
#                     x_anchor="B1", y_anchor="F1",
#                     x_min_scale = 90, x_max_scale = 277, x_major_unit = 20, x_minor_unit = 10,
#                     y_min_scale = 0, y_max_scale = 90, y_major_unit = 10, y_minor_unit = 10)