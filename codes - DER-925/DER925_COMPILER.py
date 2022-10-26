import pandas as pd
import os
from openpyxl import Workbook, load_workbook
import sys

## ENTER UNIT NUMBER ####
unit = sys.argv[1]
#########################




# Python program to return title to result 
# of excel sheet. 
  
# # Returns resul when we pass title. 
# def titleToNumber(s):
#     result = 0; 
#     for B in range(len(s)): 
#         result *= 26; 
#         result += ord(s[B]) - ord('A') + 1; 
#     return result; 


# print(titleToNumber("CDA"))



# point = "B12"

# def col_row_converter(string="B1"):
#     col = ''.join(char for char in string if char.isalpha())
#     col = ord(col)-ord('A')+1
#     row = ''.join(char for char in string if char.isdigit())
#     return col, row
# col, row = col_row_converter()
# print(col, row)



# input()




"""DATAFRAMES"""


# destination file
dest = f'Unit {unit} Data Pack.xlsx'


# source files
file = f'Unit {unit} (PF, THD, EFF, NL, Harmonics).xlsm'
path = f"/DER 925 Reference Data/PF, THD, EFF, NL, Harmonics/Unit {unit}/"
filename = os.getcwd() + path + file

"""PF, THD, EFF, NL, Harmonics"""

# line regulation
df_linereg_48V = pd.read_excel(filename, sheet_name='LineReg 48V', skiprows=4, usecols="E:P", nrows= 12)
df_linereg_36V = pd.read_excel(filename, sheet_name='LineReg 36V', skiprows=4, usecols="E:P", nrows= 12)
df_linereg_24V = pd.read_excel(filename, sheet_name='LineReg 24V', skiprows=4, usecols="E:P", nrows= 12)

# no load, light off
df_nl = pd.read_excel(filename, sheet_name='NL', skiprows=4, usecols="E:P", nrows= 12)
df_nl_lightoff = pd.read_excel(filename, sheet_name='NL lightoff', skiprows=4, usecols="E:P", nrows= 12)

# harmonics 
df_harmonics_120Vac_24V = pd.read_excel(filename, sheet_name='120Vac 24V', skiprows=10, usecols="E:P", nrows= 26)
df_harmonics_230Vac_24V = pd.read_excel(filename, sheet_name='230Vac 24V', skiprows=10, usecols="E:P", nrows= 26)
df_harmonics_277Vac_24V = pd.read_excel(filename, sheet_name='277Vac 24V', skiprows=10, usecols="E:P", nrows= 26)

df_harmonics_120Vac_36V = pd.read_excel(filename, sheet_name='120Vac 36V', skiprows=10, usecols="E:P", nrows= 26)
df_harmonics_230Vac_36V = pd.read_excel(filename, sheet_name='230Vac 36V', skiprows=10, usecols="E:P", nrows= 26)
df_harmonics_277Vac_36V = pd.read_excel(filename, sheet_name='277Vac 36V', skiprows=10, usecols="E:P", nrows= 26)

df_harmonics_120Vac_48V = pd.read_excel(filename, sheet_name='120Vac 48V', skiprows=10, usecols="E:P", nrows= 26)
df_harmonics_230Vac_48V = pd.read_excel(filename, sheet_name='230Vac 48V', skiprows=10, usecols="E:P", nrows= 26)
df_harmonics_277Vac_48V = pd.read_excel(filename, sheet_name='277Vac 48V', skiprows=10, usecols="E:P", nrows= 26)

"""ANALOG DIMMING"""

file = f"Unit {unit}.xlsm"
path = "/DER 925 Reference Data/Analog Dimming/"
filename = os.getcwd() + path + file

df_analog_120Vac_48V = pd.read_excel(filename, sheet_name='48V 120Vac', skiprows=4, usecols="E:P", nrows= 12)
df_analog_230Vac_48V = pd.read_excel(filename, sheet_name='48V 230Vac', skiprows=4, usecols="E:P", nrows= 12)
df_analog_277Vac_48V = pd.read_excel(filename, sheet_name='48V 277Vac', skiprows=4, usecols="E:P", nrows= 12)

df_analog_120Vac_36V = pd.read_excel(filename, sheet_name='36V 120VAC', skiprows=4, usecols="E:P", nrows= 12)
df_analog_230Vac_36V = pd.read_excel(filename, sheet_name='36V 230Vac', skiprows=4, usecols="E:P", nrows= 12)
df_analog_277Vac_36V = pd.read_excel(filename, sheet_name='36V 277Vac', skiprows=4, usecols="E:P", nrows= 12)

df_analog_120Vac_24V = pd.read_excel(filename, sheet_name='24V 120Vac', skiprows=4, usecols="E:P", nrows= 12)
df_analog_230Vac_24V = pd.read_excel(filename, sheet_name='24V 230Vac', skiprows=4, usecols="E:P", nrows= 12)
df_analog_277Vac_24V = pd.read_excel(filename, sheet_name='24V 277Vac', skiprows=4, usecols="E:P", nrows= 12)

"""RESISTOR DIMMING"""

file = f"Unit {unit}.xlsm"
path = "/DER 925 Reference Data/Resistor Dimming/"
filename = os.getcwd() + path + file

df_resistor_120Vac_48V = pd.read_excel(filename, sheet_name='48V 120Vac', skiprows=4, usecols="E:P", nrows= 12)
df_resistor_230Vac_48V = pd.read_excel(filename, sheet_name='48V 230Vac', skiprows=4, usecols="E:P", nrows= 12)
df_resistor_277Vac_48V = pd.read_excel(filename, sheet_name='48V 277Vac', skiprows=4, usecols="E:P", nrows= 12)

df_resistor_120Vac_36V = pd.read_excel(filename, sheet_name='36V 120VAC', skiprows=4, usecols="E:P", nrows= 12)
df_resistor_230Vac_36V = pd.read_excel(filename, sheet_name='36V 230Vac', skiprows=4, usecols="E:P", nrows= 12)
df_resistor_277Vac_36V = pd.read_excel(filename, sheet_name='36V 277Vac', skiprows=4, usecols="E:P", nrows= 12)

df_resistor_120Vac_24V = pd.read_excel(filename, sheet_name='24V 120Vac', skiprows=4, usecols="E:P", nrows= 12)
df_resistor_230Vac_24V = pd.read_excel(filename, sheet_name='24V 230Vac', skiprows=4, usecols="E:P", nrows= 12)
df_resistor_277Vac_24V = pd.read_excel(filename, sheet_name='24V 277Vac', skiprows=4, usecols="E:P", nrows= 12)

"""PWM DIMMING"""

file = f"Unit {unit} (2KHz, 400Hz).xlsm"
path = "/DER 925 Reference Data/PWM Dimming/"
filename = os.getcwd() + path + file

df_pwm_2kHz_120Vac_48V = pd.read_excel(filename, sheet_name='48V 120Vac', skiprows=4, usecols="E:P", nrows= 12)
df_pwm_2kHz_230Vac_48V = pd.read_excel(filename, sheet_name='48V 230Vac', skiprows=4, usecols="E:P", nrows= 12)
df_pwm_2kHz_277Vac_48V = pd.read_excel(filename, sheet_name='48V 277Vac', skiprows=4, usecols="E:P", nrows= 12)

df_pwm_2kHz_120Vac_36V = pd.read_excel(filename, sheet_name='36V 120VAC', skiprows=4, usecols="E:P", nrows= 12)
df_pwm_2kHz_230Vac_36V = pd.read_excel(filename, sheet_name='36V 230Vac', skiprows=4, usecols="E:P", nrows= 12)
df_pwm_2kHz_277Vac_36V = pd.read_excel(filename, sheet_name='36V 277Vac', skiprows=4, usecols="E:P", nrows= 12)

df_pwm_2kHz_120Vac_24V = pd.read_excel(filename, sheet_name='24V 120Vac', skiprows=4, usecols="E:P", nrows= 12)
df_pwm_2kHz_230Vac_24V = pd.read_excel(filename, sheet_name='24V 230Vac', skiprows=4, usecols="E:P", nrows= 12)
df_pwm_2kHz_277Vac_24V = pd.read_excel(filename, sheet_name='24V 277Vac', skiprows=4, usecols="E:P", nrows= 12)

df_pwm_400Hz_230Vac_48V = pd.read_excel(filename, sheet_name='48V 230Vac 400Hz', skiprows=4, usecols="E:P", nrows= 12)

"""DALI DIMMING"""

file = f"Unit {unit} DALI.xlsm"
path = "/DER 925 Reference Data/DALI Dimming/"
filename = os.getcwd() + path + file

df_dali_230Vac_48V = pd.read_excel(filename, sheet_name='48V 230Vac', skiprows=4, usecols="E:P", nrows= 86)


"""I2C DIMMING (Linear)"""

file = f"Unit {unit} (Linear).xlsm"
path = "/DER 925 Reference Data/I2C Dimming/Linear/"
filename = os.getcwd() + path + file

df_i2c_linear_120Vac_48V = pd.read_excel(filename, sheet_name='48V 120Vac', skiprows=4, usecols="E:P", nrows= 86)
df_i2c_linear_277Vac_48V = pd.read_excel(filename, sheet_name='48V 277Vac', skiprows=4, usecols="E:P", nrows= 86)

df_i2c_linear_120Vac_24V = pd.read_excel(filename, sheet_name='24V 120Vac', skiprows=4, usecols="E:P", nrows= 86)
df_i2c_linear_277Vac_24V = pd.read_excel(filename, sheet_name='24V 277Vac', skiprows=4, usecols="E:P", nrows= 86)


"""I2C DIMMING (Logarithmic)"""

file = f"Unit {unit} (Logarithmic).xlsm"
path = "/DER 925 Reference Data/I2C Dimming/Logarithmic/"
filename = os.getcwd() + path + file

df_i2c_logarithmic_120Vac_48V = pd.read_excel(filename, sheet_name='48V 120Vac', skiprows=4, usecols="E:P", nrows= 86)
df_i2c_logarithmic_277Vac_48V = pd.read_excel(filename, sheet_name='48V 277Vac', skiprows=4, usecols="E:P", nrows= 86)

df_i2c_logarithmic_120Vac_24V = pd.read_excel(filename, sheet_name='24V 120Vac', skiprows=4, usecols="E:P", nrows= 86)
df_i2c_logarithmic_277Vac_24V = pd.read_excel(filename, sheet_name='24V 277Vac', skiprows=4, usecols="E:P", nrows= 86)


"""RSET DIMMING"""

file = f"Unit {unit} (RSET).xlsm"
path = "/DER 925 Reference Data/LED Set/"
filename = os.getcwd() + path + file

df_rset_120Vac_48V = pd.read_excel(filename, sheet_name='48V 120Vac', skiprows=4, usecols="E:P", nrows= 10)
df_rset_277Vac_48V = pd.read_excel(filename, sheet_name='48V 277Vac', skiprows=4, usecols="E:P", nrows= 10)

df_rset_120Vac_24V = pd.read_excel(filename, sheet_name='24V 120Vac', skiprows=4, usecols="E:P", nrows= 10)
df_rset_277Vac_24V = pd.read_excel(filename, sheet_name='24V 277Vac', skiprows=4, usecols="E:P", nrows= 10)


# input()













"""WRITING DATA TO EXCEL"""





wb = load_workbook(dest)
 
"""PF, THD, EFF"""
sheet_name = 'PF, THD, EFF'
wb[sheet_name]["A3"] = f"DER-925 Unit {unit}"

# LINE REGULATION 48V
df = df_linereg_48V
start=12
end=23
for i in range(start,end+1):
    wb[sheet_name][f"B{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"C{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"D{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"E{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"F{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"G{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"H{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"I{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"J{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"K{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"L{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"M{i}"] = df.iloc[i-start,11]


# LINE REGULATION 36V
df = df_linereg_36V
start=28
end=39
for i in range(start,end+1):
    wb[sheet_name][f"B{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"C{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"D{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"E{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"F{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"G{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"H{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"I{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"J{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"K{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"L{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"M{i}"] = df.iloc[i-start,11]


# LINE REGULATION 24V
df = df_linereg_24V
start=44
end=55
for i in range(start,end+1):
    wb[sheet_name][f"B{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"C{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"D{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"E{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"F{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"G{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"H{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"I{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"J{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"K{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"L{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"M{i}"] = df.iloc[i-start,11]

"""NO LOAD"""
sheet_name = 'No Load Input Power'

# NO LOAD INPUT POWER

df = df_nl
start=10
end=21
for i in range(start,end+1):
    wb[sheet_name][f"B{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"C{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"D{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"E{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"F{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"G{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"H{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"I{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"J{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"K{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"L{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"M{i}"] = df.iloc[i-start,11]

# LIGHT OFF POWER
df = df_nl_lightoff
start=28
end=39
for i in range(start,end+1):
    wb[sheet_name][f"B{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"C{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"D{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"E{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"F{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"G{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"H{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"I{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"J{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"K{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"L{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"M{i}"] = df.iloc[i-start,11]


""""HARMONICS"""
sheet_name = 'Harmonics'

# 48V

df = df_harmonics_120Vac_48V
start=11
end=36
for i in range(start,end+1):
    wb[sheet_name][f"B{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"C{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"D{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"E{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"F{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"G{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"H{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"I{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"J{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"K{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"L{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"M{i}"] = df.iloc[i-start,11]

df = df_harmonics_230Vac_48V
start=42
end=67
for i in range(start,end+1):
    wb[sheet_name][f"B{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"C{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"D{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"E{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"F{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"G{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"H{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"I{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"J{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"K{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"L{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"M{i}"] = df.iloc[i-start,11]

df = df_harmonics_277Vac_48V
start=73
end=98
for i in range(start,end+1):
    wb[sheet_name][f"B{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"C{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"D{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"E{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"F{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"G{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"H{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"I{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"J{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"K{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"L{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"M{i}"] = df.iloc[i-start,11]

# 36V

df = df_harmonics_120Vac_36V
start=11
end=36
for i in range(start,end+1):
    wb[sheet_name][f"AA{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"AB{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"AC{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"AD{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"AE{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"AF{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"AG{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"AH{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"AI{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"AJ{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"AK{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"AL{i}"] = df.iloc[i-start,11]

df = df_harmonics_230Vac_36V
start=42
end=67
for i in range(start,end+1):
    wb[sheet_name][f"AA{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"AB{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"AC{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"AD{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"AE{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"AF{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"AG{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"AH{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"AI{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"AJ{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"AK{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"AL{i}"] = df.iloc[i-start,11]

df = df_harmonics_277Vac_36V
start=73
end=98
for i in range(start,end+1):
    wb[sheet_name][f"AA{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"AB{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"AC{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"AD{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"AE{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"AF{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"AG{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"AH{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"AI{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"AJ{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"AK{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"AL{i}"] = df.iloc[i-start,11]

# 24V

df = df_harmonics_120Vac_24V
start=11
end=36
for i in range(start,end+1):
    wb[sheet_name][f"AZ{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"BA{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"BB{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"BC{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"BD{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"BE{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"BF{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"BG{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"BH{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"BI{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"BJ{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"BK{i}"] = df.iloc[i-start,11]

df = df_harmonics_230Vac_24V
start=42
end=67
for i in range(start,end+1):
    wb[sheet_name][f"AZ{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"BA{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"BB{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"BC{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"BD{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"BE{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"BF{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"BG{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"BH{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"BI{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"BJ{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"BK{i}"] = df.iloc[i-start,11]

df = df_harmonics_277Vac_24V
start=73
end=98
for i in range(start,end+1):
    wb[sheet_name][f"AZ{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"BA{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"BB{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"BC{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"BD{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"BE{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"BF{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"BG{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"BH{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"BI{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"BJ{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"BK{i}"] = df.iloc[i-start,11]















"""ANALOG DIMMING"""
sheet_name = 'Analog Dimming'

# 48V

df = df_analog_120Vac_48V
start=12
end=22
for i in range(start,end+1):
    wb[sheet_name][f"C{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"D{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"E{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"F{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"G{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"H{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"I{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"J{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"K{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"L{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"M{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"N{i}"] = df.iloc[i-start,11]

df = df_analog_230Vac_48V
start=12
end=22
for i in range(start,end+1):
    wb[sheet_name][f"R{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"S{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"T{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"U{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"V{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"W{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"X{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"Y{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"Z{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"AA{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"AB{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"AC{i}"] = df.iloc[i-start,11]

df = df_analog_277Vac_48V
start=12
end=22
for i in range(start,end+1):
    wb[sheet_name][f"AG{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"AH{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"AI{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"AJ{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"AK{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"AL{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"AM{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"AN{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"AO{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"AP{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"AQ{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"AR{i}"] = df.iloc[i-start,11]

# 36V

df = df_analog_120Vac_36V
start=12
end=22
for i in range(start,end+1):
    wb[sheet_name][f"AW{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"AX{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"AY{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"AZ{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"BA{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"BB{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"BC{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"BD{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"BE{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"BF{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"BG{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"BH{i}"] = df.iloc[i-start,11]

df = df_analog_230Vac_36V
start=12
end=22
for i in range(start,end+1):
    wb[sheet_name][f"BL{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"BM{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"BN{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"BO{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"BP{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"BQ{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"BR{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"BS{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"BT{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"BU{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"BV{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"BW{i}"] = df.iloc[i-start,11]

df = df_analog_277Vac_36V
start=12
end=22
for i in range(start,end+1):
    wb[sheet_name][f"CA{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"CB{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"CC{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"CD{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"CE{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"CF{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"CG{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"CH{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"CI{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"CJ{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"CK{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"CL{i}"] = df.iloc[i-start,11]


# 24V

df = df_analog_120Vac_24V
start=12
end=22
for i in range(start,end+1):
    wb[sheet_name][f"CQ{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"CR{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"CS{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"CT{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"CU{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"CV{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"CW{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"CX{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"CY{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"CZ{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"DA{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"DB{i}"] = df.iloc[i-start,11]

df = df_analog_230Vac_24V
start=12
end=22
for i in range(start,end+1):
    wb[sheet_name][f"DF{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"DG{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"DH{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"DI{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"DJ{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"DK{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"DL{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"DM{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"DN{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"DO{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"DP{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"DQ{i}"] = df.iloc[i-start,11]

df = df_analog_277Vac_24V
start=12
end=22
for i in range(start,end+1):
    wb[sheet_name][f"DU{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"DV{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"DW{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"DX{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"DY{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"DZ{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"EA{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"EB{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"EC{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"ED{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"EE{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"EF{i}"] = df.iloc[i-start,11]
















"""RESISTOR DIMMING"""
sheet_name = 'Resistor Dimming'

# 48V

df = df_resistor_120Vac_48V
start=12
end=22
for i in range(start,end+1):
    wb[sheet_name][f"C{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"D{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"E{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"F{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"G{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"H{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"I{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"J{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"K{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"L{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"M{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"N{i}"] = df.iloc[i-start,11]

df = df_resistor_230Vac_48V
start=12
end=22
for i in range(start,end+1):
    wb[sheet_name][f"R{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"S{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"T{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"U{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"V{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"W{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"X{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"Y{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"Z{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"AA{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"AB{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"AC{i}"] = df.iloc[i-start,11]

df = df_resistor_277Vac_48V
start=12
end=22
for i in range(start,end+1):
    wb[sheet_name][f"AG{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"AH{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"AI{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"AJ{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"AK{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"AL{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"AM{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"AN{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"AO{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"AP{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"AQ{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"AR{i}"] = df.iloc[i-start,11]

# 36V

df = df_resistor_120Vac_36V
start=12
end=22
for i in range(start,end+1):
    wb[sheet_name][f"AW{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"AX{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"AY{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"AZ{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"BA{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"BB{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"BC{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"BD{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"BE{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"BF{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"BG{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"BH{i}"] = df.iloc[i-start,11]

df = df_resistor_230Vac_36V
start=12
end=22
for i in range(start,end+1):
    wb[sheet_name][f"BL{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"BM{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"BN{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"BO{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"BP{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"BQ{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"BR{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"BS{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"BT{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"BU{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"BV{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"BW{i}"] = df.iloc[i-start,11]

df = df_resistor_277Vac_36V
start=12
end=22
for i in range(start,end+1):
    wb[sheet_name][f"CA{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"CB{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"CC{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"CD{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"CE{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"CF{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"CG{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"CH{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"CI{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"CJ{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"CK{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"CL{i}"] = df.iloc[i-start,11]


# 24V

df = df_resistor_120Vac_24V
start=12
end=22
for i in range(start,end+1):
    wb[sheet_name][f"CQ{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"CR{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"CS{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"CT{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"CU{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"CV{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"CW{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"CX{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"CY{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"CZ{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"DA{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"DB{i}"] = df.iloc[i-start,11]

df = df_resistor_230Vac_24V
start=12
end=22
for i in range(start,end+1):
    wb[sheet_name][f"DF{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"DG{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"DH{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"DI{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"DJ{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"DK{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"DL{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"DM{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"DN{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"DO{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"DP{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"DQ{i}"] = df.iloc[i-start,11]

df = df_resistor_277Vac_24V
start=12
end=22
for i in range(start,end+1):
    wb[sheet_name][f"DU{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"DV{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"DW{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"DX{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"DY{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"DZ{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"EA{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"EB{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"EC{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"ED{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"EE{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"EF{i}"] = df.iloc[i-start,11]

















"""PWM Dimming (400 Hz)"""
sheet_name = 'PWM Dimming 400 Hz'

df = df_pwm_400Hz_230Vac_48V
start=12
end=22

for i in range(start,end+1):
    wb[sheet_name][f"R{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"S{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"T{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"U{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"V{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"W{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"X{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"Y{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"Z{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"AA{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"AB{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"AC{i}"] = df.iloc[i-start,11]













"""PWM Dimming (2000 Hz)"""
sheet_name = 'PWM Dimming 2000 Hz'

# 48V

df = df_pwm_2kHz_120Vac_48V
start=12
end=22
for i in range(start,end+1):
    wb[sheet_name][f"C{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"D{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"E{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"F{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"G{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"H{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"I{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"J{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"K{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"L{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"M{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"N{i}"] = df.iloc[i-start,11]

df = df_pwm_2kHz_230Vac_48V
start=12
end=22
for i in range(start,end+1):
    wb[sheet_name][f"R{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"S{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"T{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"U{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"V{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"W{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"X{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"Y{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"Z{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"AA{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"AB{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"AC{i}"] = df.iloc[i-start,11]

df = df_pwm_2kHz_277Vac_48V
start=12
end=22
for i in range(start,end+1):
    wb[sheet_name][f"AG{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"AH{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"AI{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"AJ{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"AK{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"AL{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"AM{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"AN{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"AO{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"AP{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"AQ{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"AR{i}"] = df.iloc[i-start,11]

# 36V

df = df_pwm_2kHz_120Vac_36V
start=12
end=22
for i in range(start,end+1):
    wb[sheet_name][f"AW{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"AX{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"AY{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"AZ{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"BA{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"BB{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"BC{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"BD{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"BE{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"BF{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"BG{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"BH{i}"] = df.iloc[i-start,11]

df = df_pwm_2kHz_230Vac_36V
start=12
end=22
for i in range(start,end+1):
    wb[sheet_name][f"BL{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"BM{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"BN{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"BO{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"BP{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"BQ{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"BR{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"BS{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"BT{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"BU{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"BV{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"BW{i}"] = df.iloc[i-start,11]

df = df_pwm_2kHz_277Vac_36V
start=12
end=22
for i in range(start,end+1):
    wb[sheet_name][f"CA{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"CB{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"CC{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"CD{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"CE{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"CF{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"CG{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"CH{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"CI{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"CJ{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"CK{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"CL{i}"] = df.iloc[i-start,11]


# 24V

df = df_pwm_2kHz_120Vac_24V
start=12
end=22
for i in range(start,end+1):
    wb[sheet_name][f"CQ{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"CR{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"CS{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"CT{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"CU{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"CV{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"CW{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"CX{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"CY{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"CZ{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"DA{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"DB{i}"] = df.iloc[i-start,11]

df = df_pwm_2kHz_230Vac_24V
start=12
end=22
for i in range(start,end+1):
    wb[sheet_name][f"DF{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"DG{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"DH{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"DI{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"DJ{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"DK{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"DL{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"DM{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"DN{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"DO{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"DP{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"DQ{i}"] = df.iloc[i-start,11]

df = df_pwm_2kHz_277Vac_24V
start=12
end=22
for i in range(start,end+1):
    wb[sheet_name][f"DU{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"DV{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"DW{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"DX{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"DY{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"DZ{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"EA{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"EB{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"EC{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"ED{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"EE{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"EF{i}"] = df.iloc[i-start,11]








"""DALI I2C Dimming (Logarithmic)"""
sheet_name = 'DALI I2C Dimming (Logarithmic)'


df = df_dali_230Vac_48V
start=10
end=95
for i in range(start,end+1):
    wb[sheet_name][f"Q{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"R{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"S{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"T{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"U{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"V{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"W{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"X{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"Y{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"Z{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"AA{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"AB{i}"] = df.iloc[i-start,11]








"""I2C Dimming (Linear)"""
sheet_name = 'I2C Dimming (Linear)'

df = df_i2c_linear_120Vac_48V
start=10
end=95
for i in range(start,end+1):
    wb[sheet_name][f"B{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"C{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"D{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"E{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"F{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"G{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"H{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"I{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"J{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"K{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"L{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"M{i}"] = df.iloc[i-start,11]

df = df_i2c_linear_277Vac_48V
start=10
end=95
for i in range(start,end+1):
    wb[sheet_name][f"AF{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"AG{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"AH{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"AI{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"AJ{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"AK{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"AL{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"AM{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"AN{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"AO{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"AP{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"AQ{i}"] = df.iloc[i-start,11]

df = df_i2c_linear_120Vac_24V
start=10
end=95
for i in range(start,end+1):
    wb[sheet_name][f"AV{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"AW{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"AX{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"AY{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"AZ{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"BA{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"BB{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"BC{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"BD{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"BE{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"BF{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"BG{i}"] = df.iloc[i-start,11]

df = df_i2c_linear_277Vac_24V
start=10
end=95
for i in range(start,end+1):
    wb[sheet_name][f"BZ{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"CA{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"CB{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"CC{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"CD{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"CE{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"CF{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"CG{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"CH{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"CI{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"CJ{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"CK{i}"] = df.iloc[i-start,11]






"""I2C Dimming (Logarithmic)"""
sheet_name = 'I2C Dimming (Logarithmic)'

df = df_i2c_logarithmic_120Vac_48V
start=10
end=95
for i in range(start,end+1):
    wb[sheet_name][f"B{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"C{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"D{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"E{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"F{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"G{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"H{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"I{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"J{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"K{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"L{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"M{i}"] = df.iloc[i-start,11]

df = df_i2c_logarithmic_277Vac_48V
start=10
end=95
for i in range(start,end+1):
    wb[sheet_name][f"AF{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"AG{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"AH{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"AI{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"AJ{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"AK{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"AL{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"AM{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"AN{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"AO{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"AP{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"AQ{i}"] = df.iloc[i-start,11]

df = df_i2c_logarithmic_120Vac_24V
start=10
end=95
for i in range(start,end+1):
    wb[sheet_name][f"AV{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"AW{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"AX{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"AY{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"AZ{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"BA{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"BB{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"BC{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"BD{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"BE{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"BF{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"BG{i}"] = df.iloc[i-start,11]

df = df_i2c_logarithmic_277Vac_24V
start=10
end=95
for i in range(start,end+1):
    wb[sheet_name][f"BZ{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"CA{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"CB{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"CC{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"CD{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"CE{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"CF{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"CG{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"CH{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"CI{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"CJ{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"CK{i}"] = df.iloc[i-start,11]










"""RSET Dimming"""
sheet_name = 'RSET Dimming'

df = df_rset_120Vac_48V
start=10
end=19
for i in range(start,end+1):
    wb[sheet_name][f"B{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"C{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"D{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"E{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"F{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"G{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"H{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"I{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"J{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"K{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"L{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"M{i}"] = df.iloc[i-start,11]

df = df_rset_277Vac_48V
start=10
end=19
for i in range(start,end+1):
    wb[sheet_name][f"AF{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"AG{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"AH{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"AI{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"AJ{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"AK{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"AL{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"AM{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"AN{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"AO{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"AP{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"AQ{i}"] = df.iloc[i-start,11]

df = df_rset_120Vac_24V
start=10
end=19
for i in range(start,end+1):
    wb[sheet_name][f"AV{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"AW{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"AX{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"AY{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"AZ{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"BA{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"BB{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"BC{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"BD{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"BE{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"BF{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"BG{i}"] = df.iloc[i-start,11]

df = df_rset_277Vac_24V
start=10
end=19
for i in range(start,end+1):
    wb[sheet_name][f"BZ{i}"] = df.iloc[i-start,0]
    wb[sheet_name][f"CA{i}"] = df.iloc[i-start,1]
    wb[sheet_name][f"CB{i}"] = df.iloc[i-start,2]
    wb[sheet_name][f"CC{i}"] = df.iloc[i-start,3]
    wb[sheet_name][f"CD{i}"] = df.iloc[i-start,4]
    wb[sheet_name][f"CE{i}"] = df.iloc[i-start,5]
    wb[sheet_name][f"CF{i}"] = df.iloc[i-start,6]
    wb[sheet_name][f"CG{i}"] = df.iloc[i-start,7]
    wb[sheet_name][f"CH{i}"] = df.iloc[i-start,8]
    wb[sheet_name][f"CI{i}"] = df.iloc[i-start,9]
    wb[sheet_name][f"CJ{i}"] = df.iloc[i-start,10]
    wb[sheet_name][f"CK{i}"] = df.iloc[i-start,11]





wb.save(dest)














print(f"Unit {unit} Data Pack is complete.")






# df = df_harmonics_120Vac_24V
# start=11
# end=36
# for i in range(start,end+1):
#     wb[sheet_name][f"{i}"] = df.iloc[i-start,0]
#     wb[sheet_name][f"{i}"] = df.iloc[i-start,1]
#     wb[sheet_name][f"{i}"] = df.iloc[i-start,2]
#     wb[sheet_name][f"{i}"] = df.iloc[i-start,3]
#     wb[sheet_name][f"{i}"] = df.iloc[i-start,4]
#     wb[sheet_name][f"{i}"] = df.iloc[i-start,5]
#     wb[sheet_name][f"{i}"] = df.iloc[i-start,6]
#     wb[sheet_name][f"{i}"] = df.iloc[i-start,7]
#     wb[sheet_name][f"{i}"] = df.iloc[i-start,8]
#     wb[sheet_name][f"{i}"] = df.iloc[i-start,9]
#     wb[sheet_name][f"{i}"] = df.iloc[i-start,10]
#     wb[sheet_name][f"{i}"] = df.iloc[i-start,11]