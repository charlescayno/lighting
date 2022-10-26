print("Charles Cayno | 08-Jun-2021")

global missingCounter
missingCounter = 0

"""Getting Size of Image"""
# from PIL import Image  # uses pillow 
# image_file = "24V, 120Vac 60Hz, 0-7V Analog, First Power Up (Discharged Cout).png"
# im = Image.open(image_file)
# print(im.size)
# input()
# 1280, 800 - image size


"""CUSTOM FUNCTIONS"""

def titleToNumber(s):
    result = 0; 
    for B in range(len(s)): 
        result *= 26; 
        result += ord(s[B]) - ord('A') + 1; 
    return result; 

def col_row_converter(string="B1"):
    col = ''.join(char for char in string if char.isalpha())
    col = ord(col)-ord('A')+1
    row = ''.join(char for char in string if char.isdigit())
    return col, row

def get_col(string="B1"):
    col = ''.join(char for char in string if char.isalpha())
    return col

def colnum_string(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string

def print_img_to_excel(anchor="B2", end=10, cond="46V, 120Vac 60Hz", filepath="waveforms/"):

    global missingCounter

    col, row = col_row_converter(anchor)
    for i in range(0,end,1):
        image = f"{cond}, {i}-{end}V Analog, First Power Up (Discharged Cout).png"
        img_path = filepath + image

        row = int(row)
        col = int(col)
        label_cell_row = i*17+row-1
        label_cell_col = col - 1
        image_cell_row = i*17+row+1

        worksheet.write(label_cell_row, label_cell_col, image)

        if os.path.exists(img_path):
            worksheet.insert_image(f'{colnum_string(col)}{image_cell_row}', img_path, {'x_scale': x_scale, 'y_scale': y_scale})
        else:
            missingCounter+=1

    col, row = col_row_converter(anchor)
    col = col + 9
    for i in range(0,end,1):
        image = f"{cond}, {i}-{end}V Analog, No AC reset (Not Discharged Cout).png"
        img_path = filepath + image
        
        row = int(row)
        col = int(col)
        label_cell_row = i*17+row-1
        label_cell_col = col - 1
        image_cell_row = i*17+row+1
        
        worksheet.write(label_cell_row, label_cell_col, image)
        
        if os.path.exists(img_path):
            worksheet.insert_image(f'{colnum_string(col)}{image_cell_row}', img_path, {'x_scale': x_scale, 'y_scale': y_scale})
        else:
            missingCounter+=1
    col, row = col_row_converter(anchor)
    col = col + 18
    for i in range(0,end,1):
        image = f"{cond}, {end}-{i}V Analog.png"
        img_path = filepath + image
        
        row = int(row)
        col = int(col)
        label_cell_row = i*17+row-1
        label_cell_col = col - 1
        image_cell_row = i*17+row+1
        
        worksheet.write(label_cell_row, label_cell_col, image)
        
        if os.path.exists(img_path):
            worksheet.insert_image(f'{colnum_string(col)}{image_cell_row}', img_path, {'x_scale': x_scale, 'y_scale': y_scale})
        else:
            missingCounter+=1


def transfer_to_excel(condition='46V 120Vac', condition_1='46V, 120Vac 60Hz'):
    path = f"/waveforms/Analog Dimming/{condition}/"
    filepath = os.getcwd() + path
    print_img_to_excel(anchor="B2", end=10, cond=condition_1, filepath=filepath)
    print_img_to_excel(anchor="B172", end=7, cond=condition_1, filepath=filepath)
    print_img_to_excel(anchor="B291", end=5, cond=condition_1, filepath=filepath)
    print_img_to_excel(anchor="B376", end=4, cond=condition_1, filepath=filepath)
    print_img_to_excel(anchor="B444", end=3, cond=condition_1, filepath=filepath)
    print_img_to_excel(anchor="B495", end=2, cond=condition_1, filepath=filepath)
    print_img_to_excel(anchor="B529", end=1, cond=condition_1, filepath=filepath)

import os
import xlsxwriter

workbook = xlsxwriter.Workbook('Analog Dimming Transient Waveforms.xlsx')

# image settings
x_scale = 0.3
y_scale = 0.3

# user input

"""46V"""
condition = "46V 120Vac"
condition_1 = "46V, 120Vac 60Hz"
worksheet = workbook.add_worksheet(condition_1[0:11])
transfer_to_excel(condition, condition_1)
condition = "46V 230Vac"
condition_1 = "46V, 230Vac 50Hz"
worksheet = workbook.add_worksheet(condition_1[0:11])
transfer_to_excel(condition, condition_1)
condition = "46V 277Vac"
condition_1 = "46V, 277Vac 60Hz"
worksheet = workbook.add_worksheet(condition_1[0:11])
transfer_to_excel(condition, condition_1)

"""36V"""
condition = "36V 120Vac"
condition_1 = "36V, 120Vac 60Hz"
worksheet = workbook.add_worksheet(condition_1[0:11])
transfer_to_excel(condition, condition_1)
condition = "36V 230Vac"
condition_1 = "36V, 230Vac 60Hz"
worksheet = workbook.add_worksheet(condition_1[0:11])
transfer_to_excel(condition, condition_1)
condition = "36V 277Vac"
condition_1 = "36V, 277Vac 60Hz"
worksheet = workbook.add_worksheet(condition_1[0:11])
transfer_to_excel(condition, condition_1)

"""24V"""
condition = "24V 120Vac"
condition_1 = "24V, 120Vac 60Hz"
worksheet = workbook.add_worksheet(condition_1[0:11])
transfer_to_excel(condition, condition_1)
condition = "24V 230Vac"
condition_1 = "24V, 230Vac 50Hz"
worksheet = workbook.add_worksheet(condition_1[0:11])
transfer_to_excel(condition, condition_1)
condition = "24V 277Vac"
condition_1 = "24V, 277Vac 60Hz"
worksheet = workbook.add_worksheet(condition_1[0:11])
transfer_to_excel(condition, condition_1)


print()
print(f'Analog Dimming Transient Waveforms.xlsx')

print()
print("On-going compilation...")
print()

workbook.close()

if missingCounter == 0:
    print(f'Analog Dimming Transient Waveforms.xlsx -> compilation complete.')
    print()
else:
    print(f"{missingCounter} waveforms are still missing.")
    print(f"{(int)(missingCounter/96)} test conditions left.")
    print()
    print(f'Analog Dimming Transient Waveforms.xlsx -> INITIAL compilation complete.')