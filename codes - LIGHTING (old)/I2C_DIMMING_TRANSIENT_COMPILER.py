print("Charles Cayno | 28-Jun-2021")

# import sys
# # user input #################################################################
# pwm_frequency = sys.argv[1] # 2000, 400 
# ##############################################################################


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
    return result 

# Converts string-type excel address to int-type column and row
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


col, row = col_row_converter("B1")
print(col,row)
print(type(col))





# class BlockPrinter():

#     def writer(self, anchor, condition):
#         self.col, self.row = col_row_converter(anchor)
#         return self.col, self.row


# FirstPowerUp = BlockPrinter()

# x,y = FirstPowerUp.writer("B1")

# print(x,y)


def printer():

    image = f"{condition}, {start}-{end} bit, I2C Dimming, {test_condition}.png"
    image_path = os.getcwd() + f"waveforms/I2C Dimming {load}V {vin}Vac/" + image

    

def print_I2C_img_to_excel(anchor="B2", condition="46V, 120Vac 60Hz", filepath="waveforms/"):

    I2C_COMMAND_LIST = [0, 50, 100, 150, 200, 250, 254]

    for i in range(len(I2C_COMMAND_LIST)):
        for cmdbit in I2C_COMMAND_LIST:
            if I2C_COMMAND_LIST[i] < cmdbit:

                printer()




def print_img_to_excel(anchor="B2", end=100, cond="46V, 120Vac 60Hz", pwm_frequency=2000, filepath="waveforms/"):

    global missingCounter

    col, row = col_row_converter(anchor)
    
    for i in range(0,end,10):
        
        image = f"{cond}, {i}-{end}% PWM Dimming, f={pwm_frequency}Hz, First Power Up (Discharged Cout).png"
        img_path = filepath + image

        row = int(row)
        col = int(col)

        i = int(i/10)
        label_cell_row = i*17+row-1
        label_cell_col = col - 1
        image_cell_row = i*17+row+1
        
        worksheet.write(label_cell_row, label_cell_col, image)
        
        if os.path.exists(img_path):
            worksheet.insert_image(f'{colnum_string(col)}{image_cell_row}', img_path, {'x_scale': x_scale, 'y_scale': y_scale})
            # print(image)
        else:
            print(f"{image} is still missing.")
            missingCounter+=1
    
    col, row = col_row_converter(anchor)
    col = col + 9
    for i in range(0,end,10):
        image = f"{cond}, {i}-{end}% PWM Dimming, f={pwm_frequency}Hz, No AC reset (Not Discharged Cout).png"
        img_path = filepath + image

        row = int(row)
        col = int(col)
        i = int(i/10)
        label_cell_row = i*17+row-1
        label_cell_col = col - 1
        image_cell_row = i*17+row+1
        
        worksheet.write(label_cell_row, label_cell_col, image)
        
        if os.path.exists(img_path):
            worksheet.insert_image(f'{colnum_string(col)}{image_cell_row}', img_path, {'x_scale': x_scale, 'y_scale': y_scale})
            # print(image)
        else:
            print(f"{image} is still missing.")
            missingCounter+=1
        
        
    col, row = col_row_converter(anchor)
    col = col + 18
    for i in range(0,end,10):
        image = f"{cond}, {end}-{i}% PWM Dimming, f={pwm_frequency}Hz.png"
        img_path = filepath + image

        row = int(row)
        col = int(col)
        i = int(i/10)
        label_cell_row = i*17+row-1
        label_cell_col = col - 1
        image_cell_row = i*17+row+1
        
        worksheet.write(label_cell_row, label_cell_col, image)
        
        if os.path.exists(img_path):
            worksheet.insert_image(f'{colnum_string(col)}{image_cell_row}', img_path, {'x_scale': x_scale, 'y_scale': y_scale})
            # print(image)
        else:
            print(f"{image} is still missing.")
            missingCounter+=1

# def transfer_to_excel(condition='PWM Dimming (f=2000Hz) 24V 120Vac', condition_1='46V, 120Vac 60Hz', pwm_frequency=2000):
#     path = f"/waveforms/PWM Dimming/{condition}/"
#     filepath = os.getcwd() + path
#     print_img_to_excel(anchor="B2", end=100, cond=condition_1, pwm_frequency=pwm_frequency, filepath=filepath)
#     print_img_to_excel(anchor="B172", end=70, cond=condition_1, pwm_frequency=pwm_frequency, filepath=filepath)
#     print_img_to_excel(anchor="B291", end=50, cond=condition_1, pwm_frequency=pwm_frequency, filepath=filepath)
#     print_img_to_excel(anchor="B376", end=40, cond=condition_1, pwm_frequency=pwm_frequency, filepath=filepath)
#     print_img_to_excel(anchor="B444", end=30, cond=condition_1, pwm_frequency=pwm_frequency, filepath=filepath)
#     print_img_to_excel(anchor="B495", end=20, cond=condition_1, pwm_frequency=pwm_frequency, filepath=filepath)
#     print_img_to_excel(anchor="B529", end=10, cond=condition_1, pwm_frequency=pwm_frequency, filepath=filepath)


# """MAIN"""

# import os
# import xlsxwriter


# workbook = xlsxwriter.Workbook(f'PWM Dimming Transient Waveforms f={pwm_frequency}Hz.xlsx')

# # image settings
# x_scale = 0.3
# y_scale = 0.3


# """46V"""
# condition = f"PWM Dimming (f={pwm_frequency}Hz) 46V 120Vac"
# condition_1 = "46V, 120Vac 60Hz"
# worksheet = workbook.add_worksheet(condition_1[0:11])
# transfer_to_excel(condition, condition_1, pwm_frequency)
# condition = f"PWM Dimming (f={pwm_frequency}Hz) 46V 230Vac"
# condition_1 = "46V, 230Vac 50Hz"
# worksheet = workbook.add_worksheet(condition_1[0:11])
# transfer_to_excel(condition, condition_1, pwm_frequency)
# condition = f"PWM Dimming (f={pwm_frequency}Hz) 46V 277Vac"
# condition_1 = "46V, 277Vac 60Hz"
# worksheet = workbook.add_worksheet(condition_1[0:11])
# transfer_to_excel(condition, condition_1, pwm_frequency)

# """36V"""
# condition = f"PWM Dimming (f={pwm_frequency}Hz) 36V 120Vac"
# condition_1 = "36V, 120Vac 60Hz"
# worksheet = workbook.add_worksheet(condition_1[0:11])
# transfer_to_excel(condition, condition_1, pwm_frequency)
# condition = f"PWM Dimming (f={pwm_frequency}Hz) 36V 230Vac"
# condition_1 = "36V, 230Vac 50Hz"
# worksheet = workbook.add_worksheet(condition_1[0:11])
# transfer_to_excel(condition, condition_1, pwm_frequency)
# condition = f"PWM Dimming (f={pwm_frequency}Hz) 36V 277Vac"
# condition_1 = "36V, 277Vac 60Hz"
# worksheet = workbook.add_worksheet(condition_1[0:11])
# transfer_to_excel(condition, condition_1, pwm_frequency)

# """24V"""
# condition = f"PWM Dimming (f={pwm_frequency}Hz) 24V 120Vac"
# condition_1 = "24V, 120Vac 60Hz"
# worksheet = workbook.add_worksheet(condition_1[0:11])
# transfer_to_excel(condition, condition_1, pwm_frequency)
# condition = f"PWM Dimming (f={pwm_frequency}Hz) 24V 230Vac"
# condition_1 = "24V, 230Vac 50Hz"
# worksheet = workbook.add_worksheet(condition_1[0:11])
# transfer_to_excel(condition, condition_1, pwm_frequency)
# condition = f"PWM Dimming (f={pwm_frequency}Hz) 24V 277Vac"
# condition_1 = "24V, 277Vac 60Hz"
# worksheet = workbook.add_worksheet(condition_1[0:11])
# transfer_to_excel(condition, condition_1, pwm_frequency)

# print()
# print(f'PWM Dimming Transient Waveforms (f={pwm_frequency}Hz)')

# print()
# print("On-going compilation...")
# print()

# workbook.close()

# if missingCounter == 0:
#     print(f'PWM Dimming Transient Waveforms f={pwm_frequency}Hz.xlsx -> compilation complete.')
#     print()
# else:
#     print(f"{missingCounter} waveforms are still missing.")
#     print(f"{(int)(missingCounter/96)} test conditions left.")
#     print()
#     print(f'PWM Dimming Transient Waveforms f={pwm_frequency}Hz.xlsx -> INITIAL compilation complete.')
