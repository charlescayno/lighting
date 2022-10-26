import pyvisa
import tkinter as tk

address = {"AC":"GPIB0::5::INSTR"}

class ACSource(object):
    def __init__(self):
        rm = pyvisa.ResourceManager()
        device_address = address["AC"]

        try:
            self.device = rm.open_resource(device_address)
            self.device.write("*IDN?")
        except Exception as e:
            print (e)
# Initialze
            
    def turn_on(self):       
        self.device.write("VOLT:OFFSET 0")
        self.device.write("VOLT 0")  
        self.device.write("OUTP ON")
        self.device.write("*CLS")   
        
    def turn_off(self):
        self.device.write("VOLT:OFFSET 0")
        self.device.write("VOLT 0")
        self.device.write("OUTP OFF")        
        self.device.write("*CLS")

# AC Source
        
    def set_volt(self, ac_voltage):        
        self.device.write("OUTP: COUP AC")
        self.device.write("VOLT:OFFSET 0")      
        self.device.write("VOLT %s" % (ac_voltage))
        self.device.write("*CLS")
            
# DC Source
        
    def set_offset(self, dc_offset):
        self.device.write("OUTP: COUP DC")        
        self.device.write("VOLT 0")      
        self.device.write("VOLTAGE:OFFSET %s" % (dc_offset))
        self.device.write("*CLS")
        
# Commands

def ac_source_reset():
    try:
        ac_source = ACSource()
        ac_source.turn_off()
    except Exception as e:
        print (e)

def ac_source_set_volt_50():
    try:   
        ac_source = ACSource()
        ac_voltage = 50
        set_volt(ac_source, ac_voltage)
    except Exception as e:
        print (e)
        

def ac_source_set_volt_100():
    try:
        ac_source = ACSource()
        ac_source.set_volt_100()
    except Exception as e:
        print (e)

def ac_source_set_volt_220():
    try:
        ac_source = ACSource()
        ac_source.set_volt_220()
    except Exception as e:
        print (e)

def ac_source_set_offset_310():
    try:
        ac_source = ACSource()
        ac_source.set_offset_310()
    except Exception as e:
        print (e)
        

def ac_source_set_offset_200():
    try:
        ac_source = ACSource()
        ac_source.set_offset_200()
    except Exception as e:
        print (e)

def ac_source_set_offset_100():
    try:
        ac_source = ACSource()
        ac_source.set_offset_100()
    except Exception as e:
        print (e)
        
def ac_source_set_offset_50():
    try:
        ac_source = ACSource()
        ac_source.set_offset_50()
    except Exception as e:
        print (e)
     
def ac_source_output_on():
    try:
        ac_source = ACSource()
        ac_source.turn_on()
    except Exception as e:
        print (e)

def ac_source_output_off():
    try:
        ac_source = ACSource()
        ac_source.turn_off()
    except Exception as e:
        print (e)

window = tk.Tk()
window.title("AC Source GUI")

label_output = tk.Label(text = "Output")
label_output.grid(row=0,column=0,sticky="w",padx=5,pady=5)

label_separator = tk.Label(text = "            ")
label_separator.grid(row=0,column=1)

button_output_on = tk.Button(text="ON",height=2,width=5,command=ac_source_output_on,fg="yellow",bg="green")
button_output_on.grid(row=0,column=3,pady=5)

button_output_off = tk.Button(text="OFF",height=2,width=5,command=ac_source_output_off,fg="yellow",bg="red")
button_output_off.grid(row=0,column=4,pady=5)

button_reset = tk.Button(text="CLEAR SETTINGS",height=2,width=33,command=ac_source_reset,fg="yellow",bg="green")
button_reset.grid(row=1,column=0,pady=5,columnspan=5)

label_presets = tk.Label(text = "DC Presets:                                                            ")
label_presets['bg'] = "blue"
label_presets['fg'] = "yellow"
label_presets.grid(row=2,column=0,sticky="w",padx=5,columnspan=5)

label_offset_50 = tk.Label(text = "50 VDC")
label_offset_50.grid(row=3,column=0,pady=5)

button_offset_50 = tk.Button(text="SET",height=2,width=5,command=ac_source_set_offset_50,fg="blue",bg="orange")
button_offset_50.grid(row=3,column=4,pady=5)

label_offset_100 = tk.Label(text = "100 VDC")
label_offset_100.grid(row=4,column=0,pady=5)

button_offset_100 = tk.Button(text="SET",height=2,width=5,command=ac_source_set_offset_100,fg="blue",bg="orange")
button_offset_100.grid(row=4,column=4,pady=5)

label_offset_200 = tk.Label(text = "200 VDC")
label_offset_200.grid(row=5,column=0,pady=5)

button_offset_200 = tk.Button(text="SET",height=2,width=5,command=ac_source_set_offset_200,fg="blue",bg="orange")
button_offset_200.grid(row=5,column=4,pady=5)

label_offset_310 = tk.Label(text = "310 VDC")
label_offset_310.grid(row=6,column=0,pady=5)

button_offset_310 = tk.Button(text="SET",height=2,width=5,command=ac_source_set_offset_310,fg="blue",bg="orange")
button_offset_310.grid(row=6,column=4,pady=5)

label_presets = tk.Label(text = "AC Presets:                                                           ")
label_presets['bg'] = "blue"
label_presets['fg'] = "yellow"
label_presets.grid(row=7,column=0,sticky="w",padx=5,columnspan=5)

label_volt_50 = tk.Label(text = "50 VAC")
label_volt_50.grid(row=8,column=0,pady=5)

button_volt_50 = tk.Button(text="SET",height=2,width=5,command=ac_source_set_volt_50,fg="purple",bg="yellow")
button_volt_50.grid(row=8,column=4,pady=5)

label_volt_100 = tk.Label(text = "100 VAC")
label_volt_100.grid(row=9,column=0,pady=5)

button_volt_100 = tk.Button(text="SET",height=2,width=5,command=ac_source_set_volt_100,fg="purple",bg="yellow")
button_volt_100.grid(row=9,column=4,pady=5)

label_volt_220 = tk.Label(text = "220 VAC")
label_volt_220.grid(row=10,column=0,pady=5)

button_volt_220 = tk.Button(text="SET",height=2,width=5,command=ac_source_set_volt_220,fg="purple",bg="yellow")
button_volt_220.grid(row=10,column=4,pady=5)

window.mainloop()
