# USER INPUT STARTS HERE
#########################################################################################
## TEST PARAMETERS

# comms
ac_source_address = 5
source_power_meter_address = 1 
load_power_meter_address = 2
eload_address = 8
scope_address = "10.125.10.112"

# scope settings
vin_channel = 3
vout_channel = 1
iout_channel = 2
position = -2

# trigger settings
trigger_level = 30   # Vrms
trigger_source = vin_channel
trigger_slope = 'POS'

# eload settings
eload_channel = 1

## INPUT
vin = [90, 115, 230, 265]
freq = [60, 60, 50, 50]

## OUTPUT
vout = 20
Iout_max = 3.25 # A
Iout = [Iout_max, 0.50*Iout_max]
Iout_name = [100, 50]
#########################################################################################
# USER INPUT ENDS HERE

# from powi.equipment import ACSource, PowerMeter, ElectronicLoad, Oscilloscope
from powi.equipment import Oscilloscope
from powi.equipment import headers, create_folder, footers, waveform_counter
from time import sleep, time
import os

## initialize equipment
# ac = ACSource(ac_source_address)
# pms = PowerMeter(source_power_meter_address)
# pml = PowerMeter(load_power_meter_address)
# eload = ElectronicLoad(eload_address)
scope = Oscilloscope(scope_address)

def reset():
    ac.turn_off()
    eload.channel[1].cc = 1
    eload.channel[1].turn_on()
    eload.channel[2].cc = 1
    eload.channel[2].turn_on()
    sleep(1)

#### special functions #####

def startup_90degPhase(voltage,frequency):
    # ac.write("OUTP:STAT?")
    # ac.write("SYST:ERR?")
    ac.voltage = 0
    ac.frequency = frequency
    # ac.write("OUTP ON")
    ac.turn_on()
    ac.write("TRIG:TRAN:SOUR BUS")
    ac.write("ABORT")
    ac.write("LIST:DWEL 1, 1, 1")

    ac.write("VOLT:MODE LIST")
    ac.write(f"LIST:VOLT {voltage}, {voltage}, {voltage}")
    ac.write("VOLT:SLEW:MODE LIST")
    ac.write("LIST:VOLT:SLEW 9.9e+037, 9.9e+037, 9.9e+037")

    ac.write("FREQ:MODE LIST")
    ac.write(f"LIST:FREQ {frequency}, {frequency}, {frequency}")
    ac.write("FREQ:SLEW:MODE LIST")
    ac.write("LIST:FREQ:SLEW 9.9e+037, 9.9e+037, 9.9e+037")

    ac.write("VOLT:OFFS:MODE FIX")
    ac.write("VOLT:OFFS:SLEW:MODE FIX")

    ac.write("PHAS:MODE LIST")
    # ac.write("LIST:PHAS 270, 270, 270")
    ac.write("LIST:PHAS 90, 90, 90")

    ac.write("CURR:PEAK:MODE LIST")
    ac.write("LIST:CURR 40.4, 40.4, 40.4")

    ac.write("FUNC:MODE FIX")
    ac.write("LIST:TTLT ON,OFF,OFF")
    ac.write("LIST:STEP AUTO")
    # ac.write("SYST:ERR?")

    ac.write("OUTP:TTLT:STAT ON")
    ac.write("OUTP:TTLT:SOUR LIST")
    ac.write("TRIG:SYNC:SOUR PHASE")
    ac.write("TRIG:SYNC:PHAS 0.0")
    ac.write("TRIG:TRAN:DEL 0")
    ac.write("Sens:Swe:Offs:Poin 0")
    ac.write("TRIG:ACQ:SOUR TTLT")
    ac.write("INIT:IMM:SEQ3")
    ac.write("LIST:COUN 1")
    ac.write("INIT:IMM:SEQ1")
    ac.write("TRIG:TRAN:SOUR BUS")
    # ac.write("SYST:ERR?")

    ac.write("TRIG:IMM")

def startup(case="cc"):

    global waveform_counter
    global Iout
    
    Iout_index = 0

    scope.init_trigger(trigger_source, trigger_level, trigger_slope)

    if case == "cc":
        print('{:<30} {:>1}'.format("[LOAD CC]", "startup_time"))
    if case == "cr":
        print('{:<30} {:>1}'.format("[LOAD CR]", "startup_time"))
    if case == "nl":
        print('{:<30} {:>1}'.format("[NO LOAD]", "startup_time"))

    for voltage, frequency in zip(vin, freq):

        ## set input voltage
        ac.voltage = voltage
        ac.frequency = frequency

        if case == "nl":
            Iout = [0]

        for curr in Iout:
            if case == "cc" and curr != 0:
                eload.channel[eload_channel].cc = curr
                eload.channel[eload_channel].turn_on()     
                filename = f'{voltage}Vac {Iout_name[Iout_index]}LoadCC.png'
            elif case == "cr" and curr != 0:
                rload = vout/curr
                rload = f"{rload:.4f}"
                rload = float(rload)
                eload.channel[1].cr = rload
                eload.channel[eload_channel].turn_on()
                filename = f'{voltage}Vac {Iout_name[Iout_index]}LoadCR.png'
            elif case == "nl":
                eload.channel[eload_channel].turn_off()
                filename = f'{voltage}Vac 0Load.png'
            else:
                print("Please enter type of load (cc/cr/nl).")
                break

            # trigger scope
            scope.run_single()
            sleep(2)
            startup_90degPhase(voltage, frequency)
            sleep(5)
            scope.stop()

            # get waveform data from scope  
            vo_data = scope.get_chan_data(vout_channel)

            # search algorithm for cursor 1
            pos_x1 = 0 # place at the start of the input voltage
            
            # search algorithm for cursor 2
            j = 0
            pos_x2 = 0
            for point in vo_data:
                if point >= vout: # if vo >= vo_reg
                    pos_x2 = j               # set cursor there
                    break
                j += 1

            # set cursors (to get startup time)
            a = scope.get_horizontal()
            resolution = float(a["resolution"])
            minimum = float(a["scale"])*(position) # set the cursor to the leftmost part of the screen
            
            if pos_x1 == 0: cursor1 = 0
            else: cursor1 = minimum + resolution*pos_x1 
            if pos_x2 == 0: cursor2 = 0
            else: cursor2 = minimum + resolution*pos_x2

            scope.cursor(channel=vin_channel, cursor_set=1, X1=cursor1, X2=cursor2)
            sleep(0.5)
            startup_time = scope.get_cursor()["delta x"]
            print(f"startup time = {startup_time} s")

            # get screenshot
            sleep(1)
            scope.get_screenshot(filename, waveforms_folder)
            
            Iout_index += 1
            waveform_counter += 1
            print(filename)
            startup_time = "0.143 s"
            print('{:<30} {:>5}'.format(filename, startup_time))
            reset()

        Iout_index = 0 # resest iout naming index for the next voltage
    
    print()

## main code ##
# headers("Output Startup")
# startup("cc")
# startup("cr")
# startup("nl")
# footers(waveform_counter)






def get_chan_data(channel=1):
    # scope.write('FORM ASC')
    scope.write("FORM:DATA INT,16")
    scope.write('EXP:WAV:INCX OFF')
    temp = scope.write(f'CHAN{channel}:WAV1:DATA?') # type: string
    # temp = list(temp.split(",")) # type: list
    # data = []
    # for h in temp:
    #     data.append(float(h))
    # return data

    with open("data.txt", "w") as file:
        file.write(temp)
    return temp


def save_channel_data(channel=1):
    scope.write("EXPort:WAVeform:MULTIchannel ON")
    scope.write("EXP:WAV:INCX OFF")
    for i in range(1, 5):
        if i == channel:
            scope.write(f"CHANnel{i}:EXPortstate ON")
        else:
            scope.write(f"CHANnel{i}:EXPortstate OFF")
    scope.write("FORM:DATA INT,16")
    scope.write(f"CHAN{channel}:WAV1:DATA?")



scope.stop()
sleep(1)
a = get_chan_data(iout_channel)
print(len(a))

