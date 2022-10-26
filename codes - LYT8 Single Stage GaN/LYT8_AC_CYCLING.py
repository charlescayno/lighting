from equipment_settings import *

ac = ACSource(EQUIPMENT_ADDRESS.AC_SOURCE)
pms = PowerMeter(EQUIPMENT_ADDRESS.POWER_METER_SOURCE)
pml = PowerMeter(EQUIPMENT_ADDRESS.POWER_METER_LOAD)
eload = ElectronicLoad(EQUIPMENT_ADDRESS.ELOAD)
scope = Oscilloscope(EQUIPMENT_ADDRESS.SCOPE)
eload_channel = EQUIPMENT_ADDRESS.ELOAD_CHANNEL


load_type = input(">> Load Type (cc/cr/nl): ")
test = "AC Cycling"
operation = "CV"
condition = "1007 rework"
project = "APPS EVAL/Single Stage GAN/Test Data (CV)"
excel_name = f"{test}_{condition}"
waveforms_folder = f'C:/Users/{username}/Desktop/{project}/{date}/{test}/'
path = path_maker(f'{waveforms_folder}')


iout_channel = 1
vin_channel = 2
ids_channel = 3
vout_channel = 4

"""USER INPUT"""
vin_list = [90,115,130,180,230,265,277,300]
percent_list = [100]

pulse_count = 5
pulse_list = [0.5, 1] 

vout = 50
iout_nom = 1.5
cr_load = vout/iout_nom

iout_list = [(percent/100) * iout_nom for percent in percent_list]
cr_list = [(cr_load/(percent/100) if percent != 0 else 0) for percent in percent_list]



if load_type == 'cc': load_list = iout_list
if load_type == 'cr': load_list = cr_list


"""SCOPE CHANNEL"""
iout_channel = 1
vin_channel = 2
ids_channel = 3
vout_channel = 4


class SCOPE_CONFIG():

    """
        MEASUREMENT SETTINGS OPTIONS: "MAX,MIN,RMS,MEAN,PDELta"
        HIGH | LOW | AMPLitude | MAXimum | MINimum | PDELta |
                MEAN | RMS | STDDev | POVershoot | NOVershoot | AREA |
                RTIMe | FTIMe | PPULse | NPULse | PERiod | FREQuency |
                PDCYcle | NDCYcle | CYCarea | CYCMean | CYCRms |
                CYCStddev | PULCnt | DELay | PHASe | BWIDth | PSWitching |
                NSWitching | PULSetrain | EDGecount | SHT | SHR | DTOTrigger |
                PROBemeter | SLERising | SLEFalling
    """
    """
        COLOR OPTIONS:

        - LIGHT_BLUE
        - YELLOW
        - PINK
        - GREEN
        - BLUE
        - ORANGE

        returns : state of the channel ('ON' or 'OFF')
    """

    TRIGGER_CHANNEL = 4
    TRIGGER_LEVEL = 10
    TRIGGER_EDGE = 'NEG'

    TIME_POSTIION = 20
    TIME_SCALE = 20e-3

    ZOOM_ENABLE = False
    ZOOM_POS = 50
    ZOOM_REL_SCALE = 1

    class CH1():
        REL_X_POS = 20
        ENABLE = 'ON'

        SCALE = 0.3
        POSITION = -4
        BANDWIDTH = 20
        LABEL = "LYT8 VOUT"
        MEASURE = "MAX,MIN"
        COLOR = "GREEN"
        COUPLING = "DCLimit"
        OFFSET = 0

    class CH2():
        REL_X_POS = 40
        ENABLE = 'ON'

        SCALE = 300
        POSITION = 3
        BANDWIDTH = 20
        LABEL = "VIN"
        MEASURE = "MAX,RMS"
        COLOR = "YELLOW"
        COUPLING = "AC"
        OFFSET = 0

    class CH3():
        REL_X_POS = 60
        ENABLE = 'ON'

        SCALE = 2
        POSITION = -4
        BANDWIDTH = 500
        LABEL = "PRI_IDS"
        MEASURE = "MAX"
        COLOR = "LIGHT_BLUE"
        COUPLING = "DCLimit"
        OFFSET = 0

    class CH4():
        REL_X_POS = 80
        ENABLE = 'ON'
        
        SCALE = 10
        POSITION = -4
        BANDWIDTH = 20
        LABEL = "LYT8 VOUT"
        MEASURE = "MAX,MIN"
        COLOR = "PINK"
        COUPLING = "DCLimit"
        OFFSET = 0

    class CURSOR():

        class CURSOR_1():
            CHANNEL = 1
            ENABLE = False
            X1 = 0
            X2 = 0
            Y1 = 0
            Y2 = 0
            TYPE = 'HOR'
        
        class CURSOR_2():
            CHANNEL = 2
            ENABLE = False
            X1 = 0
            X2 = 0
            Y1 = 0
            Y2 = 0
            TYPE = 'HOR'
        
        class CURSOR_3():
            CHANNEL = 3
            ENABLE = False
            X1 = 0
            X2 = 0
            Y1 = 0
            Y2 = 0
            TYPE = 'HOR'
        
        class CURSOR_4():
            CHANNEL = 4
            ENABLE = False
            X1 = 0
            X2 = 0
            Y1 = 0
            Y2 = 0
            TYPE = 'HOR'

def scope_settings():

    CH1 = SCOPE_CONFIG.CH1
    CH2 = SCOPE_CONFIG.CH2
    CH3 = SCOPE_CONFIG.CH3
    CH4 = SCOPE_CONFIG.CH4
    CURSOR_1 = SCOPE_CONFIG.CURSOR.CURSOR_1
    CURSOR_2 = SCOPE_CONFIG.CURSOR.CURSOR_2
    CURSOR_3 = SCOPE_CONFIG.CURSOR.CURSOR_3
    CURSOR_4 = SCOPE_CONFIG.CURSOR.CURSOR_4
        
    scope.channel_settings(state=CH1.ENABLE, channel=1, scale=CH1.SCALE, position=CH1.POSITION, label=CH1.LABEL,
                            color=CH1.COLOR, rel_x_position=CH1.REL_X_POS, bandwidth=CH1.BANDWIDTH, coupling=CH1.COUPLING, offset=CH1.OFFSET)
    
    scope.channel_settings(state=CH2.ENABLE, channel=2, scale=CH2.SCALE, position=CH2.POSITION, label=CH2.LABEL,
                            color=CH2.COLOR, rel_x_position=CH2.REL_X_POS, bandwidth=CH2.BANDWIDTH, coupling=CH2.COUPLING, offset=CH2.OFFSET)
    
    scope.channel_settings(state=CH3.ENABLE, channel=3, scale=CH3.SCALE, position=CH3.POSITION, label=CH3.LABEL,
                            color=CH3.COLOR, rel_x_position=CH3.REL_X_POS, bandwidth=CH3.BANDWIDTH, coupling=CH3.COUPLING, offset=CH3.OFFSET)
    
    scope.channel_settings(state=CH4.ENABLE, channel=4, scale=CH4.SCALE, position=CH4.POSITION, label=CH4.LABEL,
                            color=CH4.COLOR, rel_x_position=CH4.REL_X_POS, bandwidth=CH4.BANDWIDTH, coupling=CH4.COUPLING, offset=CH4.OFFSET)
    
    if CH1.ENABLE != 'OFF': scope.measure(1, CH1.MEASURE)
    if CH2.ENABLE != 'OFF': scope.measure(2, CH2.MEASURE)
    if CH3.ENABLE != 'OFF': scope.measure(3, CH3.MEASURE)
    if CH4.ENABLE != 'OFF': scope.measure(4, CH4.MEASURE)

    scope.record_length(50E6)
    scope.time_position(SCOPE_CONFIG.TIME_POSTIION)
    scope.time_scale(SCOPE_CONFIG.TIME_POSTIION)

    scope.remove_zoom()
    if SCOPE_CONFIG.ZOOM_ENABLE == True:
        scope.add_zoom(rel_pos=SCOPE_CONFIG.ZOOM_POS, rel_scale=SCOPE_CONFIG.ZOOM_REL_SCALE)
    
    trigger_channel = SCOPE_CONFIG.TRIGGER_CHANNEL
    trigger_level = SCOPE_CONFIG.TRIGGER_LEVEL
    trigger_edge = SCOPE_CONFIG.TRIGGER_EDGE
    scope.edge_trigger(trigger_channel, trigger_level, trigger_edge)

    scope.stop()

    if CURSOR_1.ENABLE: scope.cursor(channel=CURSOR_1.CHANNEL, cursor_set=1, X1=CURSOR_1.X1, X2=CURSOR_1.X2, Y1=CURSOR_1.Y1, Y2=CURSOR_1.Y2, type=CURSOR_1.TYPE)
    if CURSOR_2.ENABLE: scope.cursor(channel=CURSOR_2.CHANNEL, cursor_set=2, X1=CURSOR_2.X1, X2=CURSOR_2.X2, Y1=CURSOR_2.Y1, Y2=CURSOR_2.Y2, type=CURSOR_2.TYPE)
    if CURSOR_3.ENABLE: scope.cursor(channel=CURSOR_3.CHANNEL, cursor_set=3, X1=CURSOR_3.X1, X2=CURSOR_3.X2, Y1=CURSOR_3.Y1, Y2=CURSOR_3.Y2, type=CURSOR_3.TYPE)
    if CURSOR_4.ENABLE: scope.cursor(channel=CURSOR_4.CHANNEL, cursor_set=4, X1=CURSOR_4.X1, X2=CURSOR_4.X2, Y1=CURSOR_4.Y1, Y2=CURSOR_4.Y2, type=CURSOR_4.TYPE)

"""GENERIC FUNCTIONS"""

def discharge_output(times):    
    ac.turn_off()
    for i in range(times):
        for i in range(1,9):
            eload.channel[i].cr = 10
            eload.channel[i].turn_on()
        sleep(0.5)

        for i in range(1,9):
            eload.channel[i].turn_off()
            eload.channel[i].short_off()
        sleep(1)

        for i in range(1,9):
            eload.channel[i].cc = 1
            eload.channel[i].short_on()
            eload.channel[i].turn_on()
        sleep(0.5)

        for i in range(1,9):
            eload.channel[i].turn_off()
            eload.channel[i].short_off()
        sleep(0.5)

def screenshot(filename, path):
    global waveform_counter

    scope.get_screenshot(filename, path)
    print(filename)
    waveform_counter += 1

def collect_data(vin, iout, pulse):

    ## EDIT THIS FUNCTION DEPENDING ON TEST
    
    """COLLECT DATA"""
    vac = vin

    labels, values = scope.get_measure(vout_channel)
    vout_max = float(values[0])

    

    labels, values = scope.get_measure(ids_channel)
    ids_max = float(values[0])


    vout_overshoot = abs(100*(vout_max-vout)/vout)

    if vout_overshoot > 13.4: vout_overshoot_result = 'FAIL'
    else: vout_overshoot_result = 'PASS'

    if load_type == 'cc':
        load_percent = int(iout*100/iout_nom)
        load_value = float(f"{iout:.2f}")

    if load_type == 'cr':
        load_percent = int(cr_load*100/iout)
        load_value = float(f"{iout:.2f}")

    if load_type == 'nl':
        load_percent = 0
        load_value = 0

    if load_type != 'nl':
        labels, values = scope.get_measure(iout_channel)
        iout_max = float(values[0])
    else: iout_max = 0

    output_list = [pulse, load_percent, load_value, vac, vout_max, vout_overshoot, vout_overshoot_result, iout_max, ids_max]
    
    return output_list

def _set_time_scale(pulse):
    time_division = pulse
    time_scale = 2*time_division
    start_soak = 3*time_scale
    off_time = pulse
    on_time = pulse
    end_soak = 3*time_scale
    delay = start_soak + end_soak + (pulse_count*(off_time + on_time)) + off_time
    scope.time_scale(time_scale)

    return start_soak, off_time, on_time, end_soak

def operation():

    pms.current_range(2)
    pml.current_range(2)


    if load_type == 'cr':

        ################### COPY PASTE THIS CODE IN HEADER OF MAIN ###################
        df_header_list = ['Pulse Time (s)', 'CR Load (%)', 'CR Load (ohms)', 'Vin (VAC)',
                        'Vout_max (V)', 'Overshoot (%)', 'PASS/FAIL', 'Iout_max (A)', 'Ids_max (A)']
        df = create_header_list(df_header_list)
        ##############################################################################

        estimated_time = len(pulse_list)*len(cr_list)*len(vin_list)*8*15/60
        print(f"Estimated time: {estimated_time} mins.")

        scope.cursor(channel=vout_channel, cursor_set=1, X1=1, X2=1, Y1=0.92*vout, Y2=1.08*vout, type='HOR')


        for pulse in pulse_list:

            start_soak, off_time, on_time, end_soak = _set_time_scale(pulse)
            
            for cr in cr_list:
                
                for vin in vin_list:

                    start = start_timer()

                    eload.channel[eload_channel].cr = cr
                    eload.channel[eload_channel].turn_on()

                    scope.run_single()
                    soak(1)
                    ac.voltage = vin
                    ac.turn_on()

                    ac.ac_cycling(pulse_count, vin, start_soak, off_time, on_time, end_soak)
                    discharge_output(times=1)

                    percent = int(cr_load*100/cr) if cr != 0 else 0
                    crr = round(cr, 2)
                    filename = f"{pulse}s On-Off, {pulse_count} pulses - {vin}Vac, {crr}Ohms ({percent} CR load).png"
                    screenshot(filename, path)

                    output_list = collect_data(vin, cr, pulse)
                    export_to_excel(df, waveforms_folder, output_list, excel_name, sheet_name=f"summary_{load_type}", anchor="B2")

                    end_timer(start)

        print(f"\n\nFinal Data: \n")
        print(df)
                

    if load_type == 'cc':

        ################### COPY PASTE THIS CODE IN HEADER OF MAIN ###################
        df_header_list = ['Pulse Time (s)', 'CC Load (%)', 'CC Load (A)', 'Vin (VAC)',
                        'Vout_max (V)', 'Overshoot (%)', 'PASS/FAIL', 'Iout_max (A)', 'Ids_max (A)']
        df = create_header_list(df_header_list)
        ##############################################################################

        for pulse in pulse_list:

            start_soak, off_time, on_time, end_soak = _set_time_scale(pulse)

            for iout in iout_list:

                for vin in vin_list:


                    eload.channel[eload_channel].von = 0
                    eload.channel[eload_channel].cc = iout
                    eload.channel[eload_channel].turn_on()

                    scope.run_single()
                    soak(1)
                    ac.voltage = vin
                    ac.turn_on()
                    
                    ac.ac_cycling(pulse_count, vin, start_soak, off_time, on_time, end_soak)
                    scope.stop()
                    discharge_output(times=1)

                    percent = int(iout*100/iout_nom)
                    filename = f"{pulse}s On-Off, {pulse_count} pulses - {vin}Vac, {iout}A ({percent} CC load).png"
                    screenshot(filename, path)

                    output_list = collect_data(vin, iout, pulse)
                    export_to_excel(df, waveforms_folder, output_list, excel_name, sheet_name=f"summary_{load_type}", anchor="B2")
        
        print(f"\n\nFinal Data: \n")
        print(df)

    if load_type == 'nl':

        eload.channel[eload_channel].turn_off()

        ################### COPY PASTE THIS CODE IN HEADER OF MAIN ###################
        df_header_list = ['Pulse Time (s)', 'No Load', 'No Load', 'Vin (VAC)',
                        'Vout_max (V)', 'Overshoot (%)', 'PASS/FAIL', 'Iout_max (A)', 'Ids_max (A)']
        df = create_header_list(df_header_list)
        ##############################################################################

        scope.channel_settings(state="OFF", channel=iout_channel, scale=ch1_scale, position=ch1_position, label=ch1_label,
                            color=ch1_color, rel_x_position=ch1_rel_x_position, bandwidth=ch1_bw, coupling=ch1_coupling, offset=0)

        scope.edge_trigger(ids_channel, 3, 'POS')

        for pulse in pulse_list:

            start_soak, off_time, on_time, end_soak = _set_time_scale(pulse)

            for vin in vin_list:

                eload.channel[eload_channel].turn_off()
                
                scope.run_single()
                soak(1)
                ac.voltage = vin
                ac.turn_on()

                ac.ac_cycling(pulse_count, vin, start_soak, off_time, on_time, end_soak)
                discharge_output(times=1)

                filename = f"{pulse}s On-Off, {pulse_count} pulses - {vin}Vac, No Load.png"
                screenshot(filename, path)

                iout = 0
                output_list = collect_data(vin, iout, pulse)
                export_to_excel(df, waveforms_folder, output_list, excel_name, sheet_name=f"summary_{load_type}", anchor="B2")

        print(f"\n\nFinal Data: \n")
        print(df)

def main():
    global waveform_counter
    discharge_output(times=1)
    scope_settings()
    operation()
        
if __name__ == "__main__":
    headers(test)
    main()
    footers(waveform_counter)