from functools import wraps
from math import isnan
from time import sleep, time
from abc import ABC, abstractmethod
import atexit
import os
import sys

try:
    import pyvisa
 

except:
    import pip
    pip.main(['install', 'pyvisa'])

    import pyvisa

GPIB_NO = 0
PORT = 0

type_to_address = {
    'gpib': f'GPIB{GPIB_NO}',
    'lan': 'TCPIP',
    'usb': 'USB',
    'serial': 'ASRL'
    }

def reject_nan(func):
    @wraps(func)
    def wrapped(*args, **kwargs):
        for _ in range(10):
            response = func(*args, **kwargs)
            if not isnan(response):
                return response
            sleep(0.2)
        return None
    return wrapped


''' GENERAL CLASS ''' 
class Equipment(ABC):
    def __init__(self, address, comm_type='gpib', timeout=20000):
        rm = pyvisa.ResourceManager()
        print(rm.list_resources())
        print(rm.list_resources()[0])
        if comm_type == 'serial':
            self.address = f'{type_to_address[comm_type]}{address}'
            print(self.address)
        else:
            self.address = f'{type_to_address[comm_type]}::{address}'
        self.timeout = timeout
        try:
            self.device = rm.open_resource(self.address)
            self.device.timeout = self.timeout
            atexit.register(self.cleanup)
        except Exception as e:
            print("\n\n\n" + "*"*100)
            print("CHECK COMMS ADDRESSES")
            print("*"*100 + "\n\n\n")
            raise pyvisa.VisaIOError(e)
            

        self._id = 0

    def __repr__(self):
        return (f'{self.__class__.__name__}'
                 '(address={self.address}, timeout={self.timeout})')

    @abstractmethod
    def cleanup(self):
        pass

    @property
    def id(self):
        self._id = self.write('*IDN?')
        return self._id

    def close(self):
        self.device.close()

    def write(self, command):
        response = None
        reply_available = False
        try:
            if "BDMM" in command:
                reply_available = True
                command = command.split(':')[1]

            if "?" in command:
                response = self.device.query(command).strip()
            else:
                self.device.write(command)

            if reply_available:
                self.device.read()
                reply_available = False

        except Exception as e:
            raise pyvisa.VisaIOError(e)

        return response

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''' EQUIPMENTS ''' 
class Yokogawa_Oscilloscope_DLM5058(Equipment):
    def __init__(self, address, comm_type='lan', timeout=30000):
        super().__init__(address, comm_type, timeout)

 ##################### CH ON/OFF ###################    
    def channel_state(self, channel = 1, state ='ON'):
        self.write(f':CHAN{channel}:DISP {state}')

    def channel_state_all_off(self): 
        for channel in range(1, 9): 
            self.write(f':CHAN{channel}:DISP OFF')

 ##################### V&T /DIV #################### 
    
    def set_vdiv(self, channel = 1, vdiv = '1'): #5mV to 100V
        self.write(f':CHAN{channel}:VDIV {vdiv}V')

    def set_tdiv(self, tdiv = '10US'): #500ps to 500s
        self.write(f':TIM:TDIV {tdiv}')

 ##################### POSITION ####################

    def set_vert_position(self, channel = 1, vert_pos = 0): #-4div to 4div
        self.write(f':CHAN{channel}:POS {vert_pos}')

    def set_horiz_position(self, horiz_pos = 0): #Percentage 0% to 100%
        self.write(f':TRIG:POS {horiz_pos}')

 ##################### ATTENUATION #################

    def set_probe_attenuation(self, channel = 1, attenuation = '10'): #Add C{attenuation} if current
        self.write(f':CHAN{channel}:PROBE:MODE {attenuation}')

 ##################### TRIGGER #####################

    def set_trigger_mode(self, trig_mode = 'AUTO'): #AUTO, NORMal, A(uto)LEVel, NSINgle
        self.write(f':TRIG:MODE {trig_mode}')

    def set_trigger_level(self, channel = 1, trig_lev = '1V', slope = 'RISE'): 
        self.write(f':TRIG:SIMP:SOUR {channel}')
        self.write(f':TRIG:SIMP:LEV {trig_lev}')
        self.write(f':TRIG:SIMP:SLOP {slope}')

    def set_trigger_ncount(self, ncount = 0): 
        self.write(f':TRIG:SCOUNT {ncount}')

  ##################### MISC #################

    def set_label(self, channel = 1, label = ''):
        self.write(f':CHAN{channel}:LAB:DEF "{label}"')
        self.write(f':CHAN{channel}:LAB:DISP ON')

    def set_bw(self, channel = 1, bw = ''):
        self.write(f':CHAN{channel}:BWIDTH {bw}')

    def set_coupling(self, channel = 1, coupling = ''):
        self.write(f':CHAN{channel}:COUPLING {coupling}')

    def invert(self, channel = 1, state = 'OFF'):
        self.write(f':CHAN{channel}:INVERT {state}')

    def set_color(self, channel = 1, color = 'BLUE'): #{BLUE|BGReen|CYAN|DBLue|GRAY|GREen|LBLue|LGReen|MAGenta|MGReen|ORANge|PINK|PURPle|RED|SPINk|YELLow}
        self.write(f':DISP:COLOR:CHAN{channel} {color}')

 ###################### PROBE QUERY ###############
    def probe_query(self, channel = 1):
        result = {
            "probe": self.write(f':CHAN{channel}:PROBE:MODE?').split(' ')[1:],
            "vdiv": self.write(f':CHAN{channel}:VDIV?').split(' ')[1:],
            "coup": self.write(f':CHAN{channel}:COUP?').split(' ')[1:],
            "bw": self.write(f':CHAN{channel}:BWID?').split(' ')[1:],
            "inv": self.write(f':CHAN{channel}:INV?').split(' ')[1:],
            "label": self.write(f':CHAN{channel}:LABEL?').split(' ')[1:],
            "on/off": self.write(f':CHAN{channel}:DISPLAY?').split(' ')[1:],
            "color": self.write(f':DISPLAY:COLOR:CHAN{channel}?').split(' ')[1:],
            "position": self.write(f':CHAN{channel}:POS?').split(' ')[1:]
        }
        return result

 ###################### HORIZONTAL QUERY ##########
    def horizontal_query(self):
        result = {
            "trigChannel": self.write(f':TRIG:SIMP:SOUR?').split(' ')[1:],
            "trigLevel": self.write(f':TRIG:SIMP:LEV?').split(' ')[1:],
            "trigSlope": self.write(f':TRIG:SIMP:SLOP?').split(' ')[1:],
            "trigMode": self.write(f':TRIG:MODE?').split(' ')[1:],
            "position": self.write(f':TRIG:POS?').split(' ')[1:],
            "ncount": self.write(f':TRIG:SCOUNT?').split(' ')[1:],
            "tdiv": self.write(f':TIM:TDIV?').split(' ')[1:]
        }
        return result

    def trig_level_query(self):
        return self.write(f':TRIG:SIMP:LEV?').split(' ')[1:][0]
 ################### SNGLE ####################
    def single(self,timeout = 0):
        return self.write(f':SSTART? {timeout}').split(' ')[1]

 ################### RUN STOP ################
    def run(self):
        self.write(f':START')    
    
    def stop(self):
        self.write(f':STOP')
 ##################### MEASURE #####################

    def disp_meas_off(self, channel = 1): 
        self.write(f':MEAS:CHAN{channel}:ALL OFF')

    def disp_meas_off_all(self): 
        for channel in range(1, 9): 
            self.write(f':MEAS:CHAN{channel}:ALL OFF')

    def disp_meas(self, channel = 1, parameter = 'FREQ', state = 'ON'): 
        self.write(f':MEAS:CHAN{channel}:{parameter}:STATE {state}') 
        ''' List of Parameters Drop-Down
            AMPLitude
            AVERage
            AVGFreq
            AVGPeriod
            BWIDth
            DELay
            DT
            DUTYcycle
            ENUMber
            FALL
            FREQuency
            HIGH
            LOW
            MAXimum
            MINimum
            NOVershoot
            NWIDth
            PERiod
            PNUMber
            POVershoot
            PTOPeak
            PWIDth
            RISE
            RMS
            SDEViation
            TY1Integ
            TY2Integ
            V1
            V2'''

    def get_meas_value(self, channel = 1, parameter = 'FREQ'):
        sleep(0.05) 
        result = self.write(f':MEAS:CHAN{channel}:{parameter}:VAL?').split(' ')[1]
        return result 

    def disp_meas_statistics(self, mode = 'CONT'): #OFF|ON|CONTinuous|CYCLe|HISTory
        self.write(f':MEAS:MODE {mode}')

    def set_sample_count(self, count = 1000): 
        self.write(f':ACQ:COUNT {count}')
        self.write(f':TRIG:MODE NORM')
        status = int(self.write(f':STATus:CONDition?'))
        while status > 0:
            status = int(self.write(f':STATus:CONDition?'))

    def get_min_mean_max_sd(self, channel = 1, parameter = 'FREQ'):
        result = {
            "maximum": self.write(f':MEASure:CHANnel{channel}:{parameter}:MAXimum?').split(' ')[1],
            "mean": self.write(f':MEASure:CHANnel{channel}:{parameter}:MEAN?').split(' ')[1],
            "minimum": self.write(f':MEASure:CHANnel{channel}:{parameter}:MINimum?').split(' ')[1],
            "sdeviation": self.write(f':MEASure:CHANnel{channel}:{parameter}:SDEViation?').split(' ')[1]
        }
        return result

    def query_meas_state(self, channel = 1):
        if self.write(f':MEASure:CHANnel{channel}:DEL:REF:SOURCE?').split(' ')[1] != 'TRIG':
            reference = self.write(f':MEASure:CHANnel{channel}:DEL:REF:TRACE?').split(' ')[1]
            if len(buffer)<2:
                reference = 'CH'+reference
        else:
            reference = 'TRIG'

        result = {
            "max": self.write(f':MEASure:CHANnel{channel}:MAX:STAT?').split(' ')[1],
            "min": self.write(f':MEASure:CHANnel{channel}:MIN:STAT?').split(' ')[1],
            "pp": self.write(f':MEASure:CHANnel{channel}:PTOP:STAT?').split(' ')[1],
            "high": self.write(f':MEASure:CHANnel{channel}:HIGH:STAT?').split(' ')[1],
            "low": self.write(f':MEASure:CHANnel{channel}:LOW:STAT?').split(' ')[1],

            "amp": self.write(f':MEASure:CHANnel{channel}:AMPL:STAT?').split(' ')[1],
            "rms": self.write(f':MEASure:CHANnel{channel}:RMS:STAT?').split(' ')[1],
            "mean": self.write(f':MEASure:CHANnel{channel}:AVER:STAT?').split(' ')[1],
            "sdev": self.write(f':MEASure:CHANnel{channel}:SDEV:STAT?').split(' ')[1],
            
            "pover": self.write(f':MEASure:CHANnel{channel}:POV:STAT?').split(' ')[1],
            "nover": self.write(f':MEASure:CHANnel{channel}:NOV:STAT?').split(' ')[1],
            "pulse": self.write(f':MEASure:CHANnel{channel}:PNUM:STAT?').split(' ')[1],
            "edge": self.write(f':MEASure:CHANnel{channel}:ENUM:STAT?').split(' ')[1],
            
            "v1": self.write(f':MEASure:CHANnel{channel}:V1:STAT?').split(' ')[1],
            "v2": self.write(f':MEASure:CHANnel{channel}:V2:STAT?').split(' ')[1],
            "delta": self.write(f':MEASure:CHANnel{channel}:DT:STAT?').split(' ')[1],
            "integpos": self.write(f':MEASure:CHANnel{channel}:TY1I:STAT?').split(' ')[1],
            "integ": self.write(f':MEASure:CHANnel{channel}:TY2I:STAT?').split(' ')[1],
            
            "freq": self.write(f':MEASure:CHANnel{channel}:FREQ:STAT?').split(' ')[1],
            "period": self.write(f':MEASure:CHANnel{channel}:PER:STAT?').split(' ')[1],
            "avefreq": self.write(f':MEASure:CHANnel{channel}:AVGF:STAT?').split(' ')[1],
            "aveper": self.write(f':MEASure:CHANnel{channel}:AVGP:STAT?').split(' ')[1],
            "burst": self.write(f':MEASure:CHANnel{channel}:BWID:STAT?').split(' ')[1],
            
            "rise": self.write(f':MEASure:CHANnel{channel}:RISE:STAT?').split(' ')[1],
            "fall": self.write(f':MEASure:CHANnel{channel}:FALL:STAT?').split(' ')[1],
            "poswid": self.write(f':MEASure:CHANnel{channel}:PWID:STAT?').split(' ')[1],
            "negwid": self.write(f':MEASure:CHANnel{channel}:NWID:STAT?').split(' ')[1],
            "duty": self.write(f':MEASure:CHANnel{channel}:DUTY:STAT?').split(' ')[1],
            
            "delay": '1' if self.write(f':MEASure:CHANnel{channel}:DEL:STAT?').split(' ')[1] == 'ON' else '0',

            "sourpol": '1' if self.write(f':MEASure:CHANnel{channel}:DEL:MEAS:SLOPE?').split(' ')[1] == 'FALL' else '0',
            "sourcount": self.write(f':MEASure:CHANnel{channel}:DEL:MEAS:COUNT?').split(' ')[1],
            
            "ref": buffer, 
            
            "refpol": '1' if self.write(f':MEASure:CHANnel{channel}:DEL:REF:SLOPE?').split(' ')[1] == 'RISE' else '0',
            "refcount": self.write(f':MEASure:CHANnel{channel}:DEL:REF:COUNT?').split(' ')[1],

            "statistics": '0' if (self.write(f':MEAS:MODE?').split(' ')[1] == 'ON') or (self.write(f':MEAS:MODE?').split(' ')[1] == 'OFF') else '1'
        }
        return result 

    def set_delay_source_slope(self, channel = 1, slope = 'RISE'):
        self.write(f':MEASure:CHANnel{channel}:DEL:MEAS:SLOPE {slope}')

    def set_delay_source_count(self, channel = 1, count = 1):
        self.write(f':MEASure:CHANnel{channel}:DEL:MEAS:COUNT {count}')

    def set_delay_ref_slope(self, channel = 1, slope = 'RISE'):
        self.write(f':MEASure:CHANnel{channel}:DEL:REF:SLOPE {slope}')

    def set_delay_ref_count(self, channel = 1, count = 1):
        self.write(f':MEASure:CHANnel{channel}:DEL:REF:COUNT {count}')

    def set_delay_ref_source(self, channel = 1, source = 'TRIG'):
        self.write(f':MEASure:CHANnel{channel}:DEL:REF:SOURCE {source}')
    
    def set_delay_ref_trace(self, channel = 1, trace = 'TRIG'):
        self.write(f':MEASure:CHANnel{channel}:DEL:REF:TRACE {trace}')

 ##################### SCREENSHOT ###################
    def set_drive(self, drive = 'USB'): #FLAShmem|NETWork|USB 
        self.write(f' :IMAGE:SAVE:DRIVE {drive}')
    
    def set_background(self, background = 'REV'): #COLor|GRAY|OFF|REVerse 
        self.write(f' :IMAGE:TONE {background}')

    def set_filename(self, filename = ''): 
        self.write(f' :IMAGE:SAVE:NAME "{filename}"')

    def screenshot(self): 
        self.write(f' :IMAGE:EXECUTE')
    
 ##################### CLOSE ########################
    def stop(self):
        self.write(':STOP')

    def cleanup(self):
        self.close()


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

class Tektronix_SigGen_AFG31000(Equipment):
    def __init__(self, address, comm_type='gpib', timeout=20000):
        super().__init__(address, comm_type, timeout)

 ##################### CH ON/OFF ###################    
    def channel_state(self, channel = 1, state ='ON'):
        self.write(f':OUTP{channel}:STAT {state}')

    def channel_state_all_off(self): 
        for channel in range(1, 3): 
            self.write(f':OUTP{channel}:STAT OFF')

    def set_load_impedance(self, channel = 1, impedance = 'INF'): 
        self.write(f':OUTP{channel}:IMPedance {impedance}')

 ##################### CONTINUOUS ###################
    def out_cont_sine(self, channel = 1, frequency = '10kHz', phase = '0', low = '-500mV', high = '500mV', units = 'VPP'): 
        self.write(f' :SOURce{channel}:FUNCtion:SHAPe SIN')
        self.write(f' :SOURce{channel}:FREQuency:FIXed {frequency}')
        self.write(f' :SOURce{channel}:PHASe:ADJust {phase} DEG')
        self.write(f' :SOURce{channel}:VOLTage:LEVel:IMMediate:LOW {low}')
        self.write(f' :SOURce{channel}:VOLTage:LEVel:IMMediate:HIGH {high}')
        self.write(f' :SOURce{channel}:VOLTage:UNIT {units}')

    def out_cont_square(self, channel = 1, frequency = '10kHz', phase = '0', low = '-500mV', high = '500mV', units = 'VPP'): 
        self.write(f' :SOURce{channel}:FUNCtion:SHAPe SQU')
        self.write(f' :SOURce{channel}:FREQuency:FIXed {frequency}')
        self.write(f' :SOURce{channel}:PHASe:ADJust {phase} DEG')
        self.write(f' :SOURce{channel}:VOLTage:LEVel:IMMediate:LOW {low}')
        self.write(f' :SOURce{channel}:VOLTage:LEVel:IMMediate:HIGH {high}')
        self.write(f' :SOURce{channel}:VOLTage:UNIT {units}')

    def out_cont_ramp(self, channel = 1, frequency = '10kHz', phase = '0', low = '-500mV', high = '500mV', units = 'VPP', symmetry = 50.0): 
        self.write(f' :SOURce{channel}:FUNCtion:SHAPe RAMP')
        self.write(f' :SOURce{channel}:FREQuency:FIXed {frequency}')
        self.write(f' :SOURce{channel}:PHASe:ADJust {phase} DEG')
        self.write(f' :SOURce{channel}:VOLTage:LEVel:IMMediate:LOW {low}')
        self.write(f' :SOURce{channel}:VOLTage:LEVel:IMMediate:HIGH {high}')
        self.write(f' :SOURce{channel}:VOLTage:UNIT {units}')
        self.write(f' :SOURce{channel}:FUNCtion:RAMP:SYMMetry {symmetry}')              

    def out_cont_pulse(self, channel = 1, frequency = '10kHz', phase = '0', low = '-500mV', high = '500mV', units = 'VPP', duty = 999, width = 'ABC'): #Choose duty or width
        self.write(f' :SOURce{channel}:FUNCtion:SHAPe PULSE')
        self.write(f' :SOURce{channel}:FREQuency:FIXed {frequency}')
        self.write(f' :SOURce{channel}:PHASe:ADJust {phase} DEG')
        self.write(f' :SOURce{channel}:VOLTage:LEVel:IMMediate:LOW {low}')
        self.write(f' :SOURce{channel}:VOLTage:LEVel:IMMediate:HIGH {high}')
        self.write(f' :SOURce{channel}:VOLTage:UNIT {units}')
        
        if (duty != 999):
            self.write(f' :SOURce{channel}:PULSe:DCYCle {duty}')
        if (width != 'ABC'):
            self.write(f' :SOURce{channel}:PULSe:WIDTh {width}')
                
    def out_DC(self, channel = 1, voltage = '5V'): 
        self.write(f' :SOURce{channel}:FUNCtion:SHAPe DC')
        self.write(f' :SOURce{channel}:VOLTage:LEVel:IMMediate:OFFSet {voltage}')

 ##################### CLOSE ########################
    def stop(self):
        self.write(':STOP')

    def cleanup(self):
        self.close()


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

class Keithley_DC_2230G(Equipment):
    def __init__(self, address, comm_type='gpib', timeout=20000):
        super().__init__(address, comm_type, timeout)

 #################### SET VOL&CURR #################
    def set_volt_curr(self, channel = 1, voltage = 0.0, current = 0.1):
        self.write(f'APPL CH{channel}, {voltage}, {current}')

    def channel_state(self, channel = 1, state = 'ON'): #Individual
        self.write(f'INST CH{channel}')  
        self.write(f'CHAN:OUTP {state}')

    def channel_state_all(self, state = 'ON'): #For all channels
        self.write(f'OUTP {state}')          

 #################### CLOSE ########################
    def stop(self):
        self.write(':STOP')

    def cleanup(self):
        self.close()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

class Iseg_HVDC_HPS1k5W(Equipment):
    def __init__(self, address, comm_type='gpib', timeout=20000):
        super().__init__(address, comm_type, timeout)

 #################### SET VOL&CURR #################
    def set_volt_curr(self, voltage = 0.0, current = 0.1):
        self.write(f':VOLT {voltage}')
        self.write(f':CURR {current}')
        
    def state(self, state = 'ON'): #Individual
        self.write(f':VOLT {state}')
         

 #################### CLOSE ########################
    def stop(self):
        self.write(':STOP')

    def cleanup(self):
        self.close()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

class Fluke_BDMM_8808A(Equipment):

    def __init__(self, address, comm_type='serial', timeout=20000):
        super().__init__(address, comm_type, timeout)

 #################### BDMM #################
    def remote_mode(self):
        self.write(f'BDMM:REMS')

    def set_func(self, function = 'VDC'): #VDC, VAC, AAC, ADC, OHMS, CONTinuity, DIODE, FREQ
        self.write(f'BDMM:{function}')

    def get_meas_value(self):
        value = self.write(f'BDMM:VAL?').split(' ')[0]
        return value

    def meas_hold(self):
        return self.write(f'BDMM:HOLD')

 #################### CLOSE ########################
    def stop(self):
        self.write(':STOP')

    def cleanup(self):
        self.close()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

class Rigol_ELoad_DL3021(Equipment):
    def __init__(self, address, comm_type='usb', timeout=20000):
        super().__init__(address, comm_type, timeout)

 #################### Eload #################

    def state(self, state = 'ON'): 
        self.write(f' :SOUR:INP:STAT {state}')

    def set_mode(self, mode = 'CURR', level = 0): #CURRent, RESistance, VOLTage, POWer
        self.write(f' :SOUR:FUNC {mode}')
        self.write(f' :SOUR:{mode}:LEV:IMM {level}')

    def short(self):
        self.write(f' SYSTEM:KEY 33')

 #################### CLOSE ########################
    def stop(self):
        self.write(':STOP')

    def cleanup(self):
        self.close()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

class Temptronics_Thermostream_ATS535(Equipment):
    def __init__(self, address, comm_type='gpib', timeout=20000):
        super().__init__(address, comm_type, timeout)

 #################### CONTROLS #################
    def set_nozzle(self, direction = 'DOWN'):
        if direction == "UP":
            self.write(f'HEAD 1')
        elif direction == "DOWN":
            self.write(f'HEAD 0')

    def set_flow(self, flow = 'ON'):
        if flow == "ON":
            self.write(f'FLOW 1')
        elif flow == "OFF":
            self.write(f'FLOW 0')
        
    def setpoint_temp(self, temp = 25.0):
        self.write(f'SETP {temp}') #-99.9 to 225.0C

    def query_temp(self):
        return self.write(f'TEMP?')

 #################### CLOSE ########################
    def stop(self):
        self.write(':STOP')

    def cleanup(self):
        self.close()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

class Yokogawa_PowerMeter_WT310E(Equipment):
    def __init__(self, address, comm_type='gpib', timeout=20000):
        super().__init__(address, comm_type, timeout)

        self.volt_disp_ch = 1
        self.curr_disp_ch = 2
        self.power_disp_ch = 3
        self.pfactor_disp_ch = 4
 #################### SETTINGS #################
    def set_mode(self, mode = 'DC'): #RMS | VMEAN | DC
        self.write(f':INPUT:MODE {mode}')

    def set_voltage_range(self, Range = 'AUTO'):
        if Range == 'AUTO':
            self.write(f':INPUT:VOLTAGE {Range}')
        else:
            self.write(f':INPUT:CFACTOR 3 ')    
            self.write(f':INPUT:VOLTAGE:RANGE {Range}') #15, 30, 60, 150, 300, 600(V)
    
    def set_current_range(self, Range = 'AUTO'):
        if Range == 'AUTO':
            self.write(f':INPUT:CURRENT {Range}')
        else:
            self.write(f':INPUT:CFACTOR 3 ')    
            self.write(f':INPUT:CURRENT:RANGE {Range}') #5, 10, 20, 50, 100, 200, 500(mA), 1, 2, 5, 10, 20(A)
 ##################### DISPLAY ####################
    def item_ch_update(self, disp_mode, item_channel):
        if disp_mode == 'U':
            self.volt_disp_ch = item_channel
        elif disp_mode == 'I':
            self.current_disp_ch = item_channel
        elif disp_mode == 'P':
            self.power_disp_ch = item_channel
        elif disp_mode == 'LAMBDA':
            self.pfactor_disp_ch = item_channel

    def display(self, disp_A = 'U', disp_B = 'I', disp_C = 'P', disp_D = 'LAMBDA'): #Voltage Current Power and PF only available can add more as per request
        self.item_ch_update(disp_mode = disp_A, item_channel = 1)
        self.item_ch_update(disp_mode = disp_B, item_channel = 2)
        self.item_ch_update(disp_mode = disp_C, item_channel = 3)
        self.item_ch_update(disp_mode = disp_D, item_channel = 4)

        self.write(f':DISPLAY:ITEM1 {disp_A},1')
        self.write(f':DISPLAY:ITEM2 {disp_B},1')
        self.write(f':DISPLAY:ITEM3 {disp_C},1')
        self.write(f':DISPLAY:ITEM4 {disp_D},1')

    def query_voltage(self):
        print(f':NUMERIC:NORMAL:ITEM{self.volt_disp_ch} U, 1')

        self.write(f':NUMERIC:NORMAL:ITEM{self.volt_disp_ch} U, 1')
        return float(self.write(f':NUMERIC:NORMAL:VAL? {self.volt_disp_ch}'))

    def query_current(self):
        print(f':NUMERIC:NORMAL:ITEM{self.curr_disp_ch} I, 1')

        self.write(f':NUMERIC:NORMAL:NORMAL:ITEM{self.curr_disp_ch} I, 1')
        return float(self.write(f':NUMERIC:NORMAL:VAL? {self.curr_disp_ch}'))

    def query_power(self):
        print(f':NUMERIC:NORMAL:ITEM{self.power_disp_ch} P, 1')

        self.write(f':NUMERIC:NORMAL:ITEM{self.power_disp_ch} P, 1')
        return float(self.write(f':NUMERIC:NORMAL:VAL? {self.power_disp_ch}'))      

    def query_pfactor(self):
        self.write(f':NUMERIC:NORMAL:ITEM{self.pfactor_disp_ch} LAMBDA, 1')
        return float(self.write(f':NUMERIC:NORMAL:VAL? {self.pfactor_disp_ch}'))      

 #################### CLOSE ########################
    def stop(self):
        self.write(':STOP')

    def cleanup(self):
        self.close()