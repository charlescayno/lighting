from Equipment_Lib import Keithley_DC_2230G, Fluke_BDMM_8808A
from time import time,sleep

dc = Keithley_DC_2230G('1')
dc.set_volt_curr(channel ='CH1', voltage = 10.0, current = 1.0)
dc.set_volt_curr(channel ='CH2', voltage = 20.0, current = 1.0)
dc.set_volt_curr(channel ='CH3', voltage = 5.0, current = 1.0)

dc.channel_state(channel ='CH1', state = 'ON')
sleep(2)
dc.channel_state(channel ='CH2', state = 'ON')
sleep(2)
dc.channel_state(channel ='CH3', state = 'ON')
sleep(2)

dc.channel_state_all(state='OFF')