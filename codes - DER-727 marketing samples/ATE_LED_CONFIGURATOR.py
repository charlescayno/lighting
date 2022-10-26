from powi.equipment import LEDControl
import sys

led = LEDControl()
load = int(sys.argv[1])

led.voltage(load)

# from pyfirmata import Arduino, util
# import pyfirmata
# commPort = 11
# board = pyfirmata.Arduino(f'COM{commPort}')
# iterator = util.Iterator(board)
# iterator.start()
# RELAY1 = board.get_pin('d:10:o')
# RELAY1.write(1)




input("Press ENTER to end program.")
