from engineering_notation import EngNumber as eng

m = (6.5E-6-0.7E-6)/(100E3-25E3)
f_sw = input("Enter switching frequency:")
f_sw = float(f_sw)
t_on = 6.5E-6 + m*(f_sw-100E3)
t_on = eng(t_on)
print(f"expected t_on: {t_on}s")





