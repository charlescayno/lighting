from openpyxl import load_workbook
sample = load_workbook(filename="juv.xlsx")
for sheet in sample: sheet.protection.disable()
sample.save(filename="juve.xlsx")
sample.close()