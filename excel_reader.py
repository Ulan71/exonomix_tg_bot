from openpyxl import Workbook
from openpyxl import load_workbook

wb = load_workbook("sheet.xlsx")
sheet = wb.active
sheet.insert_cols(8)
for num, value in enumerate(sheet["G"]):
    print(num, value.value)
wb.save("sheet.xlsx")
