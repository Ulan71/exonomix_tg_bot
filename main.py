import pdfplumber
import pandas as pd
import time

pdf_path = r"pdf_files/test.pdf"

file_list = []

reg_number_key = "Регистрационный номер ESF"
number_key = "Номер учетной системы"
date1_key = "2. Дата выписки"
date2_key = "Дата совершения оборот"


with pdfplumber.open(pdf_path) as pdf:
    page = pdf.pages[0]
    table = page.extract_table()


pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', None)

found_data = {}

for row in table:
    for cell in row:
        if not cell == None:

            if reg_number_key in cell:
                found_data["reg_number"] = cell[-34:]

            if number_key in cell:
                found_data["number"] = cell[-12:]

            if date1_key in cell:
                found_data["date1"] = cell[-12:]

print(found_data)

# for i in range(len(table)):
#     df = pd.DataFrame(table[1:], columns=table[i])
#     if any("Наименование\nтоваров, работ,\nуслуг" in str(col) for col in df.columns):
#

df = pd.DataFrame(table[1:], columns=table[1])

# for i in range(10):
#     names = df.iloc[:, i]
#
#     for n in names:
#         print(n)
#
#         if "Наименование\nтоваров, работ,\nуслуг" in str(n):
#             print("This")
#             break
