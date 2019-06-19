#TODO: File location naming
#TODO: Sheet naming
#TODO: Remove deprecation warning

import os
os.chdir('C:\\Users\yasirc\Desktop\parseData')
import openpyxl
wb = openpyxl.load_workbook('ngt_log.xlsx')
sheet = wb.get_sheet_by_name('sheet1')
cycleTimes = []
cycleTimes.append(sheet.cell(row=2, column=2).value)
print("ROW | VALUE")
for i in range(2, sheet.max_row):
    current = sheet.cell(row=i, column=6).value
    next =  sheet.cell(row=i+1, column=6).value
    if int(current) >= 0 and int(next) <= 0 or int(current) <= 0 and int(next) >= 0:
        cycleTimes.append(sheet.cell(row=i, column=2).value)
        cycleTimes.append(sheet.cell(row=(i+1), column=2).value)
        print(i+1, next)
        
cycleTimes.append(sheet.cell(row=sheet.max_row, column=2).value)        
for i in cycleTimes:
    print(i)
