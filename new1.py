#TODO: File location naming
#TODO: Sheet naming

import os
import openpyxl

os.chdir('C:\\Users\yasirc\Desktop\parseData')
wb = openpyxl.load_workbook('ngt_log.xlsx')
sheet = wb['sheet1']
cycleTimes = []
cycleTimes.append(sheet.cell(row=2, column=2).value.time())

print("ROW | VALUE")
for i in range(2, sheet.max_row):
    current = sheet.cell(row=i, column=6).value
    next =  sheet.cell(row=i+1, column=6).value
    if int(current) >= 0 and int(next) <= 0 or int(current) <= 0 and int(next) >= 0:
        cycleTimes.append(sheet.cell(row=i, column=2).value)
        cycleTimes.append(sheet.cell(row=(i+1), column=2).value)
        print(i+1, next)

print("\nTIMES")        
cycleTimes.append(sheet.cell(row=sheet.max_row, column=2).value)        
for i in cycleTimes:
    print(i)
