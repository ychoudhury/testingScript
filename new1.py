#TODO: File location naming
#TODO: Sheet naming

import os
import openpyxl
from openpyxl.chart import LineChart, Reference
from openpyxl import load_workbook


os.chdir('C:\\Users\yasirc\Desktop\parseData') # enter correct filepath for project here
wb = openpyxl.load_workbook('ngt_log.xlsx')
sheet = wb['sheet1'] # name of the sheet that is being analyzed
cycleTimes = []
coulCount = []
cycleTimes.append(sheet.cell(row=2, column=2).value.time()) # adds 00:00:00 as start time
coulCount.append(sheet.cell(row=2, column=8).value)

print("ROW | VALUE")
for i in range(2, sheet.max_row):
    current = sheet.cell(row=i, column=6).value
    next =  sheet.cell(row=i+1, column=6).value
    if int(current) >= 0 and int(next) <= 0 or int(current) <= 0 and int(next) >= 0:
        cycleTimes.append(sheet.cell(row=i, column=2).value)
        cycleTimes.append(sheet.cell(row=(i+1), column=2).value)
        coulCount.append(sheet.cell(row=i, column=8).value)
        coulCount.append(sheet.cell(row=(i+1), column=8).value)
        print(i+1, next)

print("\nTIMES")        
cycleTimes.append(sheet.cell(row=sheet.max_row, column=2).value)        
for i in cycleTimes:
    print(i)

print("\nTIMES BETWEEN CYCLES")
for i in cycleTimes:
    print

print("\nCOLOUMB COUNT")
coulCount.append(sheet.cell(row=sheet.max_row, column=8).value)
for i in coulCount:
    print(i)
    
# input() #keeps cmd window open after script execution
