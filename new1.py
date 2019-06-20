#TODO: File location naming
#TODO: Sheet naming

import os
import openpyxl
import datetime
from datetime import date, datetime, timedelta

os.chdir('C:\\Users\yasirc\Desktop\parseData') # enter correct filepath for project here
wb = openpyxl.load_workbook('ngt_log.xlsx') # filename being analyzed
sheet = wb['sheet1'] # name of the sheet that is being analyzed
cycleTimes = []
coulCount = []
dateTimes = []
timeDeltas = []
res = []

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
    dateTimes.append(datetime.combine(date.today(), i))

print("\nTIMES BETWEEN CYCLES")

print("\nCOLOUMB COUNT")
coulCount.append(sheet.cell(row=sheet.max_row, column=8).value)
for i in coulCount:
    print(i)

# input() #keeps cmd window open after script execution

'''
refObj = openpyxl.chart.Reference(sheet, min_col=2, min_row=2, max_col=2, max_row=sheet.max_row)
seriesObj = openpyxl.chart.Series(refObj, title='First series')
chartObj = openpyxl.chart.LineChart()
chartObj.title = 'My Chart'
chartObj.append(seriesObj)
sheet.add_chart(chartObj, 'C5')
wb.save('sampleChart.xlsx')

'''
