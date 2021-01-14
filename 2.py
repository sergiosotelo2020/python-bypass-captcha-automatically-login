# import all required library 
import xlrd  
import csv 
import pandas as pd 
import sys
import chilkat
import time
from xlsxwriter.workbook import Workbook
import os
loc = 'Web2Pair - Copy.xlsx'
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
# sheet1 = xlrd.open_workbook("Web2Pair.xlsx").sheet_by_index(0) 

tt = open("T.csv", 'w', newline="")
col = csv.writer(tt) 

for row in range(sheet.nrows): 
    col.writerow(sheet.row_values(row))  
# df = pd.DataFrame(pd.read_csv("T.csv")) 
# df

tt.close()
csvv = chilkat.CkCsv()
csvv.put_HasColumnNames(True)

time.sleep(2)

success = csvv.LoadFile("T.csv")
time.sleep(2)
if (success != True):
    print(csvv.lastErrorText())
    sys.exit()

success = csvv.SetCell(0,22,"baguette")

success = csvv.SaveFile("V.csv")
if (success != True):
    print(csvv.lastErrorText())


time.sleep(1)

csvfile = "V.csv"

workbook = Workbook('Web2Pair-1.xlsx')
worksheet = workbook.add_worksheet()
with open(csvfile, 'rt', encoding='utf8') as f:
    reader = csv.reader(f)
    for r, row in enumerate(reader):
        for c, col in enumerate(row):
            worksheet.write(r, c, col)
workbook.close()  

time.sleep(2)
# os.remove("T.csv")
# os.remove("V.csv")

print("done")