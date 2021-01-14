# import all required library 
import xlrd  
import csv 
import pandas as pd 
import sys
import chilkat
import time
from xlsxwriter.workbook import Workbook
import os
loc = 'Web.xlsx'
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
time.sleep(1)

success = csvv.LoadFile("T.csv")
if (success != True):
    print(csvv.lastErrorText())
    sys.exit()
k = 0
edit_urls = ['dd', 'dd', 'dd']
statuss = ['dd', 'dd', 'dd']
ee = len(edit_urls)
print('urls:' + str(ee))
ss = len(statuss)
print('status:' + str(ss))
if ee < ss:
    ee = ss

while k < ee:
    success = csvv.SetCell(k,22,edit_urls[k])
    success = csvv.SaveFile("V.csv")
    success = csvv.SetCell(k,23,statuss[k])
    success = csvv.SaveFile("V.csv")
    k += 1


if (success != True):
    print(csvv.lastErrorText())
time.sleep(1)
csvfile = "V.csv"

workbook = Workbook('Web1.xlsx')
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