import xlrd  
import csv 
import pandas as pd 
import chilkat
import sys

sheet = xlrd.open_workbook("Web2Pair - Copy.xlsx").sheet_by_index(0) 
tt =  open("T.csv", 'w', newline="")
col = csv.writer(tt) 

for row in range(sheet.nrows): 
    col.writerow(sheet.row_values(row)) 
tt.close()
csvv = chilkat.CkCsv()
csvv.put_HasColumnNames(True)

success = csvv.LoadFile("T.csv")
if (success != True):
    print(csvv.lastErrorText())
    sys.exit()

success = csvv.SetCell(0,22,"baguette")

success = csvv.SaveFile("V1.csv")
if (success != True):
    print(csvv.lastErrorText())

