# import all required library 
import xlrd  
import csv 
import pandas as pd 
  

sheet = xlrd.open_workbook("Web2Pair.xlsx").sheet_by_index(0) 
  
col = csv.writer(open("T.csv", 'w', newline="")) 
  
for row in range(sheet.nrows): 
    col.writerow(sheet.row_values(row))  
df = pd.DataFrame(pd.read_csv("T.csv")) 
df