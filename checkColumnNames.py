#!/usr/bin/python3
from openpyxl import load_workbook
import re

def listNames(ExcelFile):
    print(ExcelFile)
    print("______________________")
    # Load an existing Excel file
    wb = load_workbook(ExcelFile)
    
    # Select a worksheet
    ws = wb["Sheet1"]  # or use wb.active
   
    num_columns = ws.max_column

    for i in range(1, num_columns+1):
        #(i.e., row 1, column 1 for A1)
        cell_object = ws.cell(row=1, column=i)  # This accesses cell C5
        
        # You can then get or set its value
        cell_value = cell_object.value
        print(cell_value)

    print("______________________")
    print("")

    


listNames("jupiter.xlsx")
listNames("saturn.xlsx")
listNames("uranus.xlsx")
listNames("neptune.xlsx")
