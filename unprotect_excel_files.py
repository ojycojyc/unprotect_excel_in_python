#!/usr/bin/env python
# coding: utf-8

import pandas as pd
import win32com.client
import xlrd
from itertools import product
from pathlib import Path

# Get the paths of folder holding all the protected and unprotected files
protected_data_dir = Path('data_protected/')
unprotected_data_dir = Path('data_unprotected/')

## Get the paths of files to unprotect
# https://stackoverflow.com/questions/39909655/listing-of-all-files-in-directory
existing_paths = [file for file in protected_data_dir.glob('**/*') if file.is_file()]

# Define function that opens and resave a password protected file
# Function is used later and has to be defined early.
# args: 
def unprotect_xlsx(file):
    xcl = win32com.client.DispatchEx('Excel.Application')
    
    pw_str = '1' 

    # construct new path with old file name
    writepath = unprotected_data_dir/file.name 
    
    wb = xcl.Workbooks.Open(file.absolute(),0,False,None,pw_str,pw_str,True)
    
    # Unprotect each sheet in the workbook
    # COM object Worksheet uses index which starts at 1
    for index in range (1,wb.Sheets.Count+1):
        ws = wb.Sheets(index)
        ws.Unprotect(pw_str) #unprotect each sheet

    # Unprotect the workbook itself
    wb.Unprotect(pw_str)

    # Disable displays, save file, close it and quit
    xcl.DisplayAlerts = False
    wb.SaveAs(writepath.absolute().__str__()) #function requires string as argument
    wb.Close(True)
    xcl.Quit()

# For each file path in the array, try to unprotect the file
for path in existing_paths:
    print(path) # print the path
    print("Try to unprotect it.")       
    try: 
        unprotect_xlsx(path)
    except Exception as f:
        print(f)
        print("Unprotecting failed")


