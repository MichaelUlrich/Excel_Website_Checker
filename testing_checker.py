# Excel Website Checker
# 
# Goals:    Be able to load Excel file format (.xlsx .xlsm)
#           Be able to parse such files for link and ping 
#           Append if link is still valid or 404'd for quicker manual sorting post parse
# Created by Michael Ulrich December 20, 2018
#
# Copyright (c) Michael Ulrich 2019

import pandas as pd

def load_file(file_name):
    open_file = ""
    try:
        open_file = pd.read_excel(file_name)
    except:
        print("Error Opening File...")
    print(open_file)
    print(open_file.values[0][1]) # returns a NumPy style array of values in xlsx file

def parse_file(file_name):
    # Blank
    print("")
def write_to_file(file_name):
    # Blank
    print("")
def main():
    file_name = "test.xlsx"
    load_file(file_name)
main()
