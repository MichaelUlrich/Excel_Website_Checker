# Excel Website Checker
# 
# Goals:    Be able to load Excel file format (.xlsx .xlsm)
#           Be able to parse such files for link and ping 
#           Append if link is still valid or 404'd for quicker manual sorting post parse
# Created by Michael Ulrich December 20, 2018
#
# Copyright (c) Michael Ulrich 2019

import pandas as pd
import numpy as np
import requests

def load_file(file_name):
    open_file = ""
    print(file_name)
    try:
        open_file = pd.read_excel(file_name)
        print(open_file)
    except:
        print("Error Opening File...") 
    print(open_file.values[0][0]) 
    ping_website(open_file.values[0][0])

def ping_website(url):
   # if 'https://' not in url and 'http://' not in url:
        # url = "https://" + url
    response = requests.get(url)
    print(response.status_code)
    return response.status_code

def parse_file(file_name):
    print("")

def write_to_file(file_name):
    print("")

def main():
    file_name = 'test.xlsx'
    load_file(file_name)
main()
