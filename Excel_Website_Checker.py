# Excel Website Checker
# 
# Goals:    Be able to load Excel file format (.xlsx .xlsm)
#           Be able to parse such files for link and ping 
#           Append if link is still valid or 404'd for quicker manual sorting post parse
# Created by Michael Ulrich December 20, 2018
#
# Copyright (c) Michael Ulrich 2018

# import panda
import numpy as np
import pandas as pd
import requests
import pyping
import socket
# load excel file and check if valid file
def load_file(file):
    print("\nI return opened/failed to open")
    open_file = pd.read_excel(file)
    print(open_file)
    print("\nSorted URL's")
   # for i in 5:
   #     website_ping(open_file.head(i))
    print(open_file.sort_values(['URL'], ascending=False))
    print("\n head 1: ")
    print(open_file.head(2)) # value at current position

# ping website, returning if successful or 404
def website_ping(url):
    if 'https://' not in url and 'http://' not in url:
        # print('broken url')
        url = "https://" + url
        print('new url = ' + url)
    # r = requests.get(url)
    # print(r.status_code)
    # print(r.json)
    # print("\nI return success ping/404")
    # print("The url passed: " + url)
    ping.verbose_ping(url, count=3)
    # delay = ping.Ping('www.wikipedia.org', timeout=2000).do()
   # except socket.error e:
   #     print ("Ping Error:", e)

# append result to final column of excel file after ping
def append_result():
    print("\nI append to last column good/bad")

def parse_file(file):
    print("\nParsing File")
    # open excel file
    xlsx = pd.ExcelFile(file)
    sheet1 = xlsx.parse(0)      # first sheet
    total_rows = sheet1.shape
    temp = total_rows[0]
    print(total_rows[0])
    row = sheet1.iloc[1]        # row
    column = sheet1.iloc[:, 1]  # column
    print("Row: " + row[0])     # will be URL cell
    for i in range(temp):
        row = sheet1.iloc[i]
        website_ping(row[0])
    print("Column: " + column[0])
# GUI?

def main():
    temp_file = "test.xlsx"
    # print("temp file: " + temp_file)
    # load_file(temp_file)
    parse_file(temp_file)
    # website_ping("FAKE_URL")
    append_result()
main()