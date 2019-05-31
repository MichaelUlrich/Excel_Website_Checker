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
import xlsxwriter

def load_file(file_name):
    open_file = "" 
    print(file_name)
    try:
        open_file = pd.read_excel(file_name)
        print(open_file)
    except:
        print("Error Opening File...") 
    file_length = len(open_file)
    url_array = ['NaN'] * file_length # Fill with default values
    response_array = ['NaN'] * file_length
    
    for i in range(file_length):
        print('Loop: ' + str(i)) # debug
        url_array[i] = open_file.values[i][0]
        response_array[i] = ping_website(url_array[i])
    print(url_array) # debug
    print(response_array) # debug
    write_to_file(file_name, url_array, response_array)
   
def ping_website(url):
    url = str(url)
    # Check if URL is valid, if not append https for request
    if 'https://' not in url and 'http://' not in url:
        url = "http://" + url
    print("Pinging: " + url + "...") # debug
    response = requests.get(url, timeout = 2)
    print(response.status_code)
    return response.status_code
   
def parse_file(file_name):
    print("")

def write_to_file(file_name, url_array, response_array):
    df = pd.DataFrame({'URL': url_array, 'Response Code': response_array}) # Need to re-write URLs?
    try:
        writer = pd.ExcelWriter(file_name, engine='xlsxwriter') # Convert Pandas Excel writer to xlsxWriter
        df.to_excel(writer, sheet_name='sheet1', index=False) # Convert the dataframe to an XlsxWriter Excel object.
        workbook = writer.book # Select xlsx book
        worksheet = writer.sheets['sheet1'] # Select specific sheet
        
        red_format = workbook.add_format({'bg_color': '#FFC7CE','font_color': '#9C0006'}) #sets font/background to red
        green_format = workbook.add_format({'bg_color': '#C6EFCE','font_color': '#006100'}) #sets font/background to green
        yellow_format = workbook.add_format({'bg_color': '#FFEB9C','font_color': '#9C6500'}) #sets font/background to yellow

        worksheet.conditional_format('B1:B1048576', {'type':'cell', 'criteria': '=', 'value': 404, 'format': red_format}) # B1:B1048576 - selects all possible cells in column B, maybe better way?
        worksheet.conditional_format('B1:B1048576', {'type':'cell', 'criteria': '=', 'value': 200, 'format': green_format})
        worksheet.conditional_format('B1:B1048576', {'type':'cell', 'criteria': '=', 'value': 429, 'format': yellow_format})
        writer.save()
    except Exception as e:
        print('Error - ', e)

def main():
    file_name = 'test.xlsx'
    total_urls = '' # Get total URLs from xlsx
     # Will declare array of size of total_urls once implemented
    load_file(file_name) # Replace with function to generate entire array
    # response_array[0] = ping_website(url_array[0]) # Replace with function that will be passed url_array, and ping each url and return array of HTTP response codes
    # print(url_array)
    # print(response_array)
    # write_to_file(file_name, url_array, response_array)
main()
