# Excel Website Checker
# 
# Goals:    1. Be able to load Excel file format (.xlsx .xlsm)
#           2. Be able to parse such files for link and ping 
#           3. Append if link is still valid or 404'd for quicker manual sorting post parse
#           4. GUI:
#               a. Load specific file
#               b. Validate and parse file
#               c. Generate new file
#               d. Add new URLs to file for validating
# Created by Michael Ulrich December 20, 2018
#
# Copyright (c) Michael Ulrich 2019

import pandas as pd
import numpy as np
import requests
import xlsxwriter
from tkinter import *

class Window(Frame):
    def __init__(self, master=None): # Initilize gui frame
        Frame.__init__(self, master) # used pass parameters
        self.master = master # reference to master tk window
        self.init_window()
    def init_window(self):
        self.master.title("Excel Checker")
        self.pack(fill=BOTH, expand=1) # use the ful window
        menu = Menu(self.master) #generate menu
        self.master.config(menu=menu)

        file= Menu(menu, tearoff=0) # create file obj
        file.add_command(label="Open") # open button
        file.add_command(label="New") # open button
        menu.add_cascade(label="File", menu=file) # file tab

        # edit = Menu(menu, tearoff=0) #generate file obj in window to add to menu bar
        menu.add_command(label="Exit", command=self.client_exit) # exit button directly on menu bar
        # menu.add_cascade(label="Edit", menu=edit) #add file to menu dropdown
        
    def client_exit(self):
        exit()


def load_file(file_name):
    open_file = "" 
    print(file_name)
    try:
        open_file = pd.read_excel(file_name)
        print(open_file)
    except:
        print("Error Opening File...") 
        exit(1) # terminate
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
    try:
        response = requests.get(url, timeout = 4)
    except requests.exceptions.RequestException as e:
        return 'Bad URL'
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
        
        RED_FRMT = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})    #sets font/background to red
        GREEN_FRMT = workbook.add_format({'bg_color': '#C6EFCE','font_color': '#006100'})   #sets font/background to green
        YELLOW_FRMT = workbook.add_format({'bg_color': '#FFEB9C','font_color': '#9C6500'})  #sets font/background to yellow
        BLACK_FRMT = workbook.add_format({'bg_color': '000000','font_color': '#DCDCDC'})    #sets background to black and font to white

        worksheet.conditional_format('B1:B1048576', {'type':'cell', 'criteria': '=', 'value': 404, 'format': RED_FRMT}) # B1:B1048576 - selects all possible cells in column B, maybe better way?
        worksheet.conditional_format('B1:B1048576', {'type':'cell', 'criteria': '=', 'value': 200, 'format': GREEN_FRMT})
        worksheet.conditional_format('B1:B1048576', {'type':'cell', 'criteria': 'equal to', 'value': '"Bad URL"', 'format': BLACK_FRMT})
        worksheet.conditional_format('B1:B1048576', {'type':'cell', 'criteria': '=', 'value': 429, 'format': YELLOW_FRMT})
        
        writer.save()
    except Exception as e:
        if 'Errno 13' in str(e):
            print('\n## File Still Open - Can not write to file')
        else: 
            print("\n## Error - ", e)

 
def main():
    file_name = 'test.xlsx'
    total_urls = '' # Get total URLs from xlsx

    root = Tk()
    root.geometry("1000x300")
    app = Window(root)
    root.mainloop()
     # Will declare array of size of total_urls once implemented
    # load_file(file_name) # Replace with function to generate entire array
    # response_array[0] = ping_website(url_array[0]) # Replace with function that will be passed url_array, and ping each url and return array of HTTP response codes
    # print(url_array)
    # print(response_array)
    # write_to_file(file_name, url_array, response_array)

main()
