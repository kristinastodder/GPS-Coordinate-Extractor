#! python3
#This program extracts the latitude from a (x,y) format

import pyperclip, openpyxl

# Open the spreadsheet and loop through values
wb = openpyxl.load_workbook('California.xlsx')
sheet = wb['Sheet1']
for r in range(2, sheet.max_row + 1):
    location = sheet.cell(row=r, column=17).value
    pyperclip.copy(location)

    #extract latitude coordinates
    latitude = location[1:10]
    print(latitude)
