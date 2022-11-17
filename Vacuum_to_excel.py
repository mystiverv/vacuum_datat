#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Jun  3 12:48:27 2021

@author: zionirving-singh
"""

import os
import xlsxwriter
import tkinter as tk 
from tkinter.filedialog import askopenfilename

    
#gets rid of tkinter blank window
root = tk.Tk()
root.withdraw()
root.wm_attributes('-topmost', 1)
 
#the following line lets the user select a file that can be processed later
file_path = askopenfilename()
#creates an alias "startpath" for whenever you want os to get a directory name
startpath = os.path.dirname(file_path)

name = (os.path.basename(file_path)).strip(".txt")
#creates a new excel workbook named pressure_summary
workbook = xlsxwriter.Workbook(os.path.join(startpath,((str(name) + " data summary" '.xlsx'))))

#creates the two subsheets in the workbook
rawData = workbook.add_worksheet('Raw Data')
press = workbook.add_worksheet('Pressure Graph')

#adds a blank chart into the pressure subsheet
pressPlot = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'})

#converts the scientific notation of the pressure data into regular numbers
num_format = workbook.add_format({'num_format': '0.0000'})

#opens the file you select and then picks out only the pressure data
filename = file_path                                 #
vac_data = []                                         # Declare an empty list named mylines.
with open (filename , 'rt') as myfile:               # Open pump data text for reading.
    for myline in myfile:                            # For each line in the file 
        vac_data.append(myline[4:12])    # isolates the pressure readout

# removes empty lines from vac_data list
while '' in vac_data:
    vac_data.remove('')
    
#gets the length of the vacuum data list to create a properly scaled time frame
lengthVacData = len(vac_data)


#converts all the scientific notation vacuum data to floating point data
floatyVacData = map(float, vac_data)

#creates an empty list then divides each x value in vac data by half to account for a pressure readout only being taken every 30 seconds
time = []
for x in range(len(vac_data)):
    time.append(x/2)
rawData.write(0, 0, 'Pressure (torr)')
rawData.write(0, 1, 'Time (mins)')
#writes the pressure values to the raw data worksheet with the specified number format
for row_num, data in enumerate(floatyVacData):
    rawData.write(row_num + 1, 0, data, num_format)

#writes the time data to the excel file 
for row_num, data in enumerate(time):
    rawData.write(row_num + 1, 1, data)
    


            

pressPlot.add_series({
    'values': ['Raw Data', 1, 0, lengthVacData, 0],
    'categories': ['Raw Data', 1, 1, lengthVacData, 1],
    'trendline': {'type': 'linear',
    'display_equation': True,
    'display_r_squared': True}
})

pressPlot.set_y_axis({
    'crossing': '-2',
    'major_gridlines': {
    'visible': False},
    'name': 'Pressure (torr)',
    'min': 0.2, 
    'max': 5,
    'num_font': {'name': 'Calibri', 'size': 10, 'bold': True}
})
pressPlot.set_x_axis({
    'crossing': '-2',
    'major_gridlines': {
    'visible': False},
    'name': 'Time (mins)',
    'min': 0, 
    'max': 120,
    'num_font': {'name': 'Calibri', 'size': 10, 'bold': True}
})

press.insert_chart('C5', pressPlot)

workbook.close()
