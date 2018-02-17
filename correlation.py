''' Experiment on trying to correlate parameters

Last edited by RayS 1/24/2016
'''

import numpy as np
import datetime as dt
#from nnLib import *
import pylab as pl
import xlrd
import xlwt
import calendar as cal
import xlsxwriter
import os
#from StockLib import *

#trainFile = '/home/ray/Documents/programming/python/stocks/web_scraping/scrape_1-9-2016X.xls'
os.chdir('/home/ray/Documents/programming/python/stocks/web_scraping') # windows directory
filename = 'excelTest.xlsx'
sheetName = "Sheet1"


## [subcode] read info into arrays

book = xlrd.open_workbook(filename) # open the workbook
data_sheet = book.sheet_by_name(sheetName) # get sheet by name

startRow = 1
endRow = None

header = data_sheet.row(0)
for (i, text) in enumerate(header):
    header[i] = header[i].value
print("header length = ", len(header))
#print(header)

data = []
colMax = 22
col = 0
while col < colMax:
    data.append(data_sheet.col_values(col,startRow, endRow))
    temp = []
    if col % 2 == 0: # on even columns (dates)
        for d in data[col]:
            if d is not "":           
                tmp = (dt.datetime.strptime(d,'%Y-%m-%d').date())
                temp.append(str(tmp))
        data[col] = temp
    else: # for odd columns (data)
        for d in data[col]:
            if d is not "":           
                temp.append(d)
        data[col] = temp
    print("length of column read = ", len(data[col]))
    col += 1
print("number of lists = ", len(data))
#print(type(data[0][5]), data[0][5])
#print(type(data[1][5]), data[1][5])

## [subcode] calculate correlation coefficients

x = 1
while x < 20:
    length = min(len(data[1]), len(data[x+2]))
    #print(length)
    coef = (np.corrcoef(data[1][:length], data[x+2][:length]))
    print(header[x+2], ": ", coef[0,1])
    x += 2

'''
list1 = [1,2,3,4]
list2 = [7,5,3,0]
corr = np.corrcoef(list1, list2)[0, 1]
print(corr)
'''
