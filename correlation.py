''' Experiment on trying to correlate parameters

Last edited by RayS 2/18/2017
'''

import numpy as np
import datetime as dt
#from nnLib import *
import pylab as pl
import xlrd
# import xlwt
import calendar as cal
import xlsxwriter
import os
#from StockLib import *
from cleanData import *  # this has the function to read in the excel file
from plotStocks import *  # this has the function to calculate moving averages, etc


if __name__ == '__main__':

    filename = 'data/clean_' + str(dt.date.today()) + '.xlsx'
    # filename = 'scrape_2018-02-06.xlsx'
    sheetName = "clean_data"
    print("input filename = " + filename)

    wb = load_workbook(filename)
    ws = wb.get_sheet_by_name(sheetName)

    data = read_in_columns(ws)
    # print(data[0][:5])

    PE_col = 7
    price_col = 1
    date_col = 0
    sample = 200
    header_length = 1

    price = data[price_col][header_length:]
    PE = data[PE_col][header_length:]
    date = data[date_col][header_length:]

    PRave = movingAverage(price, sample)
    PRscale = scale(price[:-sample], PRave)
    print("length of PRscale = " + str(len(PRscale)))

    PEave = movingAverage(PE, sample)
    PEscale = scale(PE[1:-sample], PEave)
    print("length of PEscale = " + str(len(PEscale)))

    date_obj = []
    for d in date:
        date_obj.append(dt.datetime.strptime(d, "%b %d, %Y").date())
    print("length of date_obj = " + str(len(date_obj)))

    x = 2
    while x < len(data):
        length = min(len(data[1][1:]), len(data[x][1:]))
        # length = 10
        price = data[1][1:length]
        other_list = data[x][1:length]
        print(length, len(price), len(other_list))
        # print(price)
        # print(other_list)
        # print(length, len(data[1][1:length]), len(data[x+1][1:length]))
        coef = (np.corrcoef(price, other_list))
        # coef = (np.corrcoef(data[1][1:length], data[x+1][1:length]))
        print(data[x][0], ": ", coef[0, 1])
        # print(data[x+2][1], ": ", coef)
        x += 1


#trainFile = '/home/ray/Documents/programming/python/stocks/web_scraping/scrape_1-9-2016X.xls'
# os.chdir('/home/ray/Documents/programming/python/stocks/web_scraping') # windows directory
# filename = 'excelTest.xlsx'
# sheetName = "Sheet1"


## [subcode] read info into arrays

# book = xlrd.open_workbook(filename) # open the workbook
# data_sheet = book.sheet_by_name(sheetName) # get sheet by name
#
# startRow = 1
# endRow = None
#
# header = data_sheet.row(0)
# for (i, text) in enumerate(header):
#     header[i] = header[i].value
# print("header length = ", len(header))
# #print(header)
#
# data = []
# colMax = 22
# col = 0
# while col < colMax:
#     data.append(data_sheet.col_values(col,startRow, endRow))
#     temp = []
#     if col % 2 == 0: # on even columns (dates)
#         for d in data[col]:
#             if d is not "":
#                 tmp = (dt.datetime.strptime(d,'%Y-%m-%d').date())
#                 temp.append(str(tmp))
#         data[col] = temp
#     else: # for odd columns (data)
#         for d in data[col]:
#             if d is not "":
#                 temp.append(d)
#         data[col] = temp
#     print("length of column read = ", len(data[col]))
#     col += 1
# print("number of lists = ", len(data))
#print(type(data[0][5]), data[0][5])
#print(type(data[1][5]), data[1][5])

## [subcode] calculate correlation coefficients


'''
list1 = [1,2,3,4]
list2 = [7,5,3,0]
corr = np.corrcoef(list1, list2)[0, 1]
print(corr)
'''
