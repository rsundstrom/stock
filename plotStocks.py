''' Basic plotting test
Last edited by RayS 2/17/2018
'''

import pylab as pl
from openpyxl import Workbook, load_workbook
import datetime as dt
from cleanData import *  # this has the function to read in the excel file


def movingAverage(values, samples=60):
    ave = []
    s = 0
    while s < len(values)-samples:
        ave.append(sum(values[s:s+samples])/samples)
        s += 1
    return ave

def scale(values, factors):
    scaled = []
    for a, b in zip(values, factors):
        scaled.append(a/b)
    return scaled


if __name__ == '__main__':

    filename = 'data/scrape_' + str(dt.date.today()) + '.xlsx'
    # filename = 'scrape_2018-02-06.xlsx'
    sheetName = "data"
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
    print("length of datex = " + str(len(date_obj)))

    PE1 = PE[:50]
    PE1.reverse()
    # pl.plot(price, label="Price")
    # pl.plot(PE1, label="PE")
    # pl.plot(date_obj[:100], PRscale[:100], label="Price")
    pl.plot(date_obj, price, label="Price")
    # pl.plot(date_obj[:100], PEscale[:100], label="PE")
    pl.legend(loc='lower left')
    pl.show()
