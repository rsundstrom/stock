''' Plotting to look for correlations
Last edited by RayS 3/28/2016
'''

import pylab as pl
import xlrd
import datetime as dt

# os.chdir('/home/ray/Documents/programming/python/stocks/web_scraping') # set working directory
filename = 'clean_' + str(dt.date.today()) + '.xlsx'
sheetName = "clean_data"
print("filename = " + filename)

book = xlrd.open_workbook(filename) # open the workbook
data_sheet = book.sheet_by_name(sheetName) # get sheet by name

startRow = 2
endRow = None

header = data_sheet.row(0)
for (i, text) in enumerate(header):
    header[i] = header[i].value
print("header length = ", len(header))
#print(header)

data = []
colMax = 12
col = 0
while col < colMax:
    data.append(data_sheet.col_values(col, startRow, endRow))
    temp = []
    # if col % 2 == 0: # on even columns (dates)
    if col == 0: # on even columns (dates)
        for d in data[col]:
            if d is not "":           
                tmp = (dt.datetime.strptime(d, '%b %d, %Y').date())
                temp.append(str(tmp))
        data[col] = temp
    else:
        for d in data[col]:
            if d is not "":           
                tmp = (d)
                temp.append(float(tmp))
        data[col] = temp
        
    print("length of column read = ", len(data[col]))
    col += 1
print("number of lists = ", len(data))

## [subcode] Scaling functions

def movingAverage(values, samples=60):
    ave = []
    s = 0
    while s < len(values)-samples:
        ave.append(sum(values[s:s+samples])/samples)
        s += 1
    return ave

def scale(values, factors):
    scaled = []
    for a,b in zip(values, factors):
        scaled.append(a/b)
    return scaled
    

## [subcode]

y = 7
z = 1
sample = 200

PRave = movingAverage(data[z], sample)
PRscale = scale(data[z][:-sample], PRave)
print("length of PRscale = " + str(len(PRscale)))
 
PEave = movingAverage(data[y], sample)
PEscale = scale(data[y][:-sample], PEave)
print("length of PEscale = " + str(len(PEscale)))

datex = []
for d in data[0]:
    datex.append(dt.datetime.strptime(d,'%Y-%m-%d').date())
print("length of datex = " + str(len(datex)))
   
#pl.plot(data[y], label=header[y])
#pl.plot(ave60, label=header[y] + " 60 MA")
pl.plot(datex[:1500], PRscale[:1500], label=header[z])
pl.plot(datex[:1500], PEscale[:1500], label=header[y])
pl.legend(loc='lower left')
pl.show()

## [subcode]

date = []
x = 6
y = 7

print(len(data[x]), len(data[y]))
#data[x].reverse()
#data[y].reverse()
for x in data[x]:
    date.append(dt.datetime.strptime(x,'%Y-%m-%d').date())
#print(date)


#pl.plot(date, data[y], label=header[y])
pl.plot(data[y], label=header[y])
pl.plot(ave, label=header[y])
pl.legend(loc='lower left')
pl.show()

