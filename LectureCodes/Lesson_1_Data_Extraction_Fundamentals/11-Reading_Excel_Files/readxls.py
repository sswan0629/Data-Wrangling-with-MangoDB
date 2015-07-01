#!/usr/bin/env python
"""
Your task is as follows:
- read the provided Excel file
- find and return the min and max values for the COAST region (Column index: 1)
- find and return the time value for the min and max entries
- the time values should be returned as Python tuples

Please see the test function for the expected return format
"""

import xlrd
from zipfile import ZipFile
dataZipFile = "/Users/swan/Documents/MYFILES/Dropbox/MOOC/DataWranglingWithMangoDB/2013_ERCOT_Hourly_Load_Data"
datafile = "2013_ERCOT_Hourly_Load_Data.xls"


def open_zip(datafile):
    with ZipFile('{0}.zip'.format(datafile), 'r') as myzip:
        myzip.extractall()


def parse_file(datafile):
    workbook = xlrd.open_workbook(datafile)
    sheet = workbook.sheet_by_index(0)
    ### example on how you can get the data
    sheet_data = [[sheet.cell_value(r, col) for r in range(sheet.nrows)] for col in range(sheet.ncols)] #read into list
    col_coast = sheet_data[1]
    col_hours = sheet_data[0]
    col_coast.pop(0)
    min_coast = min(col_coast)
    max_coast = max(col_coast)

    min_time_decimal = col_hours[col_coast.index(min_coast)+1] #plus 1 due to the delete the column names
    min_time = xlrd.xldate_as_tuple(min_time_decimal, 0)
    max_time_decimal = col_hours[col_coast.index(max_coast)+1]
    max_time = xlrd.xldate_as_tuple(max_time_decimal,0)

    avg = sum(col_coast) / float(len(col_coast))

    ### other useful methods:
    # print "\nROWS, COLUMNS, and CELLS:"
    # print "Number of rows in the sheet:", 
    # print sheet.nrows
    # print "Type of data in cell (row 3, col 2):", 
    # print sheet.cell_type(3, 2)
    # print "Value in cell (row 3, col 2):", 
    # print sheet.cell_value(3, 2)
    # print "Get a slice of values in column 3, from rows 1-3:"
    # print sheet.col_values(3, start_rowx=1, end_rowx=4)

    # print "\nDATES:"
    # print "Type of data in cell (row 1, col 0):", 
    # print sheet.cell_type(1, 0)
    # exceltime = sheet.cell_value(1, 0)
    # print "Time in Excel format:",
    # print exceltime
    # print "Convert time to a Python datetime tuple, from the Excel float:",
    # print xlrd.xldate_as_tuple(exceltime, 0)
    
    
    data = {
            'maxtime': (0, 0, 0, 0, 0, 0),
            'maxvalue': 0,
            'mintime': (0, 0, 0, 0, 0, 0),
            'minvalue': 0,
            'avgcoast': 0
    }
    data['maxvalue'] = max_coast
    data['minvalue'] = min_coast
    data['mintime']  = min_time
    data['maxtime']  = max_time
    data['avgcoast'] = avg

    return data

def parse_file_lecture(datafile):
    workbook = xlrd.open_workbook(datafile)
    sheet       = workbook.sheet_by_index(0)

    data = [[sheet.cell_value(r, col) for col in range(sheet.ncols)] for r in range(sheet.nrows)]

    cv = sheet.col_values(1, startrowx = 1, endrowx = None)

    maxval = max(cv)
    minval = min(cv)

    maxpos = cv.index(maxval) + 1
    minpos = cv.index(minval) + 1

    maxtime = sheet.cell_value(maxpos, 0)
    mintime = sheet.cell_value(minpos, 0)
    realmaxtime = xlrd.xldate_as_tuple(maxtime, 0)
    realmintime = xlrd.xldate_as_tuple(mintime, 0)
    avg = sum(cv) / float(len(cv))

    data = {
            'maxtime': realmaxtime,
            'maxvalue': maxval,
            'mintime': realmintime,
            'minvalue': minval,
            'avgcoast':   avg
    }
    return data

def test():
    open_zip(dataZipFile)
    data = parse_file(datafile)

    assert data['maxtime'] == (2013, 8, 13, 17, 0, 0)
    assert round(data['maxvalue'], 10) == round(18779.02551, 10)

test()

