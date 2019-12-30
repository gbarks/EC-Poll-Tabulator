#!/usr/bin/env python3

# =========================================================================
#  ElloCoaster reader for detailed blank ballots and regular voter ballots
#  Author: Grant Barker
#
#  Requires Python 3
# =========================================================================

from coaster import Coaster
from openpyxl import load_workbook

def blank_to_none(string):
    if string:
        return string
    else:
        return None

def read_detailed_ballot(filepath, id_col=7, name_col=2, park_col=4):
    coasters = {}
    wb = load_workbook(filepath)
    ws = wb.worksheets[0]
    for row in ws.iter_rows(min_row=2): 
        rcid = row[id_col].value
        url = None
        if "HYPERLINK" in row[id_col].value:
            rcid = row[id_col].value[:-2].split('"')[-1]
            url = row[id_col].value.split('"')[1]

        name = row[name_col].value
        park = row[park_col].value

        # silly exception for Gravity Group coasters listed under M&V on RCDB
        designer = blank_to_none(row[16].value)
        if designer == "Martin & Vleminckx":
            if name != "Coastersaurus" and park != "Legoland Florida":
                designer = "The Gravity Group, LLC"

        c = Coaster(rcid,                         #rcid
                    name,                         #name
                    park,                         #park
                    url,                          #url
                    blank_to_none(row[3].value),  #altname
                    blank_to_none(row[5].value),  #country
                    blank_to_none(row[6].value),  #fullcity
                    blank_to_none(row[8].value),  #location
                    blank_to_none(row[9].value),  #state
                    blank_to_none(row[10].value), #city
                    blank_to_none(row[11].value), #status
                    blank_to_none(row[12].value), #opendate
                    blank_to_none(row[13].value), #closedate
                    blank_to_none(row[14].value), #rctype
                    blank_to_none(row[15].value), #scale
                    designer                    , #make
                    blank_to_none(row[17].value), #model
                    blank_to_none(row[18].value), #submodel
                    blank_to_none(row[19].value)) #tracks
        coasters[rcid] = c

    return coasters

### ========================================================================
###  xls support gracefully copied from the helpful users at StackOverflow
###  https://stackoverflow.com/questions/9918646/how-to-convert-xls-to-xlsx
### ========================================================================
import xlrd
from openpyxl.workbook import Workbook
from openpyxl.reader.excel import load_workbook, InvalidFileException

def open_xls_as_xlsx(filename):
    # first open using xlrd
    book = xlrd.open_workbook(filename)
    index = 0
    nrows, ncols = 0, 0
    while nrows * ncols == 0:
        sheet = book.sheet_by_index(index)
        nrows = sheet.nrows
        ncols = sheet.ncols
        index += 1

    # prepare a xlsx sheet
    book1 = Workbook()
    sheet1 = book1.get_active_sheet()

    for row in xrange(0, nrows):
        for col in xrange(0, ncols):
            sheet1.cell(row=row, column=col).value = sheet.cell_value(row, col)

    return book1
### ========================================================================

def read_voter_ballot(filepath):

