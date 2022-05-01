from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.utils  import get_column_letter 
from typing import List
# import time

wb = load_workbook(filename = 'V2.xlsx', data_only = True )
# work_book_destination = load_workbook(filename = 'book12.xlsx')
# wa = work_book_destination.active
ws = wb.active 

# csv to xlsx

template = load_workbook("test1.xlsx") #Add file name
temp_sheet = template.active #Add Sheet name
# temp_sheet = template.worksheets[0]



def copyRange(startCol, startRow, endCol, endRow):
    rangeSelected = []
    #Loops through selected Rows
    for i in range(startRow,endRow + 1,1):
        #Appends the row to a RowSelected list
        rowSelected = []
        for j in range(startCol,endCol+1,1):
            rowSelected.append(ws.cell(row = i, column = j).value)
        #Adds the RowSelected List and nests inside the rangeSelected
        rangeSelected.append(rowSelected)
 
    return rangeSelected


def pasteRange(startCol, startRow, endCol, endRow,copiedData):
    countRow = 0
    for i in range(startRow,endRow+1,1):
        countCol = 0
        for j in range(startCol,endCol+1,1):
            
            temp_sheet.cell(row = i, column = j).value = copiedData[countRow][countCol]
            countCol += 1
        countRow += 1


def createData():
    print("Processing...")
    selectedRange = copyRange(2,38,7,67)
    pastingRange = pasteRange(3,154,8,183,selectedRange)
    template.save("test1.xlsx")
    print("Range copied and pasted!")

createData()


def fajr_salah():
    print("Processing_fajr_salah...")
    selectedRange = copyRange(13,38,13,67)
    pastingRange = pasteRange(9,154,9,183,selectedRange)
    template.save("test1.xlsx")
    print("Range copied and pasted!")


fajr_salah()


                    
def magrib_salah():
    print("Processing_magrib_salah...")
    selectedRange = copyRange(16,38,16,67)
    pastingRange = pasteRange(12,154,12,183,selectedRange)
    template.save("test1.xlsx")
    print("Range copied and pasted!")

magrib_salah()




















