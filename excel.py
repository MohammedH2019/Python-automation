from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.utils  import get_column_letter 
from typing import List
# import time

wb = load_workbook(filename = '2021 05-06_v2.xlsx')
# work_book_destination = load_workbook(filename = 'book12.xlsx')
# wa = work_book_destination.active
ws = wb.active 

# csv to xlsx

template = load_workbook("MyMasjid3-Timetable.xlsx") #Add file name
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
    selectedRange = copyRange(2,4,6,34)
    pastingRange = pasteRange(3,123,7,153,selectedRange)
    template.save("MyMasjid3-Timetable.xlsx")
    print("Range copied and pasted!")

createData()








                    



























