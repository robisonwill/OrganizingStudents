from openpyxl import Workbook

import random

myWorkbook = Workbook()
# set active sheet
currentWs = myWorkbook.active
# set title for current worksheet
currentWs.title = "Customers"

cols = ["A", "B", "C", "D", "E", "F"]

for iRow in range(1,11) :
    for iCol in range(6) :
        currentWs[ cols[iCol] + str(iRow)] = random.randint(0,100)
''' This is another way to do it that is simpler for some people:
for row in currentWs.iter_rows(min_row= 1, max_row= 10, min_col= 1, max_col= 6) :
    for cell in row :
        cell.value = random.randint(0,100)
'''
newSheets = []

for row in currentWs.iter_rows(min_row= 1, max_row= 10, min_col= 1, max_col= 1) :
    for cell in row :
        if cell.value not in newSheets :
            myWorkbook.create_sheet(f"{cell.value}")
            newSheets.append(cell.value)



currentWs["A12"] = "SUM:"
currentWs["A13"] = "=SUM(A1:A10)"

# B min C max D count E Avg F Countif
currentWs["B12"] = "MIN:"
currentWs["B13"] = "=MIN(B1:B10)"

currentWs["C12"] = "MAX:"
currentWs["C13"] = "=MAX(C1:C10)"

currentWs["D12"] = "COUNT:"
currentWs["D13"] = "=COUNT(D1:D10)"

currentWs["E12"] = "AVG:"
currentWs["E13"] = "=AVERAGE(E1:E10)"

currentWs["F12"] = "COUNTIF:"
currentWs["F13"] = '=COUNTIF(F1:F10, ">50")'

myWorkbook.save(filename="firstexcel.xlsx")

myWorkbook.close()