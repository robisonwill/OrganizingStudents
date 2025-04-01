from openpyxl import Workbook
from openpyxl import Poorly_Organized_Data_1


'''
for row in currentWs.iter_rows(min_row= 1, max_row= 10, min_col= 1, max_col= 1) :
    for cell in row :
        if cell.value not in newSheets :
            myWorkbook.create_sheet(f"{cell.value}")
            newSheets.append(cell.value)
'''