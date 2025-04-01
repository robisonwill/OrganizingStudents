import openpyxl
from openpyxl import Workbook

importedWorkbook = openpyxl.load_workbook("Poorly_Organized_Data_1.xlsx")

currSheet = importedWorkbook.active

createWorkbook = Workbook()

newSheets = []
for row in currSheet.iter_rows(min_row= 2, min_col= 1, max_col= 1) :
    for cell in row :
        if cell.value not in newSheets :
            createWorkbook.create_sheet(f"{cell.value}")
            newSheets.append(cell.value)
        else:
            pass

createWorkbook.save(filename="organizedData.xlsx")

createWorkbook.close()