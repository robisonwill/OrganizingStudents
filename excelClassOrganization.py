import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font

importedWorkbook = openpyxl.load_workbook("Poorly_Organized_Data_1.xlsx")

def summaryInfo(ws):
    ws["F1"] = "Summary Stastics"
    bold_font = Font(bold=True)
    ws["F1"].font = bold_font

    ws["G1"] = "Value"
    bold_font = Font(bold=True)
    ws["G1"].font = bold_font

    ws["F2"] = "Highest Grade"
    ws["G2"] = "MAX(D:D)"

    ws["F2"] = "Highest Grade"
    ws["G2"] = "MAX(D:D)"

    ws["F3"] = "Lowest Grade"
    ws["G3"] = "MIN(D:D)"

    ws["F4"] = "Mean Grade"
    ws["G4"] = "MEAN(D:D)"

    ws["F5"] = "Median Grade"
    ws["G5"] = "MEDIAN(D:D)"

    ws["F6"] = "Number of Students"
    ws["G6"] = '=COUNT(A:A)'

currSheet = importedWorkbook.active

createWorkbook = Workbook()

newSheets = []
for row in currSheet.iter_rows(min_row= 2, min_col= 1, max_col= 1) :
    for cell in row :
        if cell.value not in newSheets :
            createWorkbook.create_sheet(f"{cell.value}")
            summaryInfo(currSheet)
            newSheets.append(cell.value)
        else:
            pass

createWorkbook.save(filename="organizedData.xlsx")

createWorkbook.close()