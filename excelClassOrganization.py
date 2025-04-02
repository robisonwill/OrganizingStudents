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



studentList = []
for row in currSheet.iter_rows(min_row=2, values_only=True):
    studentList.append(row)

for student in studentList:
    currSheet = createWorkbook[student[0]]
    student_info = student[1].split("_")
    student_info.append(student[2])
    currSheet.append(student_info)




createWorkbook.close()
createWorkbook.save(filename="organizedData.xlsx")