import datetime
from openpyxl import load_workbook
import re
import sys

#************************* Load XLSX file for analyze **************************************#

time = datetime.datetime.now().strftime("_%I_%M_%S_%B_%d_%Y")
# print("Введите название файла :\n", "Пример - C:/Users/%USERNAME%/PycharmProjects/untitled/first_file.xlsx")
# file_name = input() + ".xlsx"
file_name = 'C:/Users/ohrabar/Desktop/SiteTimeDetailsReport.xlsx'

#************************* Load TXT file with EmployeesFileName ****************************#
employeesFileName = 'C:/Users/ohrabar/Desktop/names.txt'
file = open(employeesFileName, 'r')
s = file.read()
employees = re.split(r', ', s)
file.close()

#************************** Load in the workbook ********************************************#
wb = load_workbook(file_name)
# Get sheet names
print(wb.sheetnames)

sheet = wb['Site Time Details Report']
print("Works in sheet: ", str(sheet.title))

ch = 'A'
row = 1

sheet['T1'] = "Total Time"      # создаем ячейку "Total Time" в которую будем вносить значения общего времени

EmployeeColumn = 0
DateEntryColumn = 0
EmployeeTimeOnTaskColumn = 0
TotalTimeColumn = 0
rows = {}

while chr(ord(ch)) <= "Z":      # Limitation a range of columns
    # print("Column Symbol:", ch)

    cell = sheet[str(ch) + str(row)]
    # print("Value: ", cell.value)
    if cell.value == "Employee":
        EmployeeColumn = cell.column
    if cell.value == "Entry Date":
        DateEntryColumn = cell.column
    if cell.value == "Employee Time On Task":
        EmployeeTimeOnTaskColumn = cell.column
    if cell.value == "Total Time":
        TotalTimeColumn = cell.column
    # print("Row: ",cell.row)
    # print("Column: ",cell.column)
    # print("Coordinate: ", cell.coordinate, "\n==================|")
    ch = chr(ord(ch) + 1)
    # row += 1

# print("First_Employee           : ", str(EmployeeColumn))
# print("Second_EntryDate         : ", str(DateEntryColumn))
# print("Third EmployeeTimeOnTask : ", str(EmployeeTimeOnTaskColumn))
# print("Total Time               : ", str(TotalTimeColumn))

for i in employees:
    print("\n")
    print(i)
##################################################################
    total_hours = 0
    empty_dates = []
    row_2 = 4
    # print(sheet[str(column_1) + str(row_2)].value)
    while sheet[str(EmployeeColumn) + str(row_2)].value is not None:
        # print(i, '/n')
        # print(sheet[str(column_1) + str(row_2)].value)
        if sheet[str(EmployeeColumn) + str(row_2)].value == i:
            #print(str(sheet[str(column_3) + str(row + 3)].value))
            if sheet[str(EmployeeTimeOnTaskColumn) + str(row_2)].value is None:
                empty_dates.append(str(sheet[str(DateEntryColumn) + str(row_2)].value))
            else:
                total_hours += int(sheet[str(EmployeeTimeOnTaskColumn) + str(row_2)].value)
        row_2 += 1
    print("Quantity of hours: ", total_hours)
    if len(empty_dates) != 0:
       print("Empty dates: ", empty_dates)




    # if not list(empty_dates):
    #     if list(empty_dates):
    #         print("OK")
#################################################################

# while sheet[str(EmployeeColumn) + str(row + 1)].value is not None:
#     a = sheet[str(EmployeeColumn) + str(row + 1)].value
#     b = sheet[str(DateEntryColumn) + str(row + 1)].value
#     res = a + b
#     sheet[str(TotalTimeColumn) + str(row + 1)] = res
#     print(res)
#     row += 1
#

#wb.save("reports/template" + time + ".xlsx")

