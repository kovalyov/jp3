import datetime
from openpyxl import load_workbook

time = datetime.datetime.now().strftime("_%I_%M_%S_%B_%d_%Y")
print("Введите название файла :\n", "Пример - C:/Users/%USERNAME%/PycharmProjects/untitled/first_file.xlsx")
# file_name = input() + ".xlsx"
file_name = 'first_file.xlsx'

# Load in the workbook
wb = load_workbook(file_name)

# Get sheet names
print(wb.sheetnames)

sheet = wb['Sheet1']
print("Работаем во вкладке: ", str(sheet.title))

ch = 'A'
row = 1

column_1 = 0
column_2 = 0
column_3 = 0
while chr(ord(ch)) <= "Z":  # Limitation a range of columns
    # print("Column Symbol:", ch)

    cell = sheet[str(ch) + str(row)]
    # print("Value: ", cell.value)
    if cell.value == "FirstValue":
        column_1 = cell.column
    if cell.value == "SecondValue":
        column_2 = cell.column
    if cell.value == "Total":
        column_3 = cell.column
    # print("Row: ",cell.row)
    # print("Column: ",cell.column)
    # print("Coordinate: ", cell.coordinate, "\n==================|")
    ch = chr(ord(ch) + 1)
    # row += 1

print("First: ", str(column_1))
print("Second: ", str(column_2))
print("Third: ", str(column_3))

while sheet[str(column_1) + str(row + 1)].value is not None:
    a = sheet[str(column_1) + str(row + 1)].value
    b = sheet[str(column_2) + str(row + 1)].value
    res = a + b
    sheet[str(column_3) + str(row + 1)] = res
    print(res)
    row += 1

wb.save("reports/first_file" + time + ".xlsx")
