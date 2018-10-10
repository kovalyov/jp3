import openpyxl

print("Введите название файла :", )
# file_name = input() + ".xlsx"
file_name = 'test.xlsx'

# Load in the workbook
wb = openpyxl.load_workbook(file_name)

# Get sheet names
print(wb.sheetnames)

# sheet = wb.get_sheet_by_name('Sheet3')    #DeprecationWarning: Call to deprecated function get_sheet_by_name (Use wb[sheetname]).
# sheet = wb.active # Choose last active sheet before closing a file
sheet = wb['Sheet3']
print(str(sheet.title))

ch = 'A'
row = 1
while chr(ord(ch)) <= "H":      # Limitation a range of columns
    print("Column Symbol:", ch)
    cell = sheet[str(ch) + str(row)]
    print("Value: ", cell.value)
    # print("Row: ",cell.row)
    # print("Column: ",cell.column)
    print("Coordinate: ", cell.coordinate, "\n==================|")
    ch = chr(ord(ch) + 1)
    row += 1
