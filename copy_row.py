import openpyxl

# Open the workbook
wb = openpyxl.load_workbook('example.xlsx')

# Select the worksheet
ws = wb.active

# Print the row
for cell in ws[2]:
    print(cell.value)