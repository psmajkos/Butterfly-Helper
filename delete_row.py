import openpyxl

# Open the workbook
wb = openpyxl.load_workbook('example.xlsx')

# Select the worksheet
ws = wb.active

# Delete the row
ws.delete_rows(2)

# Save the changes
wb.save('example.xlsx')
