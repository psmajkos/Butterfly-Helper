import openpyxl

# Create a new workbook
workbook = openpyxl.Workbook()

# Create a new sheet
new_sheet = workbook.create_sheet("Sheet")

# Save the workbook
workbook.save('done.xlsx')
