import openpyxl

# Load the workbook
workbook = openpyxl.load_workbook('example.xlsx')

# Select the sheet you want to add a record to
sheet = workbook['Sheet']


# Add a new record to the sheet
new_record = ['ZP23069716', 'CO2303421', 5]
sheet.append(new_record)

# Save the workbook
workbook.save('example.xlsx')
