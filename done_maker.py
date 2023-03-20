import openpyxl

# Create a new workbook
workbook = openpyxl.Workbook()

# Get the active worksheet
worksheet = workbook.active

# # Write data to the worksheet
# worksheet['A1'] = 'ZP'
# worksheet['B1'] = 'Order'
# worksheet['C1'] = 'Quantity' 

# Save the workbook
workbook.save('done.xlsx')