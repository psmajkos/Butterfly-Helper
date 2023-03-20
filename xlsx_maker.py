import openpyxl

# Create a new workbook
workbook = openpyxl.Workbook()

# Get the active worksheet
worksheet = workbook.active

# Write data to the worksheet
worksheet['A1'] = 'ZP'
worksheet['B1'] = 'Order'
worksheet['C1'] = 'Quantity' 

worksheet['A2'] = 'ZP23063611'
worksheet['B2'] = 'CO3562774'
worksheet['C2'] = 1

worksheet['A3'] = 'ZP23076514'
worksheet['B3'] = 'CO7865276'
worksheet['C3'] = 2

worksheet['A4'] = 'ZP23025612'
worksheet['C4'] = 3

worksheet['A5'] = 'ZP2305676'
worksheet['C5'] = 4

worksheet['A6'] = 'ZP23023764'
worksheet['C6'] = 5

worksheet['A7'] = 'ZP23063611'
worksheet['C7'] = 6

worksheet['A8'] = 'ZP23076514'
worksheet['C8'] = 7

worksheet['A9'] = 'ZP23025612'
worksheet['C9'] = 8

worksheet['A10'] = 'ZP2305676'
worksheet['C10'] = 9

worksheet['A11'] = 'ZP23023764'
worksheet['B11'] = 'CO7647263'
worksheet['C11'] = 10

worksheet['A12'] = 'ZP23023764'
worksheet['C12'] = 11

worksheet['A13'] = 'ZP23023764'
worksheet['C13'] = 11

worksheet['A14'] = 'ZP23023764'
worksheet['C14'] = 12

worksheet['A15'] = 'ZP23023764'
worksheet['C15'] = 13

worksheet['A16'] = 'ZP23023764'
worksheet['C16'] = 14

worksheet['A17'] = 'ZP23023764'
worksheet['C17'] = 15

worksheet['A18'] = 'ZP23023764'
worksheet['C18'] = 16

worksheet['A19'] = 'ZP23023764'
worksheet['C19'] = 17



# Save the workbook
workbook.save('example.xlsx')
