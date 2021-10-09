import openpyxl as xl

wb = xl.load_workbook('transaction.xlsx') # Creates a variable 'wb' and stores the target workbook inside
sheet = wb['Sheet1'] # Acesses the specific sheet that contains the targeted data

for row in range(2, sheet.max_row + 1): # Loops through rows 2 to 4 inclusive
    cell = sheet.cell(row, 3) # Creates a variable 'cell' and stores the data of that cell from the spreadsheet inside
    corrected_price = cell.value * 0.9 # Creates a variable 'corrected_price' and stores the value of the current cell after a 10% discount is applied
    corrected_price_cell = sheet.cell(row, 4) # Creates a variable 'corrected_price_cell' and sets it as a cell in a new column to store the corrected prices
    corrected_price_cell.value = corrected_price # Sets the value of 'corrected_price_cell' to the corrected_price variable
    
    
wb.save('transactions2.xlsx') # Saves the changes to a new spreadsheet to avoid overriding the original file as a failsafe in case bugs exist
