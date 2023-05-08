import functions

path = input("What's the name of the spreadsheet: ") + ".xlsx"
sheet = functions.getWorkbook(path)
print(functions.getAverageForColumn(sheet, 15)) 

# Loop through each column and find average of each
# Get average of overall evaluation