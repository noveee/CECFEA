import functions

path = input("What's the name of the spreadsheet: ") + ".xlsx"
sheet = functions.getWorkbook(path)

for i in range(11, 19):
    print(functions.getAverageForColumn(sheet, i)) 

# Loop through each column and find average of each
# Get average of overall evaluation