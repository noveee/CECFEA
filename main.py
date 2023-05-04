import functions

path = "TestSheet.xlsx"
sheet = functions.getWorkbook(path)
print(functions.getAverageForColumn(sheet, 2)) 