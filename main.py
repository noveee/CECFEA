import functions

path = "TestSheet.xlsx"
sheet = functions.getWorkbook(path)
print(functions.getCorrespondingRows(sheet, 3, "idk")) 