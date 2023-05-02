import openpyxl

wb = openpyxl.Workbook()
ws = wb.active
print(type(ws))

path = "TestSheet.xlsx"
wb_obj = openpyxl.load_workbook(path)
print(type(wb_obj))