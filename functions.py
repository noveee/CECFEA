import openpyxl

def getCorrectPath():
    '''
    Assist the user in getting the right file path
    '''

    better_path = input("Enter a new path to xlsx workbook: ")
    return better_path

def getWorkbook(path: str):
    '''
    Verifies that the given path works and creates a workbook object
    Then creates a sheet object which is used throughout the whole process

    :param path: Path to xlsx workbook
    '''

    # Catches any issues related to the path and attempts to handle them before moving on
    # Instead of looping endlessly, it gives the user a chance to adjust the path of the workbook
    try:
        wb_obj = openpyxl.load_workbook(path)

    except FileNotFoundError:
        print("\nThe path is incorrect or the file does not exist")
        new_path = getCorrectPath()
        wb_obj = openpyxl.load_workbook(new_path)

    except:
        print("\nThere was an issue with the given path")
        new_path = getCorrectPath()
        wb_obj = openpyxl.load_workbook(new_path)

    # Returns a working sheet for the program to use
    sheet = wb_obj.active
    return sheet

def workbookGrid(sheet: openpyxl.workbook.workbook.Workbook):
    '''
    Creates a grid of the current sheet for use
    Using a nested list

    :param sheet: The sheet obj to create a grid for
    '''


    # Each list corresponds to a ROW
    # Each value on the row corresponds to a COLUMN
 
    tot_row = sheet.max_row
    tot_column = sheet.max_column 
    sheetGrid = []


    # Nested list comprehension
    # The first row/list corresponds to the questions
    # The following rows/list corresponds to the responses

    for list in range(tot_row + 1):
        sheetGrid.append([])
        for item in range(tot_column + 1):
            print(item)
            sheetGrid[list].append(sheet.cell(row = list + 1, column = item + 1).value)

    return sheetGrid

def getColumnInfo(col: int, ):
    '''
    Get all the values from a specific column

    :param col: Column to get values from
    ''' 

def indexRow(search):
    '''
    Get all the values from a specifc row and the information it corresponds to

    :param search: Value used to index a specific row
    '''

'''
Project specific functions 
Indexing based on student ID
Indexing based on Class Number
Averging rating values
Compiling comments into one space
Outputting into specified format (xlsx, sql, etc)
'''

path = "TestSheet.xlsx"
sheet = getWorkbook(path)
print(workbookGrid(sheet))