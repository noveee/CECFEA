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

    :param sheet: The active sheet obj to create a grid for
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

def getColumnInfo(sheet: openpyxl.workbook.workbook.Workbook, col: int, ):
    '''
    Get all the values from a specific column 
    And returns a list with those values 

    :param sheet: The active sheet obj to get info from 
    :param col: Column to get values from 
    ''' 

    values = []

    # Looping through the column and appending each value after the column name to a list
    for i in range(sheet.max_row):

        # Skips the first iteration and prints the column name
        if i == 0:
            print(f"Column: {sheet.cell(row = 1, column = col).value}")
            continue
        
        # Prints the current value and appends to list
        print(sheet.cell(row = i + 1, column = col).value)
        values.append(sheet.cell(row = i + 1, column = col).value)

    return values



def indexRow(sheet: openpyxl.workbook.workbook.Workbook, search):
    '''
    Get all the values from a specifc row and the information it corresponds to

    :param sheet: The active sheet obj to get info from 
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

def getAverageForColumn(sheet: openpyxl.workbook.workbook.Workbook, col: int):
    '''
    Returns the average of the values in the given column

    :param sheet: The active sheet obj to get info from 
    :param col: Column to calculate the average values from
    '''

    values = getColumnInfo(sheet, col)
    total = 0
    count = 0

    num_or_letter = int(input("Enter 1 for number based averaging or 2 for letter based averaging: "))
    
    # For number based averaging
    if num_or_letter == 1:
        for item in values:
            total += item
            count += 1

        average = total/count

        return average
    
    # For string based averaging
    elif num_or_letter == 2:
        criteria = [input("Enter the grading criteria from the highest rating to the lowest, seperated by space\n(I.E. : Excellent Good Poor): ")]
        grading_scale = []

        # Filling an empty list for grading
        # Will compress this algorithm later
        for grade in criteria:
            grading_scale.append(0)

        all_grades = getColumnInfo(sheet, col)

        for grade in all_grades:
            # pause on this...
        
            return None

    # If the number given is not an option
    else:
        print("Not an option")

def splitSheet(sheet: openpyxl.workbook.workbook.Workbook):
    


path = "TestSheet.xlsx"
sheet = getWorkbook(path)
print(getAverageForColumn(sheet, 2)) 