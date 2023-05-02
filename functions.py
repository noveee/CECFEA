import openpyxl

def getCorrectPath():
    '''
    Assist the user in getting the right file path
    '''

    better_path = input("Enter a new path to xlsx workbook: ")
    return better_path

def getWorkbook(path: str):
    '''
    Verifies that the given path opens a workbook
    Then returns the workbook object for usage

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
    return wb_obj

getWorkbook(input("Workbook path:"))