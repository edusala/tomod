
# importing openpyxl module
import openpyxl
 
# Give the location of the file
path = "minta.xlsx"
 
# workbook object is created
wb_obj = openpyxl.load_workbook(path)
sheet_obj = wb_obj.active
max_col = sheet_obj.max_column
max_row = sheet_obj.max_row

def getFirstName(name):    
    (secondName,firstName) = name.split()
    return firstName

def getSecondName(name):    
    (secondName,firstName) = name.split()
    return secondName
def getUserName(name):    
    (secondName,firstName) = name.split()
    return secondName + "." + firstName

def convert(name):    
    (secondName,firstName) = name.split()
    return secondName + "." + firstName

# Loop will print all columns name
for i in range(1, max_row + 1):
    for j in range(1, max_col + 1):
        cell_obj = sheet_obj.cell(row = i, column = j)
        if j == 2:
            fullName = cell_obj.value
            getFirstName(fullName)
            getSecondName(fullName)
            print(getUserName(fullName))




