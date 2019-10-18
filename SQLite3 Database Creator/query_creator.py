try:
    import xlrd
except ModuleNotFoundError:
    import os
    os.system("python3 -m pip install xlrd")
    import xlrd

file_path = "excel_name.xlsx" # path to excel .xlsx file
directory = "directory_name" # django directory which contains model.py

"""
formating of data types:
abs -> absolute (removes quotes)
date -> Date time object (cell should have "dd/mm/yyy")
int -> Integer
"""

workbook = xlrd.open_workbook(file_path)
number_of_sheets = len(workbook.sheet_names())
print("from {}.models import *".format(directory))
for sheet_no in range(number_of_sheets):
    sheet = workbook.sheet_by_index(sheet_no)
    table_name = sheet.cell_value(0, 0)
    columns = 0
    while True:
        try:
            if len(str(sheet.cell_value(1, columns+1))) == 0:
                break
            else:
                columns += 1
        except:
            columns += 1
            break
    rows = 0
    while True:
        try:
            if len(str(sheet.cell_value(3+rows+1, 0))) == 0:
                break
            else:
                rows += 1
        except:
            rows += 1
            break
    column_names = []
    column_data_types = []
    for col in range(columns):
        column_names.append(str(sheet.cell_value(1, col)))
        column_data_types.append(str(sheet.cell_value(2, col)))
    for row in range(3, 3+rows):
        row_data = []
        query_data = []
        for col in range(columns):
            row_data.append(str(sheet.cell_value(row, col)))
        for (ind, value) in enumerate(row_data):
            if len(value) == 0:
                continue
            elif column_data_types[ind].lower() == 'abs':
                value = value[1:-1]
            elif column_data_types[ind].lower() == 'date':
                value = "-".join((value[1:-1]).split('/')[::-1])
            elif column_data_types[ind].lower() == 'int':
                value = int(eval(value))
                query = "{}={}".format(column_names[ind], value)
                query_data.append(query)
                continue
            query = "{}='{}'".format(column_names[ind], value)
            query_data.append(query)
        print("db_query = {}({})".format(table_name, ", ".join(query_data)))
        print("db_query.save()")
              
                


    
