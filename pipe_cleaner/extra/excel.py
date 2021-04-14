from xlrd import open_workbook
from simplejson import dump


# Opens excel file based on path and opens excel sheet based on integer
description_sheet = open_workbook(r'C:\Users\joe.ton\Documents\trr_79000.xlsx').sheet_by_index(0)
component_sheet = open_workbook(r'C:\Users\joe.ton\Documents\trr_79000.xlsx').sheet_by_index(1)

excel_data = {}

# Imports excel data from excel file to ado dictionary
def read_excel_sheets():
    for description_cell in range(description_sheet.nrows):
        key = str(description_sheet.cell(description_cell, 0))
        if 'text' in key:
            key = key[6:-1:]
        if 'number' in key:
            key = key[7:-2:]
        value = str(description_sheet.cell(description_cell, 1))
        if 'text' in value:
            value = value[6:-1:]
        if 'number' in value:
            value = value[7:-2:]
        excel_data.update({key: value})
    for component_cell in range(component_sheet.nrows):
        key = str(component_sheet.cell(component_cell, 0))
        if 'text' in key:
            key = key[6:-1:]
        if 'number' in key:
            key = key[7:-2:]
        value = str(component_sheet.cell(component_cell, 1))
        if 'text' in value:
            value = value[6:-1:]
        if 'number' in key:
            value = value[7:-2:]
        excel_data.update({key: value})

read_excel_sheets()

test = excel_data.get['TRR']

with open("filename.json", "w") as f:
    f.write(dumps(excel_data, indent=4))