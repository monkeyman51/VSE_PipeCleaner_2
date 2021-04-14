from xlrd import open_workbook
from simplejson import dumps


skudoc = open_workbook(r'C:\Users\joe.ton\Documents\skudoc.xlsx').sheet_by_index(0)

start_data = {}
more_data = {}


def start():
    with open('skudoc.json', 'w') as file:
        x = 1
        while x < skudoc.ncols:
            for configuration in range(skudoc.nrows):
                key = str(skudoc.cell(configuration, 0))
                key = unnecessary_string(key)
                value = str(skudoc.cell(configuration, x))
                value = unnecessary_string(value)
                start_data.update({key: value})
            with open("skudoc.json", "a") as f:
                for description in range(1, skudoc.ncols):
                    value = str(skudoc.cell(x, description))
                    print(value)
                    # f.write(str(description))
                    f.write(dumps(start_data, indent=4))
            x += 1
        file.close()

def unnecessary_string(string):
    if 'text:' in string:
        string = str(string)[6:-1:]
    if 'number:' in string:
        string = string[7:-2:]
    if 'empty:' in string:
        string = string[6:-2:]
    return string


start()