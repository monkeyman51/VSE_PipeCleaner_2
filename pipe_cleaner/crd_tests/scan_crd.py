import pandas as pd
import csv

crd_list = ['CRD_1',
            'CRD_2',
            'CRD_3',
            'CRD_4']


# def get_csv_from_crd(crd_number):
#     """
#
#     :param crd_number:
#     :return:
#     """
#     df = pd.read_excel(f'cover_page.xlsx', sheet_name=None)
#     df['FW-SW Configuration'].to_csv('cover_page.csv', sep=',')
#
#     with open('cover_page.csv', 'r', newline='', encoding="utf-8") as f:
#         reader = csv.reader(f)
#         for row in reader:
#             for item in row:
#                 if 'CTD' in item:
#                     print(item)

def read_csv(file_name: str, column: int) -> list:
    """
    Reads csv with file name and selected column
    :param file_name:
    :param column: iterates based on column number, zero-indexed
    :return:
    """
    import csv

    column_data: list = []
    with open(f'{file_name}.csv', 'r') as csv_file:

        # Interprets csv into a "reader"
        # Used for iterating later to store into a list
        reader = csv.reader(csv_file, delimiter=',')
        for row in reader:

            # foo = row[column]

            # Stores each row data in column into list
            column_data.append(row[column])

    return column_data

import time

start = time.time()
foo = read_csv('Monitor-Sensors_VSE0G6IWBAL-025', 97)
bar = sorted(list(set(foo)))
print(bar[-1])
end = time.time()
print(end - start)