from openpyxl import load_workbook

import os


def enter_amount_inventory() -> str:
    """
    Enter the amount for the inventory transaction.
    :return:
    """
    print(f'\n\tEnter the number of parts being scanned and press enter...')

    user_input: str = input(f'\tAmount: ')

    if user_input.isdigit():
        return user_input
    else:
        print(f'\tWARNING!!!!!')
        print(f'\tMust put numbers only.')
        enter_amount_inventory()


def main_method(default_user_name: str) -> None:
    """
    Record serial numbers for inventory
    :return:
    """
    # os.system(fr'start EXCEL.EXE record_serial_numbers.xlsx')
    workbook = load_workbook('settings/record_serial_numbers.xlsx')
    worksheet = workbook['Sheet1']

    worksheet[f'E2'] = int(enter_amount_inventory())

    try:
        workbook.save('record_serial_numbers.xlsx')
        os.system(fr'start EXCEL.EXE record_serial_numbers.xlsx')
    except PermissionError:
        print(f'\tWARNING!!!!\n')
        print(f'\trecord_serial_numbers.xlsx file is already open.')
        print(f'\tPlease close and try again.')
        print(f'\tPress enter to quit.')
        input()
