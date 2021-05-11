import shutil
import os
import sys

from time import strftime
from openpyxl import load_workbook
from openpyxl.styles import Alignment

from colorama import Fore, Style


def get_file_path(user_name: str, root_path: str, file_number: str) -> str:
    """
    Get file name dedicated towards documenting / logging data entry for inventory.  Plan to have file names
    represented as name_date_time
    :return:
    """
    return fr'{root_path}\{file_number}_{user_name}.xlsx'


def get_date_file_name(default_user_name):
    clean_name: str = default_user_name. \
        replace('.', ''). \
        lower()

    year: str = strftime('%Y')[-2:]
    current_date: str = strftime('%m%d')
    current_time: str = strftime('%I%M%S')
    return f'{current_date}{year}_{current_time}_{clean_name}.xlsx'


def get_backup_path(default_user_name: str) -> str:
    """
    Have backup file path in case original file path is corrupt.
    """
    file_name: str = get_date_file_name(default_user_name)
    current_date: str = strftime('%m%d')
    current_time: str = strftime('%I%M%S')
    granular_file_name: str = f'{current_date}_{current_time}_{file_name}.xlsx'

    return fr'Z:\Kirkland_Lab\Users\joe_ton\project_inventory\back_transaction_logs\{granular_file_name}'


def get_clean_user_name(default_user_name):
    if '.' not in default_user_name:
        indexed_clean_name: list = []

        for index, character in enumerate(default_user_name, start=0):
            index_character: str = default_user_name[index]

            if index_character.isupper() and not index == 0:
                indexed_clean_name.append(' ')
                indexed_clean_name.append(index_character)
            else:
                indexed_clean_name.append(index_character)

        return ''.join(indexed_clean_name).replace('-Ext', '').replace(' ', '').replace('_', '').lower()
    else:

        return default_user_name.replace('.', ' ').title().replace('-Ext', '').replace(' ', '').replace('_', '').lower()


def get_file_number(root_file_path: str) -> str:
    """
    Get file names for providing new unique name.
    :param root_file_path:
    :return:
    """
    try:
        raw_file_names: list = os.listdir(root_file_path)

        if len(raw_file_names) == 0:
            return '00001'
        else:
            raw_numbers: list = []
            for raw_file_name in raw_file_names:
                raw_numbers.append(raw_file_name[0:5])

            non_zero_names: list = []
            for index, raw_number in enumerate(raw_numbers, start=0):
                non_zero_characters: list = []

                flip: int = 0
                for character in raw_number:
                    if character != 0 and flip == 0:
                        non_zero_characters.append(character)
                        flip += 1
                    elif character != 0 and flip >= 1:
                        non_zero_characters.append(character)

                complete_name: str = ''.join(non_zero_characters)
                try:
                    non_zero_names.append(int(complete_name))
                except ValueError:
                    pass

            clean_non_zero_names = sorted(list(set(non_zero_names)))
            latest_file_number = int(clean_non_zero_names[-1])
            latest_file_number += 1

        return get_file_name_number(latest_file_number)

    except FileNotFoundError:

        alt_path = r'\\172.30.1.100\pxe\Kirkland_Lab\PipeCleaner\transaction_logs'
        raw_file_names: list = os.listdir(alt_path)

        if len(raw_file_names) == 0:
            return '00001'
        else:
            raw_numbers: list = []
            for raw_file_name in raw_file_names:
                raw_numbers.append(raw_file_name[0:5])

            non_zero_names: list = []
            for index, raw_number in enumerate(raw_numbers, start=0):
                non_zero_characters: list = []

                flip: int = 0
                for character in raw_number:
                    if character != 0 and flip == 0:
                        non_zero_characters.append(character)
                        flip += 1
                    elif character != 0 and flip >= 1:
                        non_zero_characters.append(character)

                complete_name: str = ''.join(non_zero_characters)
                non_zero_names.append(int(complete_name))

            clean_non_zero_names = sorted(list(set(non_zero_names)))
            latest_file_number = int(clean_non_zero_names[-1])
            latest_file_number += 1

        return get_file_name_number(latest_file_number)


def get_file_name_number(latest_file_number):
    number_length: int = len(str(latest_file_number))

    if number_length == 1:
        return f'0000{latest_file_number}'

    elif number_length == 2:
        return f'000{latest_file_number}'

    elif number_length == 3:
        return f'00{latest_file_number}'

    elif number_length == 4:
        return f'0{latest_file_number}'

    elif number_length == 5:
        return latest_file_number


def read_inventory_file() -> list:
    """
    Read inventory excel file to print in the terminal output.  Used for confirming user the content before sending.
    :return:
    """
    try:
        excel_sheet = load_workbook(filename='inventory_transaction.xlsx')['Inventory Template']

        inventory_transaction: dict = get_inventory_transaction_data(excel_sheet)
        inventory_rows: dict = get_inventory_rows(inventory_transaction)

        filled_rows: list = []
        for row in inventory_rows:
            row_data = inventory_rows[row]

            count: int = 0
            for cell_data in row_data:
                try:
                    if not cell_data or 'Fill Here' in cell_data:
                        count += 1
                except TypeError:
                    if not str(cell_data) or 'Fill Here' in str(cell_data):
                        count += 1

            if count < 10:
                filled_rows.append(row_data)

        return filled_rows

    except FileNotFoundError:
        print(f'\n\tDo {Fore.RED}Inventory Mode{Style.RESET_ALL} first...')
        input(f'\tPress enter to close Pipe Cleaner')
        sys.exit()


def get_inventory_transaction_data(excel_sheet):
    inventory_transaction: dict = {'from': get_column_data('E', excel_sheet),
                                   'to': get_column_data('F', excel_sheet),
                                   'type': get_column_data('G', excel_sheet),
                                   'part_number': get_column_data('H', excel_sheet),
                                   'qty': get_column_data('I', excel_sheet),
                                   'supplier': get_column_data('J', excel_sheet),
                                   'pipe_number': get_column_data('K', excel_sheet),
                                   'task_number': get_column_data('L', excel_sheet),
                                   'po_number': get_column_data('M', excel_sheet),
                                   'notes': get_column_data('N', excel_sheet)}
    return inventory_transaction


def get_inventory_rows(inventory_transaction) -> dict:
    inventory_rows: dict = {'row_01': [], 'row_02': [], 'row_03': [], 'row_04': [], 'row_05': [], 'row_06': [],
                            'row_07': [], 'row_08': [], 'row_09': [], 'row_10': [], 'row_11': [], 'row_12': []}

    for column_name in inventory_transaction:
        column_data: dict = inventory_transaction[column_name]

        inventory_rows['row_01'].append(column_data[0])
        inventory_rows['row_02'].append(column_data[1])
        inventory_rows['row_03'].append(column_data[2])
        inventory_rows['row_04'].append(column_data[3])
        inventory_rows['row_05'].append(column_data[4])
        inventory_rows['row_06'].append(column_data[5])
        inventory_rows['row_07'].append(column_data[6])
        inventory_rows['row_08'].append(column_data[7])
        inventory_rows['row_09'].append(column_data[8])
        inventory_rows['row_10'].append(column_data[9])
        inventory_rows['row_11'].append(column_data[10])
        inventory_rows['row_12'].append(column_data[11])

    return inventory_rows


def get_column_data(letter: str, sheet) -> list:
    column_data: list = []

    count: int = 10
    while count < 23:
        column_data.append(sheet[f'{letter}{count}'].value)
        count += 1

    return column_data


def print_file_data() -> list:
    """
    Print in terminal output in Pipe Cleaner to showcase the filled in data already before finalizing send
    data to Z: Drive
    """
    file_data: list = read_inventory_file()

    for index, row in enumerate(file_data, start=1):
        print(f'\n\tInventory Transaction #{index}:')

        if not row[0] or row[0] == 'Fill Here':
            print(f'\t\tFrom:      {Fore.RED}{row[0]}{Style.RESET_ALL}')
        else:
            print(f'\t\tFrom:      {row[0]}')

        if not row[1] or row[1] == 'Fill Here':
            print(f'\t\tTo:        {Fore.RED}{row[1]}{Style.RESET_ALL}')
        else:
            print(f'\t\tTo:        {row[1]}')

        if not row[2] or row[2] == 'Fill Here':
            print(f'\t\tType:      {Fore.RED}{row[2]}{Style.RESET_ALL}')
        else:
            print(f'\t\tType:      {row[2]}')

        if not row[3] or row[3] == 'Fill Here':
            print(f'\t\tPart #:    {Fore.RED}{row[3]}{Style.RESET_ALL}')
        else:
            print(f'\t\tPart #:    {row[3]}')

        if not row[4] or row[4] == 'Fill Here':
            print(f'\t\tQTY:       {Fore.RED}{row[4]}{Style.RESET_ALL}')
        else:
            print(f'\t\tQTY:       {row[4]}')

        if not row[5] or row[5] == 'Fill Here':
            print(f'\t\tSupplier:  {Fore.RED}{row[5]}{Style.RESET_ALL}')
        else:
            print(f'\t\tSupplier:  {row[5]}')

        if not row[6] or row[6] == 'Fill Here':
            print(f'\t\tPipe #:    {Fore.RED}{row[6]}{Style.RESET_ALL}')
        else:
            print(f'\t\tPipe #:    {row[6]}')

        if not row[7] or row[7] == 'Fill Here':
            print(f'\t\tTask # :   {Fore.RED}{row[7]}{Style.RESET_ALL}')
        else:
            print(f'\t\tTask #:    {row[7]}')

        if not row[8] or row[8] == 'Fill Here':
            print(f'\t\tPO #:      {Fore.RED}{row[8]}{Style.RESET_ALL}')
        else:
            print(f'\t\tPO #:      {row[8]}')

        if not row[9] or row[9] == 'Fill Here':
            print(f'\t\tNotes:     {Fore.RED}{row[9]}{Style.RESET_ALL}')
        else:
            print(f'\t\tNotes:     {row[9]}')

    return file_data


def get_user_send_response() -> str:
    """
    Get input from user regarding to send current inventory data from excel.
    """
    print(f'\n\tChoose between these options...')
    print(f'\ty -> Yes - Send data to Z: Drive')
    print(f'\tn -> No - Do not send data to Z: Drive')
    return input(f'\n\tChoose option: ')


def is_authorized_user(default_user_name: str, inventory_authorized: tuple) -> bool:
    """
    Checks if current user is apart of the inventory authorized users.
    :param default_user_name: current user
    :param inventory_authorized: authorized users
    :return: True or False
    """
    clean_name: str = default_user_name.casefold()

    for authorized_user in inventory_authorized:
        first_name: str = authorized_user[0].casefold()
        last_name: str = authorized_user[1].casefold()

        if first_name in clean_name and last_name in clean_name:
            return True

    else:
        return False


def get_last_entry_row(excel_sheet) -> int:
    index: int = 10
    for row in range(10_000):
        date_value = str(excel_sheet[f'D{index}'].value)
        name_value = str(excel_sheet[f'E{index}'].value)

        if date_value == 'None' and name_value == 'None':
            return index

        index += 1


def add_row_data(workbook, worksheet, file_path, file_data, last_row, user_name) -> None:
    for index, row in enumerate(file_data, start=0):
        number: int = last_row + index

        worksheet[f'C{number}'] = str('')
        worksheet[f'D{number}'] = str(strftime('%m/%d/%Y'))
        worksheet[f'E{number}'] = str(user_name)
        worksheet[f'F{number}'] = str(row[0])
        worksheet[f'G{number}'] = str(row[1])
        worksheet[f'H{number}'] = str(row[2])
        worksheet[f'I{number}'] = str(row[3])
        worksheet[f'J{number}'] = str(row[4])
        worksheet[f'K{number}'] = str(row[5])
        worksheet[f'L{number}'] = str(row[6])
        worksheet[f'M{number}'] = str(row[7])
        worksheet[f'N{number}'] = str(row[8])
        worksheet[f'O{number}'] = str(row[9])

        cell_00 = worksheet[f'C{number}']
        cell_00.alignment = Alignment(horizontal='center')

        cell_01 = worksheet[f'D{number}']
        cell_01.alignment = Alignment(horizontal='center')

        cell_02 = worksheet[f'E{number}']
        cell_02.alignment = Alignment(horizontal='center')

        cell_03 = worksheet[f'F{number}']
        cell_03.alignment = Alignment(horizontal='center')

        cell_04 = worksheet[f'G{number}']
        cell_04.alignment = Alignment(horizontal='center')

        cell_05 = worksheet[f'H{number}']
        cell_05.alignment = Alignment(horizontal='center')

        cell_06 = worksheet[f'I{number}']
        cell_06.alignment = Alignment(horizontal='center')

        cell_07 = worksheet[f'J{number}']
        cell_07.alignment = Alignment(horizontal='center')

        cell_08 = worksheet[f'K{number}']
        cell_08.alignment = Alignment(horizontal='center')

        cell_09 = worksheet[f'L{number}']
        cell_09.alignment = Alignment(horizontal='center')

        cell_10 = worksheet[f'M{number}']
        cell_10.alignment = Alignment(horizontal='center')

        cell_11 = worksheet[f'N{number}']
        cell_11.alignment = Alignment(horizontal='center')

        cell_12 = worksheet[f'O{number}']
        cell_12.alignment = Alignment(horizontal='center')

    try:
        workbook.save(file_path)
    except PermissionError:
        print(f'\n\tClose {Fore.RED}_inventory_transactions.xlsx{Style.RESET_ALL} to send.')


def get_drive_transaction_logs(file_data: list, user_name: str):
    """
    Get transaction logs found in the Z: Drive or 172.30.1.100 network.
    :return:
    """
    root_path: str = r'Z:\Kirkland_Lab\PipeCleaner\transaction_logs\_inventory_transactions.xlsx'
    alt_path: str = r'\\172.30.1.100\pxe\Kirkland_Lab\PipeCleaner\transaction_logs\_inventory_transactions.xlsx'

    try:
        workbook = load_workbook(root_path)
        worksheet = workbook['All Inventory Transaction']
        last_row = get_last_entry_row(worksheet)
        add_row_data(workbook, worksheet, root_path, file_data, last_row, user_name)

    except FileNotFoundError:
        workbook = load_workbook(alt_path)
        worksheet = workbook['All Inventory Transaction']
        last_row = get_last_entry_row(worksheet)
        add_row_data(workbook, worksheet, alt_path, file_data, last_row, user_name)


def main_method(default_user_name: str, inventory_authorized: tuple) -> None:
    """
    Send data to specific locations in Z: Drive to store information.
    :return:
    """
    user_name: str = get_clean_user_name(default_user_name)

    if is_authorized_user(default_user_name, inventory_authorized):
        print(f'\tAuthorized User [PASS]: {user_name}')
        root_file_path: str = r'Z:\Kirkland_Lab\PipeCleaner\transaction_logs'
        file_number: str = get_file_number(root_file_path)

        file_data: list = print_file_data()

        file_path: str = get_file_path(user_name, root_file_path, file_number)

        user_send_response: str = get_user_send_response()

        if 'YES' in user_send_response.upper() or user_send_response.upper() == 'Y':
            shutil.copyfile('inventory_transaction.xlsx', file_path)
            get_drive_transaction_logs(file_data, user_name)

            print(f'\n\t{Fore.GREEN}Transaction log sent to Z: Drive{Style.RESET_ALL}')
            print(f'\n\tFile Path: {root_file_path}')
            print(f'\tFile Name: {file_number}_{user_name}.xlsx')

            input(f'\n\tPress enter to exit Pipe Cleaner...')
            sys.exit()

        elif 'NO' in user_send_response.upper() or user_send_response.upper() == 'N':
            print(f'\tDid not sent transaction log to Z: Drive')

            input(f'\n\tPress enter to exit Pipe Cleaner...')
            sys.exit()

    else:
        print(f'\tAuthorized User [FAILED]: {user_name}')
        print(f'\n\tPlease notify developer or inventory admin to allow access.')
        print(F'\tPress enter to continue:')
        input()
