"""
Get total from Kirkland - 5/6/2021
"""
from openpyxl import load_workbook

from pipe_cleaner.src.data_console_server import main_method as get_console_server_data


def get_last_entry_row(excel_sheet) -> int:
    """

    :param excel_sheet:
    :return:
    """
    index: int = 2
    for row in range(10_000):
        date_value = str(excel_sheet[f'D{index}'].value)
        name_value = str(excel_sheet[f'E{index}'].value)

        if date_value == 'None' and name_value == 'None':
            return index

        elif date_value == '' and name_value == '':
            return index

        elif not date_value and not name_value:
            return index

        index += 1


def get_drive_last_row(excel_sheet) -> int:
    """

    :param excel_sheet:
    :return:
    """
    index: int = 10
    for row in range(10_000):
        date_value = str(excel_sheet[f'D{index}'].value)
        name_value = str(excel_sheet[f'E{index}'].value)

        if date_value == 'None' and name_value == 'None':
            return index

        elif date_value == '' and name_value == '':
            return index

        elif not date_value and not name_value:
            return index

        index += 1


def get_current_transactions_data(worksheet, last_row) -> list :
    """
    Grab data from Z: Drive on transactions data.
    :param worksheet:
    :param last_row:
    :return:
    """
    transactions_data: list = []

    index: int = 10
    for row in range(10_000):
        current_row: list = [worksheet[f'C{index}'].value,
                             worksheet[f'D{index}'].value,
                             worksheet[f'E{index}'].value,
                             worksheet[f'F{index}'].value,
                             worksheet[f'G{index}'].value,
                             worksheet[f'H{index}'].value,
                             worksheet[f'I{index}'].value,
                             worksheet[f'J{index}'].value,
                             worksheet[f'K{index}'].value,
                             worksheet[f'L{index}'].value,
                             worksheet[f'M{index}'].value,
                             worksheet[f'N{index}'].value,
                             worksheet[f'O{index}'].value]

        clean_data: dict = {'approved': current_row[0],
                            'date': current_row[1],
                            'name': current_row[2],
                            'from': current_row[3],
                            'to': current_row[4],
                            'type': current_row[5],
                            'part_number': current_row[6],
                            'quantity': current_row[7],
                            'supplier': current_row[8],
                            'pipe_number': current_row[9],
                            'task_number': current_row[10],
                            'po_number': current_row[11],
                            'notes': current_row[12]}

        transactions_data.append(clean_data)

        if index == last_row:
            transactions_data.pop(-1)
            return transactions_data

        else:
            index += 1


def get_excel_date(raw_data: str) -> str:
    clean_data = raw_data.replace('datetime.datetime(', '').replace(', 0, 0)', '').replace(' 00:00:00', '')
    year = clean_data[0:4]
    month = clean_data[5:7]
    day = clean_data[8:10]

    return f'{month}/{day}/{year}'


def clean_part_numbers(raw_data: str) -> str:
    clean_data = str(raw_data).split(' ')[0].replace(r'\u00a0', '').strip().upper()
    return clean_data


def clean_numbers(raw_data: str) -> int:
    clean_data = str(raw_data).strip()

    if not clean_data or clean_data.upper() == 'NONE' or clean_data == '':
        return 0
    else:
        try:
            return int(clean_data)
        except ValueError:
            return 0


def get_old_inventory_data(worksheet, last_row: int) -> list:
    """
    Grab data from Z: Drive on transactions data.
    """
    transactions_data: list = []

    index: int = 2
    for row in range(10_000):
        current_row: list = [worksheet[f'A{index}'].value,
                             worksheet[f'B{index}'].value,
                             worksheet[f'C{index}'].value,
                             worksheet[f'D{index}'].value]

        current_data: dict = {'type': str(current_row[0]).strip(),
                              'supplier': str(current_row[1]).strip(),
                              'part_number': clean_part_numbers(current_row[2]),
                              'quantity': clean_numbers(current_row[3])}

        transactions_data.append(current_data)

        if index == last_row:
            return transactions_data

        else:
            index += 1


def get_old_transactions_data(worksheet, last_row: int) -> list:
    """
    Grab data from Z: Drive on transactions data.
    :param worksheet:
    :param last_row:
    :return:
    """
    transactions_data: list = []

    index: int = 2
    for row in range(10_000):
        current_row: list = [get_excel_date(str(worksheet[f'A{index}'].value)),
                             worksheet[f'B{index}'].value,
                             worksheet[f'C{index}'].value,
                             worksheet[f'D{index}'].value,
                             worksheet[f'E{index}'].value,
                             worksheet[f'F{index}'].value,
                             clean_part_numbers(str(worksheet[f'G{index}'].value)),
                             worksheet[f'H{index}'].value,
                             worksheet[f'I{index}'].value]

        current_data: dict = {'date': current_row[0],
                              'person': current_row[1],
                              'from': str(current_row[2]).split('_')[0].lower(),
                              'to': str(current_row[2]).split('_')[-1].lower(),
                              'supplier': current_row[5],
                              'part_number': current_row[6],
                              'quantity': current_row[7],
                              'notes': current_row[8]}

        transactions_data.append(current_data)

        if index == last_row:
            transactions_data.pop(-1)
            return transactions_data

        else:
            index += 1


def clean_current_transactions(current_transactions: list) -> dict:
    """

    :param current_transactions:
    :return:
    """
    clean_data: dict = {}

    for row_data in current_transactions:
        part_number: str = row_data['part_number']
        quantity: int = clean_numbers(row_data['quantity'])

        if not part_number and not quantity:
            pass

        elif part_number in clean_data:
            clean_data[part_number]['quantity'] += quantity

        else:
            clean_data[part_number]: dict = {}
            clean_data[part_number]['quantity']: int = quantity
            clean_data[part_number]['supplier']: str = row_data['supplier']
            clean_data[part_number]['type']: str = row_data['type']

    return clean_data


def get_current_inventory_transactions() -> dict:
    """
    Get current inventory transactions.
    :return:
    """
    root_path: str = r'Z:\Kirkland_Lab\PipeCleaner\transaction_logs\_inventory_transactions.xlsx'
    alt_path: str = r'\\172.30.1.100\pxe\Kirkland_Lab\PipeCleaner\transaction_logs\_inventory_transactions.xlsx'

    try:
        workbook = load_workbook(root_path)
        worksheet = workbook['All Inventory Transaction']
        last_row: int = get_drive_last_row(worksheet)
        current_transactions = get_current_transactions_data(worksheet, last_row)
        return clean_current_transactions(current_transactions)

    except FileNotFoundError:
        workbook = load_workbook(alt_path)
        worksheet = workbook['All Inventory Transaction']
        last_row: int = get_drive_last_row(worksheet)
        current_transactions = get_current_transactions_data(worksheet, last_row)
        return clean_current_transactions(current_transactions)


def clean_old_transactions(old_transactions_data: list) -> dict:
    """
    Convert old transactions to same data structure.
    :param old_transactions_data:
    :return:
    """
    clean_data: dict = {}

    for row_data in old_transactions_data:
        part_number: str = row_data['part_number']

        if part_number in clean_data:
            clean_data[part_number]['quantity'] += int(row_data['quantity'])

        else:
            clean_data[part_number]: dict = {}
            clean_data[part_number]['quantity'] = row_data['quantity']
            clean_data[part_number]['supplier'] = row_data['supplier']

    return clean_data


def get_old_transactions() -> dict:
    base_path: str = 'settings/transaction_logs.xlsx'
    workbook = load_workbook(base_path)
    worksheet = workbook['Sheet1']

    last_row: int = get_last_entry_row(worksheet)
    old_transactions_data = get_old_transactions_data(worksheet, last_row)
    return clean_old_transactions(old_transactions_data)


def clean_old_inventory(raw_old_inventory: list) -> dict:
    """
    Clean old inventory to combine similar data
    :param raw_old_inventory:
    :return:
    """
    clean_data: dict = {}

    for row_data in raw_old_inventory:
        part_number: str = row_data['part_number']

        if part_number == 'NONE' or not part_number or part_number == '':
            pass

        else:
            try:
                if part_number in clean_data:
                    clean_data[part_number]['quantity'] += int(row_data['quantity'])

                else:
                    clean_data[part_number]: dict = {}
                    clean_data[part_number]['type'] = row_data['type']
                    clean_data[part_number]['supplier'] = row_data['supplier']
                    clean_data[part_number]['quantity'] = int(row_data['quantity'])

            except KeyError:
                clean_data[part_number]: dict = {}
                clean_data[part_number]['type'] = row_data['type']
                clean_data[part_number]['supplier'] = row_data['supplier']
                clean_data[part_number]['quantity'] = int(row_data['quantity'])

    return clean_data


def get_old_inventory():
    base_path: str = 'settings/transaction_logs.xlsx'
    workbook = load_workbook(base_path)
    worksheet = workbook['Sheet2']

    last_row: int = get_last_entry_row(worksheet)
    raw_old_inventory = get_old_inventory_data(worksheet, last_row)
    return clean_old_inventory(raw_old_inventory)


def consolidate_transactions_data(old_transactions: dict, current_transactions: dict) -> dict:
    """
    Put together old transactions and current transactions (Shared Drive) for easier comparison later.
    :param old_transactions: static data from given excel file
    :param current_transactions: snapshot data taken from shared drive
    :return:
    """
    consolidated_data: dict = {}

    for old_part_number in old_transactions:
        old_quantity: int = clean_numbers(old_transactions[old_part_number]['quantity'])

        if not old_quantity or old_quantity == 0:
            pass

        elif old_part_number in consolidated_data:
            consolidated_data[old_part_number]['quantity'] += old_quantity

        else:
            consolidated_data[old_part_number]: dict = {}
            consolidated_data[old_part_number]['supplier']: str = old_transactions[old_part_number]['supplier']
            consolidated_data[old_part_number]['quantity'] = old_quantity

    for current_part_number in current_transactions:
        old_quantity: int = clean_numbers(old_transactions[current_part_number]['quantity'])
        transaction_type: str = old_transactions[current_part_number]['type']

        if not old_quantity or old_quantity == 0:
            pass

        elif current_part_number in consolidated_data:
            consolidated_data[current_part_number]['quantity'] += old_quantity

        else:
            consolidated_data[current_part_number]: dict = {}
            consolidated_data[current_part_number]['supplier']: str = old_transactions[current_part_number]['supplier']
            consolidated_data[current_part_number]['type']: str = old_transactions[current_part_number]['type']
            consolidated_data[current_part_number]['quantity']: int = old_quantity

        if not transaction_type or transaction_type == 'None' or transaction_type == '':
            pass

        elif current_part_number in consolidated_data and 'type' in consolidated_data[current_part_number]:
            pass

        else:
            consolidated_data[current_part_number]: dict = {}

    return consolidated_data


def main_method() -> None:
    """
    Consolidate data to produce excel output of inventory update.
    """
    old_transactions: dict = get_old_transactions()
    current_transactions: dict = get_current_inventory_transactions()
    # old_inventory_data: dict = get_old_inventory()
    # console_server_data: dict = get_console_server_data()['inventory']

    # import json
    # foo = json.dumps(old_transactions, sort_keys=True, indent=4)
    # print(foo)
    # input()
