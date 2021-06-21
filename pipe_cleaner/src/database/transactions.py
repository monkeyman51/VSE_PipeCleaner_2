"""
Export monthly transactions on inventory from latest to earliest.
"""
from os import system
from time import strftime

from openpyxl import load_workbook
from openpyxl.styles import Alignment

from pipe_cleaner.src.log_database import access_database_document


def convert_year_month(year_month: str) -> str:
    """
    Move year and month to cleaner looking version for excel output.
    """
    year: str = year_month[0:4]
    month: str = year_month[5:7]

    if month == '01':
        return f'January {year}'

    elif month == '02':
        return f'February {year}'

    elif month == '03':
        return f'March {year}'

    elif month == '04':
        return f'April {year}'

    elif month == '05':
        return f'May {year}'

    elif month == '06':
        return f'June {year}'

    elif month == '07':
        return f'July {year}'

    elif month == '08':
        return f'August {year}'

    elif month == '09':
        return f'September {year}'

    elif month == '10':
        return f'October {year}'

    elif month == '11':
        return f'November {year}'

    elif month == '12':
        return f'December {year}'


def reverse_transactions(all_transactions: list) -> list:
    """

    """
    order: list = []

    for entry in all_transactions:
        order.append(entry)

    new_order: list = order[::-1]

    return new_order


def add_excel_data(document, worksheet: load_workbook, year_month: str) -> None:
    """
    Add serial number data to excel.
    """
    all_transactions: list = document.find({})
    transactions: list = reverse_transactions(all_transactions)

    for index, serial_number_data in enumerate(transactions, start=3):

        scanned: list = serial_number_data['scanned']

        worksheet[f'A{index}'] = f'# {serial_number_data["_id"]}'
        worksheet[f'B{index}'] = len(scanned)
        worksheet[f'C{index}'] = serial_number_data['part_number']

        worksheet[f'D{index}'] = serial_number_data['time']['date_logged']
        worksheet[f'E{index}'] = serial_number_data['time']['time_logged']

        worksheet[f'F{index}'] = serial_number_data['location']['current']
        worksheet[f'G{index}'] = serial_number_data['location']['previous']
        worksheet[f'H{index}'] = serial_number_data['location']['rack']
        worksheet[f'I{index}'] = serial_number_data['location']['machine']
        worksheet[f'J{index}'] = serial_number_data['location']['pipe']
        worksheet[f'K{index}'] = serial_number_data['location']['site']

        worksheet[f'L{index}'] = serial_number_data['source']['approved_by']
        worksheet[f'M{index}'] = serial_number_data['source']['verified_by']
        worksheet[f'N{index}'] = serial_number_data['source']['trr']
        worksheet[f'O{index}'] = serial_number_data['source']['version']
        worksheet[f'P{index}'] = serial_number_data['source']['comment']
        worksheet[f'Q{index}'] = serial_number_data['source']['task']

        worksheet[f'R{index}'] = ", ".join(scanned)

        for number in range(1, 18):
            current_letter = str(chr(ord('@')+number))
            worksheet[f'{current_letter}{index}'].alignment = Alignment(horizontal='center')

    worksheet['A2'] = f'{convert_year_month(year_month)} - Transactions'
    worksheet['A2'].alignment = Alignment(horizontal='center')


def get_year_month() -> str:
    """

    """
    date: str = strftime('%m/%d/%Y')
    month: str = date[0:2]
    year: str = date[6:10]
    return f'{year}_{month}'


def main_method() -> None:
    """

    """
    print(f'\n\tGetting transactions data from database...')

    year_month: str = get_year_month()
    document = access_database_document('transactions', year_month)

    workbook = load_workbook(fr'settings/transactions_template.xlsx')
    worksheet = workbook['transactions']

    print(f'\n\tCreating excel output...')
    add_excel_data(document, worksheet, year_month)
    workbook.save(fr'pipes/serial_numbers.xlsx')
    system(fr'start EXCEL.EXE pipes/serial_numbers.xlsx')

