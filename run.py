"""
Module responsible for starting Project Pipe Cleaner.
"""
import os
import sys
from getpass import getuser

from pipe_cleaner.src.dashboard.executive_summary import main_method as executive_summary
from pipe_cleaner.src.inventory.output import main_method as inventory_output
from pipe_cleaner.src.dashboard_executive_summary import create_excel_output
from pipe_cleaner.src.database.all_cage import main_method as all_cage
from pipe_cleaner.src.database.find_serial_number import main_method as find_serial_number
from pipe_cleaner.src.database.pn_library import main_method as get_pn_library
from pipe_cleaner.src.database.serial_numbers import main_method as get_serial_numbers
from pipe_cleaner.src.database.transactions import main_method as store_transactions
from pipe_cleaner.src.inventory.all_parts import main_method as all_parts
from pipe_cleaner.src.request_inventory import start_main_method_for_request_form as request_inventory
from pipe_cleaner.src.request_inventory import start_update_form
from pipe_cleaner.src.terminal import run_terminal
from pipe_cleaner.src.inventory.total import main_method as total_inventory
from pipe_cleaner.src.inventory.kirkland_total import main_method as kirkland_total
from pipe_cleaner.src.inventory.kirkland_inventory import main_method as kirkland_inventory
from pipe_cleaner.src.inventory.part_number_library import main_method as get_library
from pipe_cleaner.src.pipe_state import main as pipe_report
from pipe_cleaner.src.rack_manager_powers import main as rack_manager_powers


def end_program_procedure() -> None:
    """
    End Pipe Cleaner with run time. Automatically bring up Pipe Cleaner excel output.
    """
    open_excel_after_run(r'pipes\main_dashboard.xlsx')
    sys.exit()


def end_inventory_procedure(file_name: str) -> None:
    """
    End Pipe Cleaner with run time. Automatically bring up Pipe Cleaner excel output.
    """
    open_excel_after_run(file_name)
    sys.exit()


def open_excel_after_run(location: str) -> None:
    """
    Run excel output after finishing program
    """
    os.system(fr'start EXCEL.EXE {location}')


def consolidate_user_data() -> dict:
    """
    Put together data.
    """
    user_response['name']: str = basic_data['username']
    user_response['location']: str = basic_data['site']
    user_response['version']: str = basic_data['version']

    return user_response


def response_to_user_input() -> None:
    """
    Responding to which letter user chose.  This letter corresponds to options to move the program forward.
    """
    letter = str(user_response["letter"])

    if "R" == letter:
        rack_manager_powers()
        # user_data: dict = consolidate_user_data()
        # request_inventory(user_data)

    elif "N" in letter:
        create_excel_output(basic_data)
        end_program_procedure()

    elif "I" in letter:
        # total_inventory()
        # inventory_output()

        user_data: dict = consolidate_user_data()
        request_inventory(user_data)
        # create_inventory_template(basic_data)
        end_inventory_procedure('inventory_transaction.xlsx')

    elif "D" in letter:
        executive_summary()

    elif "P" in letter:
        pipe_report()

    elif "T" in letter:
        store_transactions()
        # create_total_inventory(basic_data)
        # end_inventory_procedure('total_kirkland_inventory.xlsx')

    elif "U" in letter:
        pass
        # print(f'\n\tForm Number:')
        # print(f'\n\t\t- Enter form number found in inventory email box.')
        # print(f'\t\t- ex. 021-000-001')
        # form_number: str = input(f'\n\t\tEnter Form Number: ')
        # start_update_form(form_number, basic_data)

    elif "S" in letter:
        get_serial_numbers()

    elif "A" in letter:
        all_cage()

    elif "K" in letter:
        kirkland_inventory()
        # all_parts()
        # kirkland_total(basic_data)

    elif "L" in letter:
        get_library(basic_data)

    elif "F" in letter:
        serial_number: str = input(f'\n\n\tEnter Serial Number: ')
        find_serial_number(serial_number)

        for number in range(1, 1_000):
            run_terminal(basic_data)
            response_to_user_input()


if __name__ == "__main__":
    basic_data: dict = {'version': '2.7.7',
                        'site': 'Kirkland Lab Site',
                        'username': getuser().replace(' ', '').strip()}

    user_response: dict = run_terminal(basic_data)
    response_to_user_input()
