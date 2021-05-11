"""
Module responsible for starting Project Pipe Cleaner.
"""

import os
import sys
import time
from getpass import getuser
import csv

from pipe_cleaner.src.dashboard_executive_summary import create_excel_output
from pipe_cleaner.src.inventory_template import create_inventory_template
from pipe_cleaner.src.terminal import run_terminal
from pipe_cleaner.src.send_inventory import main_method as send_inventory
from pipe_cleaner.src.total_inventory import main_method as create_total_inventory


def print_run_time(start_time: float) -> None:
    """
    Output in terminal time taken to run Pipe Cleaner
    :param start_time: Beginning time of Pipe Cleaner
    """
    end_time: float = time.time()
    run_time: float = end_time - start_time

    print(f'\t Run Time: {run_time}')


def end_program_procedure(start_time: float) -> None:
    """
    End Pipe Cleaner with run time. Automatically bring up Pipe Cleaner excel output.
    """
    print_run_time(start_time)

    open_excel_after_run('pipes\main_dashboard.xlsx')

    sys.exit()


def end_inventory_procedure(start_time: float) -> None:
    """
    End Pipe Cleaner with run time. Automatically bring up Pipe Cleaner excel output.
    """
    print_run_time(start_time)

    open_excel_after_run('inventory_transaction.xlsx')

    sys.exit()


def open_excel_after_run(location: str) -> None:
    """
    Run excel output after finishing program
    """
    os.system(fr'start EXCEL.EXE {location}')


def get_inventory_authorized() -> tuple:
    """
    Get people
    :return:
    """
    file_path: str = 'settings/inventory_authorized.csv'

    authorized_people: list = []
    with open(file_path, newline='') as csv_file:
        csv_data = csv.reader(csv_file, delimiter=' ')

        for row in csv_data:
            authorized_people.append(tuple(row))

    return tuple(authorized_people)


def run_pipe_cleaner(pipe_cleaner_version: str, default_user_name: str, current_location: str) -> None:
    """
    Starts Pipe Cleaner
    """
    start_time: float = time.time()
    inventory_authorized: tuple = get_inventory_authorized()

    user_response: str = run_terminal(pipe_cleaner_version, default_user_name, current_location)

    if user_response == 'N':
        create_excel_output(pipe_cleaner_version, default_user_name, current_location)
        end_program_procedure(start_time)

    elif user_response == 'I':
        create_inventory_template(pipe_cleaner_version, default_user_name, current_location)
        end_inventory_procedure(start_time)

    elif user_response == 'S':
        send_inventory(default_user_name, inventory_authorized)

    elif user_response == 'T':
        create_total_inventory()


if __name__ == "__main__":
    current_version: str = '2.6.4'
    site_location: str = 'Kirkland Lab Site'
    default_end_user_name: str = getuser().replace(' ', '').strip()

    run_pipe_cleaner(current_version, default_end_user_name, site_location)
