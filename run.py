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
from pipe_cleaner.src.request_inventory import main_method as request_inventory, start_update_form


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


def end_inventory_procedure(start_time: float, file_name: str) -> None:
    """
    End Pipe Cleaner with run time. Automatically bring up Pipe Cleaner excel output.
    """
    print_run_time(start_time)

    open_excel_after_run(file_name)

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


def consolidate_user_data(pipe_cleaner_version: str, default_user_name: str, current_location: str,
                          user_response: dict) -> dict:
    """
    Put together data.
    """
    user_response['version'] = pipe_cleaner_version
    user_response['name'] = default_user_name
    user_response['location'] = current_location
    return user_response


def run_pipe_cleaner(pipe_cleaner_version: str, default_user_name: str, current_location: str) -> None:
    """
    Starts Pipe Cleaner
    """
    start_time: float = time.time()
    inventory_authorized: tuple = get_inventory_authorized()

    user_response: dict = run_terminal(pipe_cleaner_version, default_user_name, current_location)
    user_data: dict = consolidate_user_data(pipe_cleaner_version, default_user_name, current_location, user_response)

    letter: str = user_response['letter']

    if letter == 'N':
        create_excel_output(pipe_cleaner_version, default_user_name, current_location)
        end_program_procedure(start_time)

    elif letter == 'I':
        create_inventory_template(pipe_cleaner_version, default_user_name, current_location)
        end_inventory_procedure(start_time, 'inventory_transaction.xlsx')

    elif letter == 'S':
        send_inventory(default_user_name, inventory_authorized)

    elif letter == 'R':
        request_inventory(user_data)

    elif letter == 'T':
        create_total_inventory(pipe_cleaner_version, default_user_name, current_location)
        end_inventory_procedure(start_time, 'total_kirkland_inventory.xlsx')

    elif letter == 'U':
        start_update_form()


if __name__ == "__main__":
    current_version: str = '2.6.7'
    site_location: str = 'Kirkland Lab Site'
    default_end_user_name: str = getuser().replace(' ', '').strip()

    run_pipe_cleaner(current_version, default_end_user_name, site_location)
