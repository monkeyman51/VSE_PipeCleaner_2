"""
Module responsible for starting Project Pipe Cleaner.
"""

import os
import sys
import time
from getpass import getuser

from pipe_cleaner.src.dashboard_executive_summary import create_excel_output
from pipe_cleaner.src.terminal import run_terminal


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

    open_excel_after_run()

    sys.exit()


def open_excel_after_run() -> None:
    """
    Run excel output after finishing program
    """
    os.system(r"start EXCEL.EXE pipes\main_dashboard.xlsx")


def run_pipe_cleaner(pipe_cleaner_version: str, default_user_name: str, current_location: str) -> None:
    """
    Starts Pipe Cleaner
    """
    start_time: float = time.time()

    run_terminal(pipe_cleaner_version, default_user_name, current_location)

    create_excel_output(pipe_cleaner_version, default_user_name, current_location)

    end_program_procedure(start_time)


if __name__ == "__main__":
    current_version: str = '2.5.7'
    site_location: str = 'Kirkland Lab Site'
    default_end_user_name: str = getuser().replace(' ', '').strip()
    # default_end_user_name: str = 'matthew_hoffman'

    run_pipe_cleaner(current_version, default_end_user_name, site_location)
