"""
Module for creating an excel report on Host Groups page information from Console Server. Will grab other information
from ADO (Azure DevOps), VSE (Veritas Services & Engineers) files to create a more comprehensive report.
"""
import sys
from time import strftime, time, mktime, strptime

import xlsxwriter
from colorama import Fore, Style

from pipe_cleaner.src.dashboard_compare import get_all_issues, get_total_checks, get_missing_tally, get_mismatch_tally
from pipe_cleaner.src.dashboard_compare import main_method as compare_data
from pipe_cleaner.src.dashboard_all_issues import main_method as create_issues_sheet
from pipe_cleaner.src.dashboard_main_setup import main_method as create_setup_sheet
from pipe_cleaner.src.dashboard_main_virtual_machine import main_method as create_virtual_machine_sheet
from pipe_cleaner.src.dashboard_machine_ownership import main_method as create_personal_issues_sheet
from pipe_cleaner.src.dashboard_write import main_method as write_column_data
from pipe_cleaner.src.dashboard_write import parsed_date
from pipe_cleaner.src.data_ado import main_method as get_all_ticket_data
from pipe_cleaner.src.data_console_server import main_method as get_console_server_data
from pipe_cleaner.src.dashboard_inventory import main_method as add_dashboard_inventory
from pipe_cleaner.src.dashboard_all_serial import main_method as add_all_serial
from pipe_cleaner.src.dashboard_all_part_numbers import main_method as add_all_part_numbers
from pipe_cleaner.src.excel_properties import Structure

import os


def set_sheet_structure(current_setup) -> None:
    """
    Create dashboard structure
    """
    set_excel_design(current_setup)
    add_header_data(current_setup)


def add_header_data(current_setup: dict) -> None:
    """
    Add header data on ex. username, date, version, etc.
    """
    add_header_user_name(current_setup)
    add_header_sheet_title(current_setup)
    add_header_site_location(current_setup)
    add_header_date_and_version(current_setup)
    add_header_items_being_checked(current_setup)
    add_user_info_titles(current_setup)


def add_user_info_titles(current_setup: dict):
    worksheet: xlsxwriter = current_setup.get('worksheet')
    structure: xlsxwriter = current_setup.get('structure')
    header_height: xlsxwriter = current_setup.get('header_height')
    upper_header: str = header_height - 1

    worksheet.merge_range('F6:G6', '    Pipes Total', structure.teal_left_14)
    worksheet.merge_range('F7:G7', '    Pipes Available', structure.teal_left_14)
    worksheet.merge_range('F8:G8', '    Primary TRRs', structure.teal_left_14)
    worksheet.merge_range('F9:G9', '    Secondary TRRs', structure.teal_left_14)
    worksheet.merge_range('F10:G10', '    Console Server - No TRR', structure.teal_left_14)
    worksheet.merge_range('F11:G11', '    Blocked TRRs', structure.teal_left_14)

    worksheet.merge_range(f'C{upper_header}:D{upper_header}', f'VSE', structure.teal_middle_14)
    worksheet.merge_range(f'F{upper_header}:H{upper_header}', f'Progress Summary', structure.teal_middle_14)
    worksheet.merge_range(f'J{upper_header}:R{upper_header}', f'Client - TRRs - ADO', structure.teal_middle_14)


def add_header_user_info(console_server_data, current_setup) -> None:
    """

    :param console_server_data:
    :param current_setup:
    :return:
    """
    add_user_info_titles(current_setup)


def add_user_info_totals(console_server_data: dict, current_setup: dict):

    # default_user_name: str = current_setup.get('default_user_name')
    worksheet: xlsxwriter = current_setup.get('worksheet')
    structure: xlsxwriter = current_setup.get('structure')

    pipes_total: int = current_setup.get('pipes_total')
    available: int = current_setup.get('available')
    primary_count: int = current_setup.get('primary_count')
    secondary_count: int = current_setup.get('secondary_count')
    blocked: int = current_setup.get('blocked')
    no_ticket: int = current_setup.get('no_ticket')

    worksheet.write('H6', pipes_total, structure.pale_teal_middle_12)
    worksheet.write('H7', available, structure.pale_teal_middle_12)
    worksheet.write('H8', primary_count, structure.pale_teal_middle_12)
    worksheet.write('H9', secondary_count, structure.pale_teal_middle_12)
    worksheet.write('H10', no_ticket, structure.pale_teal_middle_12)
    worksheet.write('H11', blocked, structure.pale_teal_middle_12)


def get_user_virtual_machines(user_info):
    return user_info['virtual_machines']


def add_user_virtual_machines_total(user_info: dict, current_setup: dict):
    worksheet: xlsxwriter = current_setup.get('worksheet')
    structure: xlsxwriter = current_setup.get('structure')

    try:
        hosts_vms: int = len(get_user_virtual_machines(user_info))
        worksheet.write('G7', hosts_vms, structure.pale_teal_middle_12)

    except KeyError:
        worksheet.write('G7', f'None', structure.pale_teal_middle_12)


def add_user_hosts_total(user_info: dict, current_setup: dict):
    worksheet: xlsxwriter = current_setup.get('worksheet')
    structure: xlsxwriter = current_setup.get('structure')

    try:
        hosts_total: int = len(user_info['systems'])
        worksheet.write('G6', hosts_total, structure.pale_teal_middle_12)

    except KeyError:
        worksheet.write('G6', 'None', structure.pale_teal_middle_12)


def get_user_systems(user_info) -> dict:
    return user_info['systems']


def get_user_pipes_issues(processed_issues: dict, user_systems: dict, user_unique_pipes: list) -> dict:
    """
    Get user assigned machines that have issues sorted within pipes as dictionary
    :param processed_issues: All issues
    :param user_systems:
    :param user_unique_pipes:
    :return:
    """
    pipes_issues: dict = {}

    for user_pipe in user_unique_pipes:
        try:
            pipe_issues: dict = processed_issues[user_pipe]

            pipes_issues[user_pipe] = {}
            count: int = 0
            for user_system in user_systems:
                for pipe_issue in pipe_issues:

                    if user_system in pipe_issue:
                        current_pipe_issue: dict = processed_issues[user_pipe][pipe_issue]
                        pipes_issues[user_pipe][count] = current_pipe_issue
                        count += 1
        except KeyError:
            pipes_issues[user_pipe] = ''
    return pipes_issues


def get_user_unique_pipes(user_pipes):
    return sorted(list(set(user_pipes)))


def get_user_pipe_total(user_info) -> int:
    user_systems: dict = get_user_systems(user_info)
    user_pipes: list = get_user_pipes(user_systems)
    user_unique_pipes: list = get_user_unique_pipes(user_pipes)
    return len(user_unique_pipes)


def get_user_pipes(user_systems):
    all_pipes: list = []
    for item in user_systems:
        if 'VSE' in item and '-' in item:
            all_pipes.append(user_systems[item]['pipe_name'])
    return all_pipes


def add_user_pipes_total(user_info: dict, current_setup: dict):
    worksheet: xlsxwriter = current_setup.get('worksheet')
    structure: xlsxwriter = current_setup.get('structure')

    try:
        user_pipes_total: int = get_user_pipe_total(user_info)

        worksheet.write('G5', user_pipes_total, structure.pale_teal_middle_12)

    except KeyError:
        worksheet.write('G5', f'None', structure.pale_teal_middle_12)


def default_name_period_to_underscore(default_name):
    if 'steph' in default_name and '.ak' in default_name:
        return 'steph_ak'
    else:
        return default_name.replace('.', '_').replace('-EXT', '')


def get_user_info(console_server_data, default_name) -> dict:
    default_name_underscore: str = default_name_period_to_underscore(default_name)

    try:
        return console_server_data['user_base'][default_name_underscore]

    except KeyError:
        import sys
        print(f'\n')
        print(f'\tDear {default_name.replace(".", " ").title()},')
        print(f'\tPipe Cleaner did not detect any machines checked out under {default_name} within Console Server.')
        print(f'\tPlease checkout a system in order to use Personal Issues page.')
        print(f'\n\tPress enter to continue...')
        return {}


def add_header_items_being_checked(current_setup: dict) -> None:
    """
    These items under testing are meant to be components still not 100% confident
    """
    worksheet: xlsxwriter = current_setup.get('worksheet')
    structure: xlsxwriter = current_setup.get('structure')

    worksheet.merge_range('F2:H2', 'ITEMS BEING CHECKED', structure.teal_middle_14)
    worksheet.merge_range('F3:H3', 'BIOS, BMC, CPLD, OS, Ticket', structure.pale_teal_middle_12_normal)
    worksheet.merge_range('F4:H4', 'Configured Systems, Virtual Machines', structure.pale_teal_middle_12_normal)


def add_header_date_and_version(current_setup: dict) -> None:
    """
    Adds the current date/time and Pipe Cleaner version to the header area in the top left corner
    """
    current_time: str = strftime('%I:%M %p')
    current_date: str = strftime('%m/%d/%Y')
    pipe_cleaner_version: str = current_setup.get('version')

    worksheet: xlsxwriter = current_setup.get('worksheet')
    structure: xlsxwriter = current_setup.get('structure')

    pipe_cleaner_version: str = clean_pipe_cleaner_version(pipe_cleaner_version)

    worksheet.write('B8', f'            {current_date} - {current_time} - {pipe_cleaner_version}',
                    structure.italic_blue_font)


def clean_pipe_cleaner_version(pipe_cleaner_version) -> str:
    """
    Version for documentation
    :param pipe_cleaner_version:
    :return: cleaner version
    """
    return f"v{pipe_cleaner_version.split(' ')[0]}"


def add_header_site_location(current_setup: dict) -> None:
    """
    Adds the site location to the header area in the top left corner
    """
    site_location: str = current_setup.get('site_location')

    worksheet: xlsxwriter = current_setup.get('worksheet')
    structure: xlsxwriter = current_setup.get('structure')

    worksheet.write('B6', f'        {site_location}', structure.bold_italic_blue_font)


def add_header_sheet_title(current_setup: dict) -> None:
    """
    Adds the excel sheet name to the header area in the top left corner
    """
    sheet_name: str = current_setup.get('sheet_name')

    worksheet: xlsxwriter = current_setup.get('worksheet')
    structure: xlsxwriter = current_setup.get('structure')

    worksheet.write('B5', f'  Pipe Cleaner - {sheet_name}', structure.big_blue_font)


def add_header_user_name(current_setup: dict):
    """
    Add clean user name to the top left corner.
    """
    clean_name: str = current_setup.get('clean_name')

    worksheet: xlsxwriter = current_setup.get('worksheet')
    structure: xlsxwriter = current_setup.get('structure')

    worksheet.write('B7', f'            {clean_name}', structure.bold_italic_blue_font)
    worksheet.merge_range('C10:D11', 'WORK IN PROGRESS', structure.light_red_middle_14)


def set_excel_design(current_setup: dict) -> None:
    """
    Set up excel output design/parameters.
    """
    set_rows_and_columns_sizes(current_setup)

    add_column_titles(current_setup)
    add_freeze_panes(current_setup)
    add_vse_logo_top_right(current_setup)


def add_column_titles(current_setup: dict) -> None:
    """
    Set up Column Names in the Excel table for categorizing into vertical data later
    """
    header_height: int = current_setup.get('header_height')
    left_padding: int = current_setup.get('left_padding')
    column_names: tuple = current_setup.get('column_names')

    worksheet: xlsxwriter = current_setup.get('worksheet')
    structure: xlsxwriter = current_setup.get('structure')

    # Hyperlink to Host Group page within Console Server, should be for column title "Pipe"
    host_group_column: str = get_letter_for_column_position(initial=0, left_padding=2)

    for index, column_title in enumerate(column_names, start=0):
        position: str = get_column_title_position(header_height, index, left_padding)

        if 'PIPE' in column_title.upper() and host_group_column in position[0]:
            worksheet.write_url(position, 'http://172.30.1.100/console/host_groups.php',
                                structure.teal_middle, column_title)

        elif not column_title:
            add_white_cell(position, current_setup)

        else:
            add_column_title(position, column_title, current_setup)


def add_white_cell(position: str, current_setup) -> None:
    """
    Account for empty cells that don't have column title.
    Meant for giving space between different groups of data.
    """
    worksheet: xlsxwriter = current_setup.get('worksheet')
    structure: xlsxwriter = current_setup.get('structure')

    worksheet.write(position, '', structure.white)


def add_column_title(position: str, column_title: str, current_setup: dict) -> None:
    """
    Add column title to the current excel sheet
    """
    worksheet: xlsxwriter = current_setup.get('worksheet')
    structure: xlsxwriter = current_setup.get('structure')

    worksheet.write(position, column_title, structure.teal_middle)


def get_column_title_position(header_height: int, index: int, left_padding: int) -> str:
    """
    Get position of the column title based on excel position from the letter and number ex. A1, B4, C3
    """
    letter: str = get_letter_for_column_position(index, left_padding)
    return f'{letter}{header_height}'


def convert_index_to_letter(index: int) -> str:
    """

    :param index: Current index due to how many columns we care about in the excel output sheet
    """
    lower_character = chr(ord('a') + index)
    return str(lower_character).upper()


def get_letter_for_column_position(initial: int, left_padding: int) -> str:
    """
    For positioning the column title based on starting point of the left padding.
    :return: letter of excel column
    """
    return convert_index_to_letter(initial + left_padding)


def add_vse_logo_top_right(current_setup: dict) -> None:
    """
    Creates VSE Logo on the top left corner
    :param current_setup:
    """
    worksheet: xlsxwriter = current_setup.get('worksheet')

    worksheet.insert_image('A1', 'pipe_cleaner/img/vse_logo.png')


def add_freeze_panes(current_setup: dict) -> None:
    """
    Allows information to the left to stay
    """
    header_height: int = current_setup.get('header_height')
    freeze_pane_position: int = current_setup.get('freeze_pane_position')
    worksheet: xlsxwriter = current_setup.get('worksheet')

    worksheet.freeze_panes(header_height, freeze_pane_position)


def set_rows_and_columns_sizes(current_setup) -> None:
    """
    Beginning of the Excel Structure
    """
    rows_height: tuple = current_setup.get('rows_height')
    columns_width: tuple = current_setup.get('columns_width')

    worksheet: xlsxwriter = current_setup.get('worksheet')
    structure: xlsxwriter = current_setup.get('structure')

    set_header_rows_height(rows_height, worksheet, structure)
    set_excel_column_width(columns_width, worksheet, structure)


def set_header_rows_height(rows_height: tuple, worksheet: xlsxwriter, structure: xlsxwriter) -> None:
    """
    Establishes current worksheet row vertical heights for the header.
    """
    for index, row_size in enumerate(rows_height, start=0):
        worksheet.set_row(index, row_size, structure.white)


def set_excel_column_width(columns_width: tuple, worksheet: xlsxwriter, structure: xlsxwriter) -> None:
    """
    Establishes current worksheet column widths.
    """
    for index in range(0, len(columns_width)):
        current_letter: str = convert_index_to_letter(index)

        worksheet.set_column(f'{current_letter}:{current_letter}',
                             columns_width[index],
                             structure.white)


def set_layout(worksheet, structure):
    """
    Beginning of the Excel Structure
    :return:
    """
    worksheet.set_row(0, 12, structure.white)
    worksheet.set_row(1, 20, structure.white)
    worksheet.set_row(2, 16, structure.white)
    worksheet.set_row(3, 15, structure.white)
    worksheet.set_row(4, 15, structure.white)
    worksheet.set_row(5, 15, structure.white)
    worksheet.set_row(6, 15, structure.white)
    worksheet.set_row(7, 15, structure.white)
    worksheet.set_row(8, 15, structure.white)
    worksheet.set_row(9, 15, structure.white)
    worksheet.set_row(10, 15, structure.white)
    worksheet.set_row(11, 15, structure.white)

    worksheet.set_column('A:A', 5.5, structure.white)
    worksheet.set_column('B:B', 26, structure.white)
    worksheet.set_column('C:C', 40, structure.white)
    worksheet.set_column('D:D', 25, structure.white)
    worksheet.set_column('E:E', 24, structure.white)
    worksheet.set_column('F:F', 11, structure.white)
    worksheet.set_column('G:G', 11, structure.white)
    worksheet.set_column('H:H', 11, structure.white)
    worksheet.set_column('I:I', 11, structure.white)
    worksheet.set_column('J:J', 25, structure.white)
    worksheet.set_column('K:K', 18, structure.white)
    worksheet.set_column('L:L', 25, structure.white)
    worksheet.set_column('M:M', 25, structure.white)
    worksheet.set_column('N:N', 25, structure.white)
    worksheet.set_column('O:O', 25, structure.white)
    worksheet.set_column('P:P', 25, structure.white)


def set_column_names(top_plane_height, worksheet, structure):
    """
    Set up Column Names in the Excel table for adding data later
    :param top_plane_height:
    :param worksheet:
    :param structure:
    :return:
    """
    name_to_number: dict = {}

    column_names: list = ['Pipe Name',
                          'Description',
                          'Checked Out To',
                          'Status',
                          'Setup',
                          'PM',
                          'TECH',
                          'ENG',
                          'Expected Start',
                          'Schedule']

    # Number part of the excel position
    num = str(top_plane_height)

    initial = 0
    while initial < len(column_names):
        little = chr(ord('b') + initial)
        let = str(little).upper()

        if let == 'B' or let == 'C':
            worksheet.write(f'{let}{num}', f'{column_names[initial]}', structure.teal_left)
        else:
            worksheet.write(f'{let}{num}', f'{column_names[initial]}', structure.teal_middle)

        # Create key for dictionary
        name = str(column_names[initial]).lower().replace(' ', '_')
        number = initial + 1

        name_to_number[name] = str(number)

        initial += 1

    return name_to_number


def add_dashboard_data(console_server_data: dict, current_setup: dict, azure_devops_data: dict) -> None:
    """
    Actual adding data to the Main Dashboard
    """
    worksheet: xlsxwriter = current_setup.get('worksheet')
    structure: xlsxwriter = current_setup.get('structure')

    pipe_name_column = 'B'
    description_column = 'C'
    checked_out_to_column = 'D'
    status_column = 'E'
    setup_column = 'F'
    pm_column = 'G'
    tech_column = 'H'
    eng_column = 'I'
    due_date_column = 'J'

    # Store real pipes with proper naming conventions in list for writing dashboard data later
    real_pipes: list = []
    for supposed_pipe in console_server_data:
        try:
            description = console_server_data.get(supposed_pipe, {}).get('description', 'None')

            if 'Pipe-' in supposed_pipe and 'OFFLINE' not in supposed_pipe and 'OFFLINE' not in description \
                    and '[' in supposed_pipe and ']' in supposed_pipe:

                if console_server_data.get(supposed_pipe, {}).get('compare_data', 'None') != 'None':
                    real_pipes.append(supposed_pipe)
        except AttributeError:
            pass

    print(f'\n\t=====================================================================')
    print(f'\t  Pipe Excel Reports - Collecting and Processing Data:')
    print(f'\t=====================================================================')
    print(f'\t\t  STATUS   |  REASON    |  HOST GROUP NAME')

    # Initial accounts for starting point of the dashboard data
    initial = 14
    for index, pipe in enumerate(real_pipes):

        start: int = index + initial
        check_color: int = start % 2
        current_pipe_data: dict = console_server_data.get(pipe, 'None')

        # Accounts for no pipe data
        if current_pipe_data != 'None':

            # Current data within Pipe
            current_compare_data: dict = current_pipe_data.get('compare_data', 'None')
            setup_data: dict = current_pipe_data.get('setup_data', 'None')

            current_host_group_id: str = console_server_data.get(pipe, {}).get('host_id', 'None')

            # Ticket Data
            due_dates: dict = {}
            try:
                base_ticket: list = []

                group_unique_tickets: list = current_pipe_data.get('group_unique_tickets', 'None')
                one_ticket: str = group_unique_tickets[0]
                base_ticket.append(one_ticket)
                ticket = base_ticket[0]

                due_dates['base_ticket'] = ticket

                due_dates['actual_qual_end_date'] = azure_devops_data.get(ticket, {}).get('due_dates', {}). \
                    get('actual_qual_end_date', 'None')

                due_dates['actual_qual_start_date'] = azure_devops_data.get(ticket, {}).get('due_dates', {}). \
                    get('actual_qual_start_date', 'None')

                due_dates['expected_task_completion'] = azure_devops_data.get(ticket, {}).get('due_dates', {}). \
                    get('expected_task_completion', 'None')

                due_dates['expected_task_start'] = azure_devops_data.get(ticket, {}).get('due_dates', {}). \
                    get('expected_task_start', 'None')

            except IndexError:
                pass

            # Create Mini Dashboard
            # write_pipe_dashboard(pipe, console_server_data, azure_devops_data, all_issues)
            print(f'\t\t- Collect  |  {Fore.GREEN}Success{Style.RESET_ALL}   |  {pipe}.xlsx')

            # Writes to the Main Dashboard
            write_column_data('pipe_name', structure, worksheet, start, check_color, pipe_name_column,
                              pipe, '', current_host_group_id)

            write_column_data('description', structure, worksheet, start, check_color, description_column,
                              current_pipe_data.get('description'), '', current_host_group_id)

            write_column_data('checked_out_to', structure, worksheet, start, check_color, checked_out_to_column,
                              current_pipe_data.get('checked_out_to'), '', current_host_group_id)

            write_column_data('status', structure, worksheet, start, check_color, status_column,
                              current_pipe_data.get('host_group_status'), '', current_host_group_id)

            write_column_data('setup_column', structure, worksheet, start, check_color, setup_column,
                              setup_data, '', current_host_group_id)

            write_column_data('pm_column', structure, worksheet, start, check_color, pm_column,
                              current_compare_data, pipe, current_host_group_id)

            write_column_data('tech_column', structure, worksheet, start, check_color, tech_column,
                              current_compare_data, pipe, current_host_group_id)

            write_column_data('eng_column', structure, worksheet, start, check_color, eng_column,
                              '', '', current_host_group_id)

            write_column_data('due_date_column', structure, worksheet, start, check_color, due_date_column,
                              due_dates, '', current_host_group_id)


def add_workload_data(overall_workload: dict, worksheet, structure):
    """

    :param overall_workload:
    :param worksheet:
    :param structure:
    :return:
    """

    # Set up where location of the column is
    column_vse_employee: str = 'B'
    column_pipes: str = 'C'
    column_systems: str = 'D'

    all_employees: list = []
    for person in overall_workload:
        clean_person = str(person).lower()
        if clean_person == 'none' or clean_person == '' or clean_person is None:
            pass
        else:
            all_employees.append(clean_person)

    sorted_employees = sorted(all_employees)

    # Initial accounts for starting point of the dashboard data
    initial: int = 14
    for index, person in enumerate(sorted_employees):

        clean_name = str(person).replace('_', ' ').title()
        pipes_count = str(overall_workload.get(person, {}).get('pipes', 'None'))
        systems_count = str(overall_workload.get(person, {}).get('systems', 'None'))

        # Increments
        start = index + initial
        check_color = start % 2

        if check_color == 0:
            # Pipe Name
            write_workload_blue(column_vse_employee, start, clean_name, worksheet, structure)

            # Machine Name
            write_workload_blue(column_pipes, start, pipes_count, worksheet, structure)

            # Ticket
            write_workload_blue(column_systems, start, systems_count, worksheet, structure)

        elif check_color == 1:
            # Pipe Name
            write_workload_alt(column_vse_employee, start, clean_name, worksheet, structure)

            # Machine Name
            write_workload_alt(column_pipes, start, pipes_count, worksheet, structure)

            # Ticket
            write_workload_alt(column_systems, start, systems_count, worksheet, structure)


def write_workload_blue(column: str, start_location: int, data: str, worksheet, structure):
    """
    Create output
    :param column:
    :param start_location:
    :param data:
    :param worksheet:
    :param structure:
    :return:
    """
    if ' ' in data:
        worksheet.write(f'{column}{str(start_location)}', data, structure.blue_left)
    elif data == 'None' or data == '':
        worksheet.write(f'{column}{str(start_location)}', '', structure.missing_cell)
    elif data == 'MISMATCH':
        worksheet.write(f'{column}{str(start_location)}', data, structure.neutral_cell)
    elif data == 'MISSING':
        worksheet.write(f'{column}{str(start_location)}', data, structure.bad_cell)
    elif 'VSE' in data or 'Pipe-' in data:
        worksheet.write(f'{column}{str(start_location)}', data, structure.blue_left)
    else:
        worksheet.write(f'{column}{str(start_location)}', data, structure.blue_middle)


def write_workload_alt(column: str, start_location: int, data: str, worksheet, structure):
    """
    Create output
    :param column:
    :param start_location:
    :param data:
    :param worksheet:
    :param structure:
    :return:
    """
    if ' ' in data:
        worksheet.write(f'{column}{str(start_location)}', data, structure.alt_blue_left)
    elif data == 'None' or data == '':
        worksheet.write(f'{column}{str(start_location)}', '', structure.missing_cell)
    elif data == 'MISMATCH':
        worksheet.write(f'{column}{str(start_location)}', data, structure.neutral_cell)
    elif data == 'MISSING':
        worksheet.write(f'{column}{str(start_location)}', data, structure.bad_cell)
    elif 'VSE' in data or 'Pipe-' in data:
        worksheet.write(f'{column}{str(start_location)}', data, structure.alt_blue_left)
    else:
        worksheet.write(f'{column}{str(start_location)}', data, structure.alt_blue_middle)


def get_pipe_numbers(console_server_data: dict) -> dict:
    """
    Get authentic pipes that have standards agreed on.
    :param console_server_data:
    :return: pipe number
    """
    pipe_number: int = 0
    system_number: int = 0

    for possible_pipe in console_server_data:
        try:
            host_group_status: str = console_server_data.get(possible_pipe, {}).get('host_group_status', 'None')
            if 'Pipe-' in possible_pipe and 'OFFLINE' not in possible_pipe and 'OFFLINE' not in host_group_status:
                pipe_number += 1
                systems = console_server_data.get(possible_pipe, {}).get('pipe_data', 'None')
                if systems != 'None':
                    system_number += len(systems)
        except AttributeError:
            pass

    return {'pipes': pipe_number, 'systems': system_number}


def create_executive_summary(pipe_cleaner_version: str, default_user_name: str, site_location: str) -> dict:
    """
    Current excel sheet design to setup the excel tab for data to fill in later.
    """
    sheet_title: str = 'Executive Summary'

    check_opened_pipe_cleaner()

    workbook = xlsxwriter.Workbook('pipes/main_dashboard.xlsx')

    excel_setup: dict = {'sheet_title': sheet_title,
                         'worksheet': workbook.add_worksheet(sheet_title),
                         'structure': Structure(workbook),
                         'workbook': workbook,
                         'version': pipe_cleaner_version,
                         'site_location': site_location,
                         'default_user_name': default_user_name,
                         'clean_name': get_clean_user_name(default_user_name),
                         'header_height': 14,
                         'body_position': 15,
                         'left_padding': 2,
                         'freeze_pane_position': 4,
                         'host_group_hyperlink': 'http://172.30.1.100/console/host_groups.php',
                         'rows_height': (12.0, 19.5, 16.5, 16.5, 19.5, 19.5, 19.5, 19.5, 19.5, 19.5, 15.75, 24.0),
                         'columns_width': (0.25, 0.25, 22.0, 38.0, 2.0, 16.0, 16.0, 16.0, 2.0, 12.0, 10.0, 10.0,
                                           19.0, 25.0, 21.0, 21.0, 21.0, 21.0, 25.0, 25.0, 25.0, 25.0),
                         'column_names': ('Pipe',
                                          'Host Group Description',
                                          '',
                                          'PMs',
                                          'TECHs',
                                          'ENGs',
                                          '',
                                          'Type',
                                          'TRR',
                                          'Qual',
                                          'State',
                                          'Assigned To',
                                          'Expected Start',
                                          'Expected End',
                                          'Actual Start',
                                          'Actual End'),
                         'pipes_total': 0,
                         'available': 0,
                         'no_ticket': 0,
                         'primary_count': 0,
                         'secondary_count': 0,
                         'blocked': 0}

    return excel_setup


def check_opened_pipe_cleaner():
    try:
        os.remove('pipes/main_dashboard.xlsx')
    except FileNotFoundError:
        pass
    except PermissionError:
        print(f'\tExcel file {Fore.RED}main_dashboard.xlsx{Style.RESET_ALL} already up.\n'
              f'\tPlease close excel file and restart Pipe Cleaner.')
        print(f'\n')
        print(f'\tPress {Fore.LIGHTBLUE_EX}enter{Style.RESET_ALL} to close program.\n')
        input()
        sys.exit()


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

        return ''.join(indexed_clean_name).replace('-Ext', '')
    else:
        return default_user_name.replace('.', ' ').title().replace('-Ext', '')


def process_console_server(console_server_data: dict):
    """

    :param console_server_data:
    :return:
    """
    # Get relevant pipe information
    real_pipes: list = []
    for potential_pipe in console_server_data:
        if 'Pipe-' in potential_pipe and 'OFFLINE' not in potential_pipe \
                and '(' not in potential_pipe and ')' not in potential_pipe:
            try:
                host_group_status: str = console_server_data.get(potential_pipe, {}).get('host_group_status')
                if 'OFFLINE' not in host_group_status.upper():
                    real_pipes.append(potential_pipe)
            except TypeError:
                pass

            except AttributeError:
                pass

    process_console_server_data: dict = {}
    for real_pipe in real_pipes:
        process_console_server_data[real_pipe]: dict = {}
        checked_out_to: str = console_server_data.get(real_pipe, {}).get('checked_out_to')
        description: str = console_server_data.get(real_pipe, {}).get('description')
        host_group_status: str = console_server_data.get(real_pipe, {}).get('host_group_status')
        host_id: str = console_server_data.get(real_pipe, {}).get('host_id')
        group_unique_tickets: str = console_server_data.get(real_pipe, {}).get('group_unique_tickets')

        # Get tally for Dead or Alive on connection status
        alive_tally = 0
        filled_tally = 0
        in_use_tally = 0
        pipe_data: str = console_server_data.get(real_pipe, {}).get('pipe_data')
        try:
            for machine in pipe_data:
                try:
                    connection_status: str = console_server_data.get(real_pipe, {}).get('pipe_data', {}). \
                        get(machine, {}).get('connection_status')
                    if connection_status.upper() == 'ALIVE':
                        alive_tally += 1
                except AttributeError:
                    pass
                except TypeError:
                    pass

                try:
                    ticket: str = console_server_data.get(real_pipe, {}).get('pipe_data', {}). \
                        get(machine, {}).get('ticket')
                    if ticket.upper().isdigit():
                        filled_tally += 1
                except AttributeError:
                    pass
                except TypeError:
                    pass

                try:
                    checked_out_to: str = console_server_data.get(real_pipe, {}).get('pipe_data', {}). \
                        get(machine, {}).get('checked_out_to')
                    if check_missing(checked_out_to) != 'None':
                        in_use_tally += 1
                except AttributeError:
                    pass
                except TypeError:
                    pass

            process_console_server_data[real_pipe]['total_tally'] = len(pipe_data)

        except TypeError:
            pass

        process_console_server_data[real_pipe]['alive_tally'] = alive_tally
        process_console_server_data[real_pipe]['filled_tally'] = filled_tally
        process_console_server_data[real_pipe]['in_use_tally'] = in_use_tally

        process_console_server_data[real_pipe]['host_id'] = host_id
        process_console_server_data[real_pipe]['host_group_status'] = host_group_status
        process_console_server_data[real_pipe]['description'] = description
        process_console_server_data[real_pipe]['checked_out_to'] = checked_out_to
        process_console_server_data[real_pipe]['group_unique_tickets'] = group_unique_tickets

    return process_console_server_data


def check_missing(data: str) -> str:
    """

    :param data:
    :return:
    """
    clean_data = data.replace(' ', '')
    if clean_data == 'None' or clean_data == '' or clean_data is None or not clean_data:
        return 'None'
    else:
        return data


def process_pipe_name(pipe_name: str):
    """
    Shorten pipe name to fit into excel output
    :param pipe_name:
    :return:
    """
    clean_data: str = pipe_name. \
        replace('[', ''). \
        replace(']', ''). \
        replace("'", '')

    last_part: str = clean_data.split(' ')[-1]

    return clean_data.replace('Pipe-', '').replace(last_part, '')


def write_pipe_name_column(processed_console_server: dict, console_server_data: dict, current_setup: dict) -> None:
    """
    Writes Pipe Name column in excel output
    """
    letter: str = 'C'
    worksheet: xlsxwriter = current_setup.get('worksheet')
    structure: xlsxwriter = current_setup.get('structure')
    current_position: int = current_setup.get('body_position')

    for index, pipe_name in enumerate(processed_console_server, start=0):

        current_color: xlsxwriter = get_current_color(index, structure)
        pipe_hyperlink: str = get_pipe_hyperlink(console_server_data, pipe_name)
        clean_pipe_name: str = f' {process_pipe_name(pipe_name)}'

        total_tickets: int = get_total_tickets(pipe_name, processed_console_server)
        max_number: int = current_position + total_tickets - 1

        base_position: str = get_base_position(letter, current_position)
        merge_position: str = get_merge_position(base_position, letter, max_number)

        if check_missing(pipe_name) == 'None':
            if total_tickets >= 2:
                add_empty_merge_cell(merge_position, structure, worksheet)
                add_empty_hyperlink_cell(base_position, pipe_hyperlink, current_setup)
            else:
                add_empty_hyperlink_cell(base_position, pipe_hyperlink, current_setup)
                set_single_ticket_row(current_position, worksheet)

        elif total_tickets >= 2:
            worksheet.merge_range(merge_position, clean_pipe_name, current_color)
            worksheet.write_url(base_position, pipe_hyperlink, current_color, string=clean_pipe_name)

        else:
            worksheet.write_url(base_position, pipe_hyperlink, current_color, string=clean_pipe_name)
            set_single_ticket_row(current_position, worksheet)

        worksheet.set_row(max_number, 3.75, structure.white)
        current_position: int = max_number + 2


def add_empty_hyperlink_cell(base_position: str, pipe_hyperlink: str, current_setup: dict):
    worksheet: xlsxwriter = current_setup.get('worksheet')
    structure: xlsxwriter = current_setup.get('structure')
    worksheet.write_url(base_position, pipe_hyperlink, structure.missing_cell, string='')


def set_single_ticket_row(current_position, worksheet):
    worksheet.set_row(current_position - 1, 28.5)


def add_empty_merge_cell(merge_position, structure, worksheet):
    worksheet.merge_range(merge_position, '', structure.aqua_left_12)


def get_total_tickets(pipe_name, processed_console_server) -> int:
    total_tickets: int = 0
    pipe_total_trr: int = len(processed_console_server.get(pipe_name, {}).get('group_unique_tickets'))

    return increment_total_tickets(pipe_total_trr, total_tickets)


def get_pipe_hyperlink(console_server_data, pipe_name):
    host_group_id: str = console_server_data.get(pipe_name, {}).get('host_id', 'None')
    return f'http://172.30.1.100/console/host_group_host_list.php?host_group_id={host_group_id}'


def get_merge_position(base_position, letter, max_number):
    max_position: str = get_max_position(letter, max_number)
    return f'{base_position}:{max_position}'


def increment_total_tickets(pipe_total_trr, total_tickets):
    if pipe_total_trr == 0:
        total_tickets += 1
    else:
        total_tickets += pipe_total_trr
    return total_tickets


def get_current_color(index, structure):
    result: int = index % 2
    if result == 0:
        return structure.blue_middle_22
    elif result == 1:
        return structure.alt_blue_middle_22


def get_current_color_11_left(index, structure):
    result: int = index % 2
    if result == 0:
        return structure.blue_left
    elif result == 1:
        return structure.alt_blue_left


def get_current_color_11_middle(index, structure):
    result: int = index % 2
    if result == 0:
        return structure.blue_middle
    elif result == 1:
        return structure.alt_blue_middle


def get_current_color_12_middle(index, structure):
    result: int = index % 2
    if result == 0:
        return structure.blue_middle_12
    elif result == 1:
        return structure.alt_blue_middle_12


def get_current_color_16_middle(index, structure):
    result: int = index % 2
    if result == 0:
        return structure.blue_middle_16
    elif result == 1:
        return structure.alt_blue_middle_16


def get_current_color_14_middle(index, structure):
    result: int = index % 2
    if result == 0:
        return structure.blue_middle_14
    elif result == 1:
        return structure.alt_blue_middle_14


def get_current_color_12_left(index, structure):
    result: int = index % 2
    if result == 0:
        return structure.blue_left_12
    elif result == 1:
        return structure.alt_blue_left_12


def get_max_position(letter, max_number) -> str:
    return f'{letter}{max_number}'


def get_base_position(column, current_position) -> str:
    return f'{column}{current_position}'


def write_description_column(processed_console_server, console_server_data, current_setup):
    """
    Add column data in excel output
    """
    letter: str = 'D'
    worksheet: xlsxwriter = current_setup.get('worksheet')
    structure: xlsxwriter = current_setup.get('structure')
    current_position: int = current_setup.get('body_position')

    for index, pipe_name in enumerate(processed_console_server, start=0):

        current_color: xlsxwriter = get_current_color_11_left(index, structure)
        clean_description: str = clean_pipe_description(console_server_data, pipe_name)

        total_tickets: int = get_total_tickets(pipe_name, processed_console_server)
        max_number: int = current_position + total_tickets - 1

        base_position: str = get_base_position(letter, current_position)
        merge_position: str = get_merge_position(base_position, letter, max_number)

        if check_missing(clean_description) == 'None':
            if total_tickets >= 2:
                worksheet.merge_range(merge_position, '', structure.missing_cell)
            else:
                worksheet.write(base_position, '', structure.missing_cell)

                set_single_ticket_row(current_position, worksheet)

        elif total_tickets >= 2:
            worksheet.merge_range(merge_position, clean_description, current_color)
        else:
            worksheet.write(base_position, clean_description, current_color)
            set_single_ticket_row(current_position, worksheet)

        worksheet.set_row(max_number, 3.75, structure.white)
        current_position = max_number + 2


def add_trr_type_column(processed_console_server, azure_devops_data, current_setup):
    """
    Add column data in excel output
    """
    letter: str = 'J'
    current_position: int = current_setup.get('body_position')
    worksheet: xlsxwriter = current_setup.get('worksheet')
    structure: xlsxwriter = current_setup.get('structure')

    for index, pipe_name in enumerate(processed_console_server, start=0):
        current_setup['pipes_total'] += 1
        group_unique_tickets: list = processed_console_server.get(pipe_name, {}).get('group_unique_tickets')
        current_color_11: xlsxwriter = get_current_color_11_left(index, structure)

        total_tickets: int = get_total_tickets(pipe_name, processed_console_server)
        max_number: int = current_position + total_tickets - 1

        base_position: str = get_base_position(letter, current_position)
        merge_position: str = get_merge_position(base_position, letter, max_number)

        if len(group_unique_tickets) == 0:
            worksheet.merge_range(f'{base_position}:R{current_position}', '', structure.missing_cell)
            worksheet.merge_range(f'F{current_position}:H{current_position}', 'Console Server - No TRR',
                                  structure.neutral_cell_16)
            current_setup['no_ticket'] += 1

        else:
            group_unique_tickets: list = sorted(group_unique_tickets)
            unique_tickets_content: list = []
            for unique_ticket in group_unique_tickets:
                trr_type: int = azure_devops_data.get(unique_ticket, {}).get('trr_type', '')
                ticket_state: str = clean_ticket_state(azure_devops_data, unique_ticket)

                if trr_type == 1:
                    unique_tickets_content.append('Primary')
                    if 'Signed Off' not in ticket_state and 'Test Completed' not in \
                            ticket_state and 'Done' not in ticket_state:
                        current_setup['primary_count'] += 1

                elif trr_type == 2:
                    unique_tickets_content.append('Secondary')
                    if 'Signed Off' not in ticket_state and 'Test Completed' not in \
                            ticket_state and 'Done' not in ticket_state:
                        current_setup['secondary_count'] += 1

            if len(list(set(unique_tickets_content))) == 1 and len(group_unique_tickets) == 1:
                worksheet.write(base_position, f'  {unique_tickets_content[0]}', current_color_11)

            elif len(list(set(unique_tickets_content))) == 1 and len(group_unique_tickets) >= 2:
                worksheet.merge_range(merge_position, f'  {unique_tickets_content[0]}', current_color_11)

            else:
                for count, ticket_data in enumerate(unique_tickets_content, start=0):
                    modified_position = current_position + count
                    base_position = get_base_position(letter, modified_position)
                    worksheet.write(base_position, f'  {ticket_data}', current_color_11)

        worksheet.set_row(max_number, 3.75, structure.white)
        current_position = max_number + 2

    return current_setup


def clean_pipe_description(console_server_data, pipe_name):
    description: str = console_server_data.get(pipe_name, {}).get('description', 'None')
    return f'   {description}'


def clean_trr_type(console_server_data, pipe_name):
    description: str = console_server_data.get(pipe_name, {}).get('trr_type', 'None')
    return f'   {description}'


def add_issue_data(azure_devops_data: dict, console_server_data: dict, current_setup: dict) -> dict:
    """

    """
    processed_console_server: dict = process_console_server(console_server_data)

    write_pipe_name_column(processed_console_server, console_server_data, current_setup)

    write_description_column(processed_console_server, console_server_data, current_setup)

    current_setup = add_trr_type_column(processed_console_server, azure_devops_data, current_setup)

    current_setup = write_state_column(processed_console_server, azure_devops_data, current_setup)

    write_trr_column(processed_console_server, azure_devops_data, current_setup)

    add_qual_column(processed_console_server, azure_devops_data, current_setup)

    write_assigned_to_column(processed_console_server, azure_devops_data, current_setup)

    write_expected_start_column(processed_console_server, azure_devops_data, current_setup)
    write_expected_end_column(processed_console_server, azure_devops_data, current_setup)
    write_actual_start_column(processed_console_server, azure_devops_data, current_setup)
    write_actual_end_column(processed_console_server, azure_devops_data, current_setup)

    return current_setup


def write_expected_start_column(processed_console_server, azure_devops_data, current_setup):
    """
    Writes Pipe Name column in excel output
    """
    letter: str = 'O'
    current_position: int = current_setup.get('body_position')
    worksheet: xlsxwriter = current_setup.get('worksheet')
    structure: xlsxwriter = current_setup.get('structure')

    for index, pipe_name in enumerate(processed_console_server, start=0):
        group_unique_tickets = list(sorted(processed_console_server.get(pipe_name, {}).get('group_unique_tickets')))
        group_content: list = get_group_expected_start(azure_devops_data, group_unique_tickets)

        current_color_12: xlsxwriter = get_current_color_12_middle(index, structure)
        max_number: int = get_max_number(current_position, pipe_name, processed_console_server)

        for count, unique_ticket in enumerate(group_unique_tickets, start=0):
            base_position: str = get_base_position(letter, current_position + count)
            expected_start = clean_expected_start(azure_devops_data, unique_ticket)

            if len(group_unique_tickets) >= 2 and len(group_content) == 1:
                max_vertical = get_max_vertical(base_position, group_unique_tickets, letter, current_position + count)
                worksheet.merge_range(max_vertical, expected_start, current_color_12)
                break

            elif expected_start == 'None' or not expected_start:
                worksheet.write(base_position, '', structure.missing_cell)

            else:
                worksheet.write(base_position, expected_start, current_color_12)

        current_position = max_number + 2

    return current_setup


def clean_expected_start(azure_devops_data, unique_ticket):
    expected_task_start: str = azure_devops_data.get(unique_ticket, {}).get('due_dates', {}).get('expected_task_start',
                                                                                                 'None')
    return parsed_date(expected_task_start)


def write_actual_start_column(processed_console_server: dict, azure_devops_data: dict, current_setup: dict):
    """
    Accounts for actual start date according to assigned TRR from Microsoft.
    """
    letter: str = 'Q'
    current_position: int = current_setup.get('body_position')
    worksheet: xlsxwriter = current_setup.get('worksheet')
    structure: xlsxwriter = current_setup.get('structure')

    for index, pipe_name in enumerate(processed_console_server, start=0):
        group_unique_tickets = list(sorted(processed_console_server.get(pipe_name, {}).get('group_unique_tickets')))
        group_content: list = get_group_actual_start(azure_devops_data, group_unique_tickets)

        current_color_12: xlsxwriter = get_current_color_12_middle(index, structure)
        max_number: int = get_max_number(current_position, pipe_name, processed_console_server)
        does_not_exist: bool = is_does_not_exist(azure_devops_data, group_unique_tickets)

        for count, unique_ticket in enumerate(group_unique_tickets, start=0):
            base_position: str = get_base_position(letter, current_position + count)
            actual_start: str = clean_actual_start(azure_devops_data, unique_ticket)

            if not does_not_exist:
                worksheet.write(base_position, actual_start, current_color_12)

            elif len(group_unique_tickets) >= 2 and len(group_content) == 1 and not actual_start:
                max_vertical: str = get_max_vertical(base_position, group_unique_tickets, letter,
                                                     current_position + count)
                worksheet.merge_range(max_vertical, '', structure.missing_cell)
                break

            elif len(group_unique_tickets) >= 2 and len(group_content) == 1:
                max_vertical = get_max_vertical(base_position, group_unique_tickets, letter, current_position + count)
                worksheet.merge_range(max_vertical, actual_start, current_color_12)
                break

            elif actual_start == 'None' or not actual_start:
                worksheet.write(base_position, '', structure.missing_cell)

            else:
                worksheet.write(base_position, actual_start, current_color_12)

        current_position = max_number + 2

    return current_setup


def write_actual_end_column(processed_console_server, azure_devops_data, current_setup):
    """
    Writes Pipe Name column in excel output
    """
    letter: str = 'R'
    current_position: int = current_setup.get('body_position')
    worksheet: xlsxwriter = current_setup.get('worksheet')
    structure: xlsxwriter = current_setup.get('structure')

    for index, pipe_name in enumerate(processed_console_server, start=0):
        group_unique_tickets = list(sorted(processed_console_server.get(pipe_name, {}).get('group_unique_tickets')))
        group_content: list = get_group_actual_end(azure_devops_data, group_unique_tickets)

        current_color_12: xlsxwriter = get_current_color_12_middle(index, structure)
        max_number: int = get_max_number(current_position, pipe_name, processed_console_server)
        does_not_exist: bool = is_does_not_exist(azure_devops_data, group_unique_tickets)

        for count, unique_ticket in enumerate(group_unique_tickets, start=0):
            base_position: str = get_base_position(letter, current_position + count)
            actual_end: str = clean_actual_end(azure_devops_data, unique_ticket)
            ticket_state: str = clean_ticket_state(azure_devops_data, unique_ticket)

            if not does_not_exist:
                worksheet.write(base_position, actual_end, current_color_12)

            elif len(group_unique_tickets) >= 2 and len(group_content) == 1 and not actual_end:
                max_vertical = get_max_vertical(base_position, group_unique_tickets, letter, current_position + count)
                worksheet.merge_range(max_vertical, '', structure.missing_cell)
                break

            elif len(group_unique_tickets) >= 2 and len(group_content) == 1:
                max_vertical = get_max_vertical(base_position, group_unique_tickets, letter, current_position + count)
                worksheet.merge_range(max_vertical, actual_end, current_color_12)
                break

            elif actual_end == 'None' or not actual_end:
                worksheet.write(base_position, '', structure.missing_cell)

            else:
                worksheet.write(base_position, actual_end, current_color_12)

        current_position = max_number + 2

    return current_setup


def clean_actual_end(azure_devops_data, unique_ticket):
    actual_task_end: str = azure_devops_data.get(unique_ticket, {}).get('due_dates', {}). \
        get('actual_qual_end_date', 'None')
    return parsed_date(actual_task_end)


def clean_actual_start(azure_devops_data, unique_ticket):
    actual_task_start: str = azure_devops_data.get(unique_ticket, {}).get('due_dates', {}). \
        get('actual_qual_start_date', 'None')
    return parsed_date(actual_task_start)


def write_expected_end_column(processed_console_server, azure_devops_data, current_setup):
    """
    Writes Pipe Name column in excel output
    """
    letter: str = 'P'
    current_position: int = current_setup.get('body_position')
    worksheet: xlsxwriter = current_setup.get('worksheet')
    structure: xlsxwriter = current_setup.get('structure')

    for index, pipe_name in enumerate(processed_console_server, start=0):
        group_unique_tickets = list(sorted(processed_console_server.get(pipe_name, {}).get('group_unique_tickets')))
        group_content: list = get_group_expected_end(azure_devops_data, group_unique_tickets)

        current_color_12: xlsxwriter = get_current_color_12_middle(index, structure)
        max_number: int = get_max_number(current_position, pipe_name, processed_console_server)

        for count, unique_ticket in enumerate(group_unique_tickets, start=0):
            base_position: str = get_base_position(letter, current_position + count)
            expected_end: str = clean_expected_end(azure_devops_data, unique_ticket)

            if len(group_unique_tickets) >= 2 and len(group_content) == 1:
                max_vertical = get_max_vertical(base_position, group_unique_tickets, letter, current_position + count)
                worksheet.merge_range(max_vertical, expected_end, current_color_12)
                break

            elif expected_end == 'None' or not expected_end:
                worksheet.write(base_position, '', structure.missing_cell)

            else:
                worksheet.write(base_position, expected_end, current_color_12)

        current_position = max_number + 2

    return current_setup


def clean_expected_end(azure_devops_data, unique_ticket):
    expected_task_end: str = azure_devops_data.get(unique_ticket, {}).get('due_dates', {}). \
        get('expected_task_completion', 'None')
    return parsed_date(expected_task_end)


def write_assigned_to_column(processed_console_server, azure_devops_data, current_setup):
    """
    Writes Pipe Name column in excel output
    """
    letter: str = 'N'
    current_position: int = current_setup.get('body_position')
    worksheet: xlsxwriter = current_setup.get('worksheet')
    structure: xlsxwriter = current_setup.get('structure')

    for index, pipe_name in enumerate(processed_console_server, start=0):
        group_unique_tickets = list(sorted(processed_console_server.get(pipe_name, {}).get('group_unique_tickets')))
        group_content: list = get_group_unique_assigned_to(azure_devops_data, group_unique_tickets)

        current_color_12: xlsxwriter = get_current_color_12_middle(index, structure)
        max_number: int = get_max_number(current_position, pipe_name, processed_console_server)

        for count, unique_ticket in enumerate(group_unique_tickets, start=0):
            assigned_to = clean_assigned_to(azure_devops_data, unique_ticket)
            base_position: str = get_base_position(letter, current_position + count)

            if len(group_unique_tickets) >= 2 and len(group_content) == 1:
                max_vertical = get_max_vertical(base_position, group_unique_tickets, letter, current_position + count)
                worksheet.merge_range(max_vertical, assigned_to, current_color_12)
                break

            elif assigned_to == 'None':
                worksheet.write(base_position, '', structure.missing_cell)

            else:
                worksheet.write(base_position, assigned_to, current_color_12)

        worksheet.set_row(max_number, 3.75, structure.white)
        current_position = max_number + 2

    return current_setup


def clean_assigned_to(azure_devops_data, unique_ticket):
    assigned_to: str = azure_devops_data.get(unique_ticket, {}).get('assigned_to', 'None')
    assigned_to = assigned_to.replace('@VeritasDCservices.com', '').replace('.', ' ').lower().title().strip()
    return assigned_to


def write_state_column(processed_console_server: dict, azure_devops_data: dict, current_setup: dict) -> dict:
    """
    Adds state of TRR and fills other cells based on state of TRR
    """
    letter: str = 'M'

    current_position: int = current_setup.get('body_position')
    worksheet: xlsxwriter = current_setup.get('worksheet')
    structure: xlsxwriter = current_setup.get('structure')

    for index, pipe_name in enumerate(processed_console_server, start=0):
        group_unique_tickets: list = sorted(processed_console_server.get(pipe_name, {}).get('group_unique_tickets'))
        current_color_14: xlsxwriter = get_current_color_14_middle(index, structure)

        total_tickets: int = get_total_tickets(pipe_name, processed_console_server)
        max_number: int = current_position + total_tickets - 1

        unique_tickets_state: list = get_unique_tickets_state(azure_devops_data, group_unique_tickets)
        unique_tickets_actual_end: list = get_unique_tickets_actual_end(azure_devops_data, group_unique_tickets)

        if len(group_unique_tickets) == 0:
            pass

        else:
            for count, unique_ticket in enumerate(group_unique_tickets, start=0):
                modified_position: int = current_position + count
                ticket_state: str = clean_ticket_state(azure_devops_data, unique_ticket)

                base_position: str = get_base_position(letter, modified_position)
                max_vertical: str = get_max_vertical(base_position, group_unique_tickets, letter, modified_position)
                vertical_range: str = f'{current_position + len(group_unique_tickets) - 1}'

                try:
                    if ticket_state == 'Done' or ticket_state == 'Test Completed':
                        is_review_hold: bool = get_time_difference(azure_devops_data, unique_ticket)

                        if len(group_unique_tickets) >= 2 and len(unique_tickets_state) == 1 and \
                                len(unique_tickets_actual_end) == 1:
                            worksheet.merge_range(max_vertical, ticket_state, structure.purple_middle_12)

                            if is_review_hold and len(unique_tickets_actual_end) == 1:
                                worksheet.merge_range(f'F{modified_position}:H{vertical_range}', 'REVIEW HOLD',
                                                      structure.light_grey_middle_12)
                            elif not is_review_hold and len(unique_tickets_actual_end) == 1:
                                worksheet.merge_range(f'F{modified_position}:H{vertical_range}', 'AVAILABLE',
                                                      structure.dark_grey_middle_14)
                                current_setup['available'] += 1
                            break

                        else:
                            worksheet.write(base_position, ticket_state, structure.purple_middle_12)
                            if is_review_hold:
                                worksheet.merge_range(f'F{modified_position}:H{modified_position}', 'REVIEW HOLD',
                                                      structure.light_grey_middle_12)
                            elif not is_review_hold:
                                worksheet.merge_range(f'F{modified_position}:H{modified_position}', 'AVAILABLE',
                                                      structure.dark_grey_middle_14)

                    elif ticket_state == 'Blocked':
                        current_setup['blocked'] += 1
                        if len(group_unique_tickets) >= 2 and len(unique_tickets_state) == 1:
                            worksheet.merge_range(max_vertical, ticket_state, structure.light_red_middle_12)
                            worksheet.merge_range(f'F{modified_position}:H{vertical_range}', 'BLOCKED',
                                                  structure.light_red_middle_16)
                            break
                        else:
                            worksheet.write(base_position, ticket_state, structure.light_red_middle_12)
                            worksheet.merge_range(f'F{modified_position}:H{modified_position}', 'BLOCKED',
                                                  structure.light_red_middle_12)

                    elif ticket_state == 'On Hold':
                        if len(group_unique_tickets) >= 2 and len(unique_tickets_state) == 1:
                            worksheet.merge_range(max_vertical, ticket_state, structure.light_red_middle_12)
                            worksheet.merge_range(f'F{modified_position}:H{vertical_range}', 'On Hold',
                                                  structure.light_red_middle_16)
                            break
                        else:
                            worksheet.write(base_position, ticket_state, structure.light_red_middle_12)
                            worksheet.merge_range(f'F{modified_position}:H{modified_position}', 'On Hold',
                                                  structure.light_red_middle_12)

                    elif ticket_state == 'Planning':
                        if len(group_unique_tickets) >= 2 and len(unique_tickets_state) == 1:
                            worksheet.merge_range(max_vertical, ticket_state, structure.purple_middle_12)
                            worksheet.merge_range(f'F{modified_position}:H{vertical_range}', 'PLANNING',
                                                  structure.dark_grey_middle_16)
                            break
                        else:
                            worksheet.write(base_position, ticket_state, structure.purple_middle_12)
                            worksheet.merge_range(f'F{modified_position}:H{modified_position}', 'PLANNING',
                                                  structure.dark_grey_middle_12)

                    elif ticket_state == 'In Progress':
                        if len(group_unique_tickets) >= 2 and len(unique_tickets_state) == 1:
                            worksheet.merge_range(max_vertical, ticket_state, structure.aqua_middle_12)
                            worksheet.merge_range(f'F{modified_position}:G{vertical_range}', 'COMPLETE',
                                                  structure.middle_green_14)
                            worksheet.merge_range(f'H{modified_position}:H{vertical_range}', 'PROGRESS',
                                                  current_color_14)
                            break
                        else:
                            worksheet.write(base_position, ticket_state, structure.aqua_middle_12)
                            worksheet.merge_range(f'F{modified_position}:G{modified_position}', 'COMPLETE',
                                                  structure.middle_green_14)
                            worksheet.write(f'H{modified_position}', 'PROGRESS', current_color_14)

                    elif ticket_state == 'Ready to Start':
                        if len(group_unique_tickets) >= 2 and len(unique_tickets_state) == 1:
                            worksheet.merge_range(max_vertical, ticket_state, structure.aqua_middle_12)
                            worksheet.merge_range(f'F{modified_position}:G{vertical_range}', 'COMPLETE',
                                                  structure.middle_green_14)
                            worksheet.merge_range(f'H{modified_position}:H{vertical_range}', 'READY',
                                                  current_color_14)
                            break
                        else:
                            worksheet.write(base_position, ticket_state, structure.aqua_middle_12)
                            worksheet.merge_range(f'F{modified_position}:G{modified_position}', 'COMPLETE',
                                                  structure.middle_green_14)
                            worksheet.write(f'H{modified_position}', 'READY', current_color_14)

                    elif ticket_state == 'Ready to Review':
                        if len(group_unique_tickets) >= 2 and len(unique_tickets_state) == 1:
                            worksheet.merge_range(max_vertical, ticket_state, structure.aqua_middle_12)
                            worksheet.merge_range(f'F{modified_position}:F{vertical_range}', 'PROGRESS',
                                                  current_color_14)
                            worksheet.merge_range(f'G{modified_position}:G{vertical_range}', 'PROGRESS',
                                                  current_color_14)
                            worksheet.merge_range(f'H{modified_position}:H{vertical_range}', '',
                                                  structure.missing_cell)
                            break
                        else:
                            worksheet.write(base_position, ticket_state, structure.aqua_middle_12)
                            worksheet.write(f'F{modified_position}', 'PROGRESS', current_color_14)
                            worksheet.write(f'G{modified_position}', 'PROGRESS', current_color_14)
                            worksheet.write(f'H{modified_position}', '', structure.missing_cell)

                    elif ticket_state == 'Signed Off':
                        if len(group_unique_tickets) >= 2 and len(unique_tickets_state) == 1:
                            worksheet.merge_range(max_vertical, ticket_state, structure.purple_middle_12)
                            worksheet.merge_range(f'F{modified_position}:H{vertical_range}', 'AVAILABLE',
                                                  structure.dark_grey_middle_14)
                            current_setup['available'] += 1

                            break
                        else:
                            worksheet.write(base_position, ticket_state, structure.purple_middle_12)
                            worksheet.merge_range(f'F{modified_position}:H{modified_position}', 'AVAILABLE',
                                                  structure.dark_grey_middle_14)
                            current_setup['available'] += 1

                    elif ticket_state == 'None':
                        worksheet.write(base_position, 'Does Not Exist', structure.light_red_middle_12)
                        worksheet.write(f'L{modified_position}', '', structure.missing_cell)
                        worksheet.merge_range(f'N{modified_position}:R{modified_position}', '',
                                              structure.missing_cell)
                        worksheet.merge_range(f'F{modified_position}:H{modified_position}', '',
                                              structure.missing_cell)

                except AttributeError:
                    worksheet.write(base_position, 'Does Not Exist', structure.light_red_middle_12)
                    worksheet.merge_range(f'F{modified_position}:H{modified_position}', '',
                                          structure.missing_cell)

        worksheet.set_row(max_number, 3.75, structure.white)
        current_position = max_number + 2

    return current_setup


def is_does_not_exist(azure_devops_data, group_unique_tickets) -> bool:
    for current_ticket in group_unique_tickets:
        ticket_state: str = clean_ticket_state(azure_devops_data, current_ticket)
        if 'Does Not Exist' in ticket_state:
            return False
    else:
        return True


def get_unique_tickets_state(azure_devops_data, group_unique_tickets) -> list:
    unique_tickets_content: list = []
    for unique_ticket in group_unique_tickets:
        ticket_state: str = azure_devops_data.get(unique_ticket, {}).get('state', 'None')
        unique_tickets_content.append(ticket_state)
    return list(set(unique_tickets_content))


def get_unique_tickets_actual_end(azure_devops_data, group_unique_tickets) -> list:
    unique_tickets_content: list = []
    for unique_ticket in group_unique_tickets:
        actual_task_end: str = azure_devops_data.get(unique_ticket, {}).get('due_dates', {}). \
            get('actual_qual_end_date', 'None')
        unique_tickets_content.append(actual_task_end)
    return list(set(unique_tickets_content))


def get_time_difference(azure_devops_data, unique_ticket) -> bool:
    """
    Reasons why...
    """
    days: int = get_epoch_difference(azure_devops_data, unique_ticket)

    if days <= 7:
        return True
    else:
        return False


def get_epoch_difference(azure_devops_data, unique_ticket):
    date_time: str = get_date_time(azure_devops_data, unique_ticket)
    pattern: str = '%d.%m.%Y %H:%M:%S'
    previous_epoch = int(mktime(strptime(date_time, pattern)))
    current_epoch = int(time())
    epoch_diff = current_epoch - previous_epoch
    days = int(epoch_diff / 86400)
    return days


def get_date_time(azure_devops_data, unique_ticket) -> str:
    """
    Get date time
    """
    actual_end: str = azure_devops_data.get(unique_ticket, {}).get('due_dates', {}). \
        get('actual_qual_end_date', 'None')
    month: str = actual_end[5:7]
    day: str = actual_end[8:10]
    year: str = actual_end[0:4]
    time_of_day: str = actual_end[11:-1]
    return f'{day}.{month}.{year} {time_of_day}'


def get_max_vertical(base_position, group_unique_tickets, letter, modified_position):
    max_range = str(modified_position + len(group_unique_tickets) - 1)
    return f'{base_position}:{letter}{max_range}'


def clean_ticket_state(azure_devops_data, unique_ticket):
    ticket_state: str = azure_devops_data.get(unique_ticket, {}).get('state', 'None')
    clean_ticket_state: str = ticket_state.replace('InProgress', 'In Progress'). \
        replace('Test completed', 'Test Completed'). \
        replace('Ready To Review', 'Ready to Review'). \
        replace('Ready to start', 'Ready to Start')
    return clean_ticket_state


def add_hyperlink_cell(letter: str, temp_number: int, ticket_url: str, string_data, worksheet, color):
    """

    :param color:
    :param letter:
    :param temp_number:
    :param ticket_url:
    :param string_data:
    :param worksheet:
    :return:
    """
    worksheet.write_url(f'{letter}{temp_number}', ticket_url, color, string=string_data)


def add_normal_cell(base_position, string_data, worksheet, color):
    """

    """
    worksheet.write(base_position, string_data, color)


def add_qual_column(processed_console_server, azure_devops_data, current_setup):
    """
    Writes Pipe Name column in excel output
    """
    letter: str = 'L'
    current_position: int = current_setup.get('body_position')
    worksheet: xlsxwriter = current_setup.get('worksheet')
    structure: xlsxwriter = current_setup.get('structure')

    for index, pipe_name in enumerate(processed_console_server, start=0):
        group_unique_tickets = list(sorted(processed_console_server.get(pipe_name, {}).get('group_unique_tickets')))
        group_content: list = get_group_unique_request_types(azure_devops_data, group_unique_tickets)

        current_color_12: xlsxwriter = get_current_color_12_middle(index, structure)
        max_number: int = get_max_number(current_position, pipe_name, processed_console_server)

        for count, unique_ticket in enumerate(group_unique_tickets, start=0):
            request_type: str = get_request_type(azure_devops_data, unique_ticket)
            base_position: str = get_base_position(letter, current_position + count)

            if len(group_unique_tickets) >= 2 and len(group_content) == 1:
                max_vertical = get_max_vertical(base_position, group_unique_tickets, letter, current_position + count)

                worksheet.merge_range(max_vertical, request_type, current_color_12)
                break
            elif request_type == 'None':
                worksheet.write(base_position, '', structure.missing_cell)
            else:
                worksheet.write(base_position, request_type, current_color_12)

        worksheet.set_row(max_number, 3.75, structure.white)
        current_position = max_number + 2

    return current_setup


def get_group_unique_request_types(azure_devops_data, group_unique_tickets):
    group_content: list = []
    for item in group_unique_tickets:
        request_type: str = get_request_type(azure_devops_data, item).strip()
        group_content.append(request_type)
    return list(sorted(set(group_content)))


def get_group_unique_assigned_to(azure_devops_data, group_unique_tickets):
    group_content: list = []
    for unique_ticket in group_unique_tickets:
        assigned_to: str = azure_devops_data.get(unique_ticket, {}).get('assigned_to')
        if not assigned_to:
            group_content.append('None')
        else:
            group_content.append(assigned_to)
    return list(sorted(set(group_content)))


def get_group_expected_start(azure_devops_data, group_unique_tickets):
    group_content: list = []
    for unique_ticket in group_unique_tickets:
        expected_task_end: str = azure_devops_data.get(unique_ticket, {}).get('due_dates', {}). \
            get('expected_task_start', 'None')
        if not expected_task_end:
            group_content.append('None')
        else:
            group_content.append(expected_task_end)
    return list(sorted(set(group_content)))


def get_group_expected_end(azure_devops_data, group_unique_tickets):
    group_content: list = []
    for unique_ticket in group_unique_tickets:
        expected_task_start: str = azure_devops_data.get(unique_ticket, {}).get('due_dates', {}). \
            get('expected_task_completion', 'None')
        if not expected_task_start:
            group_content.append('None')
        else:
            group_content.append(expected_task_start)
    return list(sorted(set(group_content)))


def get_group_actual_start(azure_devops_data, group_unique_tickets):
    group_content: list = []
    for unique_ticket in group_unique_tickets:
        expected_task_start: str = azure_devops_data.get(unique_ticket, {}).get('due_dates', {}). \
            get('actual_qual_start_date', 'None')
        if not expected_task_start:
            group_content.append('None')
        else:
            group_content.append(expected_task_start)
    return list(sorted(set(group_content)))


def get_group_actual_end(azure_devops_data, group_unique_tickets):
    group_content: list = []
    for unique_ticket in group_unique_tickets:
        actual_end: str = azure_devops_data.get(unique_ticket, {}).get('due_dates', {}). \
            get('actual_qual_end_date', 'None')
        if not actual_end:
            group_content.append('None')
        else:
            group_content.append(actual_end)
    return list(sorted(set(group_content)))


def get_request_type(azure_devops_data, unique_ticket):
    request_type: str = azure_devops_data.get(unique_ticket, {}).get('table_data', {}).get('request_type',
                                                                                           'None')
    return request_type.replace(' TEST', '').replace('TEST', '')


def write_trr_column(processed_console_server, azure_devops_data, current_setup) -> dict:
    """
    Writes Pipe Name column in excel output
    """
    letter: str = 'K'
    current_position: int = current_setup.get('body_position')
    worksheet: xlsxwriter = current_setup.get('worksheet')
    structure: xlsxwriter = current_setup.get('structure')

    for index, pipe_name in enumerate(processed_console_server, start=0):
        group_unique_tickets = list(sorted(processed_console_server.get(pipe_name, {}).get('group_unique_tickets')))
        current_color_11: xlsxwriter = get_current_color_11_middle(index, structure)
        max_number: int = get_max_number(current_position, pipe_name, processed_console_server)

        for count, ticket_data in enumerate(group_unique_tickets, start=0):
            ticket_url: str = f'https://azurecsi.visualstudio.com/CSI%20Commodity%20Qualification/' \
                              f'_workitems/edit/{ticket_data}'
            base_position: str = get_base_position(letter, current_position + count)
            worksheet.write_url(base_position, ticket_url, current_color_11, string=ticket_data)

        worksheet.set_row(max_number, 3.75, structure.white)
        current_position = max_number + 2

    return current_setup


def get_max_number(current_position, pipe_name, processed_console_server) -> int:
    total_tickets: int = get_total_tickets(pipe_name, processed_console_server)
    max_number: int = current_position + total_tickets - 1
    return max_number


def remove_excel_green_corners(current_setup) -> None:
    """
    Excel sometimes have green corners within a cell. Removes to clear up look of excel output.
    :param current_setup: Current worksheet
    """
    worksheet = current_setup.get('worksheet')

    worksheet.ignore_errors({'number_stored_as_text': 'A1:XFD1048576'})


def create_excel_output(basic_data: dict) -> None:
    """
    Create Dashboard here
    """
    username: str = basic_data['username']
    site: str = basic_data['site']
    version: str = basic_data['version']

    current_setup: dict = create_executive_summary(version, username, site)

    console_server_data: dict = get_console_server_data()
    azure_devops_data: dict = get_all_ticket_data(console_server_data)
    pipe_numbers: dict = get_pipe_numbers(console_server_data)

    # final_data: dict = compare_data(console_server_data, azure_devops_data)
    all_issues: list = get_all_issues()

    all_checks: list = get_total_checks()
    missing_tally: str = get_missing_tally()
    mismatch_tally: str = get_mismatch_tally()

    set_sheet_structure(current_setup)

    current_setup: dict = add_issue_data(azure_devops_data, console_server_data, current_setup)

    add_user_info_totals(console_server_data, current_setup)

    remove_excel_green_corners(current_setup)

    create_personal_issues_sheet(azure_devops_data, console_server_data, current_setup, all_issues)

    workbook: xlsxwriter = current_setup.get('workbook')
    structure: xlsxwriter = current_setup.get('structure')

    create_issues_sheet(azure_devops_data, console_server_data, workbook, structure, site, all_issues,
                        all_checks, mismatch_tally, missing_tally, pipe_numbers, version)

    create_setup_sheet(azure_devops_data, console_server_data, workbook, structure, site, all_issues,
                       all_checks, mismatch_tally, missing_tally, pipe_numbers, version)

    create_virtual_machine_sheet(console_server_data, workbook, structure, site, all_issues,
                                 all_checks, mismatch_tally, missing_tally, pipe_numbers, version)

    add_dashboard_inventory(azure_devops_data, console_server_data, current_setup, all_issues)

    add_all_serial(azure_devops_data, console_server_data, current_setup, all_issues)

    add_all_part_numbers(console_server_data, current_setup)

    try:
        workbook.close()
    except xlsxwriter.workbook.FileCreateError:
        print(f'\n\tCannot create excel output because dashboard.xlsx is {Fore.RED}ALREADY OPEN{Style.RESET_ALL}')
        print(f'\tPlease close the dashboard.xlsx to create a new dashboard.xlsx')
        print(f'\tPress {Fore.BLUE}ENTER{Style.RESET_ALL} to try again...', end='')
        input()
        create_excel_output(basic_data)
