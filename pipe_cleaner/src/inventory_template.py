"""
Module for creating an excel report on Host Groups page information from Console Server. Will grab other information
from ADO (Azure DevOps), VSE (Veritas Services & Engineers) files to create a more comprehensive report.
"""
import os
import sys
from time import strftime

import xlsxwriter
from colorama import Fore, Style

from pipe_cleaner.src.dashboard_write import main_method as write_column_data
from pipe_cleaner.src.dashboard_write import parsed_date
from pipe_cleaner.src.data_access import write_host_groups_json
from pipe_cleaner.src.excel_properties import Structure


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
    add_header_date_and_version(current_setup)


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


def get_user_pipes(user_systems):
    all_pipes: list = []
    for item in user_systems:
        if 'VSE' in item and '-' in item:
            all_pipes.append(user_systems[item]['pipe_name'])
    return all_pipes


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

    worksheet.write('C6', f' {current_date} - {current_time} - {pipe_cleaner_version}',
                    structure.italic_blue_font)


def clean_pipe_cleaner_version(pipe_cleaner_version) -> str:
    """
    Version for documentation
    :param pipe_cleaner_version:
    :return: cleaner version
    """
    return f"v{pipe_cleaner_version.split(' ')[0]}"


def add_header_sheet_title(current_setup: dict) -> None:
    """
    Adds the excel sheet name to the header area in the top left corner
    """
    sheet_name: str = current_setup.get('sheet_name')

    worksheet: xlsxwriter = current_setup.get('worksheet')
    structure: xlsxwriter = current_setup.get('structure')


def add_header_user_name(current_setup: dict):
    """
    Add clean user name to the top left corner.
    """
    clean_name: str = current_setup.get('clean_name')

    worksheet: xlsxwriter = current_setup.get('worksheet')
    structure: xlsxwriter = current_setup.get('structure')

    worksheet.write('C5', 'Inventory Transaction', structure.black_35)
    worksheet.merge_range('C8:F8', 'LOGISTICS', structure.grey_middle_11)
    worksheet.merge_range('G8:J8', 'NECESSARY', structure.grey_middle_11)
    worksheet.merge_range('K8:N8', 'ADDITIONAL', structure.grey_middle_11)


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

    for index, column_title in enumerate(column_names, start=0):
        position: str = get_column_title_position(header_height, index, left_padding)

        if not column_title:
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

    worksheet.write(position, column_title, structure.teal_middle_16)


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
    structure = current_setup['structure']

    worksheet.insert_image('A1', 'pipe_cleaner/img/vse_logo.png')
    worksheet.write('F5', 'Task Name:', structure.teal_middle_24)
    worksheet.write('G5', 'Copy Task Title Here…', structure.black_left_18)


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


def create_inventory_transaction(basic_data: dict) -> dict:
    """
    Current excel sheet design to setup the excel tab for data to fill in later.
    """
    sheet_title: str = 'Inventory Template'

    site: str = basic_data['site']
    version: str = basic_data['version']
    username: str = basic_data['username']

    check_opened_pipe_cleaner()

    workbook = xlsxwriter.Workbook('inventory_transaction.xlsx')

    excel_setup: dict = {'sheet_title': sheet_title,
                         'worksheet': workbook.add_worksheet(sheet_title),
                         'structure': Structure(workbook),
                         'workbook': workbook,
                         'version': version,
                         'site_location': site,
                         'default_user_name': username,
                         'clean_name': get_clean_user_name(username),
                         'header_height': 9,
                         'body_position': 15,
                         'left_padding': 2,
                         'freeze_pane_position': 6,
                         'host_group_hyperlink': 'http://172.30.1.100/console/host_groups.php',
                         'rows_height': (15.0, 15.0, 15.0, 15.0, 36.0, 15.75, 15.75, 15.75, 30.0, 60.0, 60.0, 60.0,
                                         60.0, 60.0, 60.0, 60.0, 60.0, 60.0, 60.0, 60.0, 60.0, 60.0),
                         'columns_width': (2.00, 2.00, 17.0, 23.0, 22.0, 22.0, 15.0, 30.0, 12.0, 18.0, 17.0, 17.0,
                                           17.0, 40.0, 21.0, 21.0, 21.0, 21.0, 25.0, 25.0, 25.0, 25.0),
                         'column_names': ('Date',
                                          'Name',
                                          'From',
                                          'To',
                                          'Type',
                                          'Part Number',
                                          'QTY',
                                          'Supplier',
                                          'Pipe #',
                                          'Cage #',
                                          'PO #',
                                          'Notes')}

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
        print(f'\tPress {Fore.BLUE}enter{Style.RESET_ALL} to close program.\n')
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


def remove_excel_green_corners(current_setup) -> None:
    """
    Excel sometimes have green corners within a cell. Removes to clear up look of excel output.
    :param current_setup: Current worksheet
    """
    worksheet = current_setup.get('worksheet')

    worksheet.ignore_errors({'number_stored_as_text': 'A1:XFD1048576'})


def add_default_data(current_setup, structure) -> None:
    worksheet: xlsxwriter = current_setup['worksheet']
    header_height: xlsxwriter = current_setup['header_height']

    current_date: str = strftime('%m/%d/%Y')
    items: list = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12']

    from_to_locations: list = ['Cage', 'Inbound', 'Outbound', 'Rack / Pipe', 'Mini-labs', 'Quarantined', 'Reserved',
                               'TurboCats', 'Taking Pictures', 'Other']

    commodity_types: list = ['DIMM', 'SSD', 'HDD', 'NVMe', 'M.2', 'Ruler', 'Disk', 'Other']

    suppliers: list = ['Samsung', 'SK Hynix', 'Micron', 'Seagate', 'HGST/Western Digital', 'Intel', 'Toshiba',
                       'Lite-On', 'Nanya', 'Aspen/Western Digital', 'Kingston', 'Toshiba/Kioxia']

    for index, item in enumerate(items, start=1):
        number = index + header_height

        clean_name = str(current_setup['default_user_name']).replace('.', ' ').replace('-EXT', '').title()

        worksheet.write(f'C{number}', current_date, structure.blank_12)
        worksheet.write(f'D{number}', clean_name, structure.blank_12)
        worksheet.write(f'E{number}', 'Fill Here', structure.blank_12)
        worksheet.write(f'F{number}', 'Fill Here', structure.blank_12)

        worksheet.data_validation(f'E{number}', {'validate': 'list', 'source': from_to_locations})
        worksheet.data_validation(f'F{number}', {'validate': 'list', 'source': from_to_locations})

        worksheet.write(f'G{number}', 'Fill Here', structure.blank_12)
        worksheet.data_validation(f'G{number}', {'validate': 'list', 'source': commodity_types})

        worksheet.write(f'H{number}', '', structure.blank_12)
        worksheet.write(f'I{number}', '', structure.blank_12)

        worksheet.write(f'J{number}', 'Fill Here', structure.blank_12)
        worksheet.data_validation(f'J{number}', {'validate': 'list', 'source': suppliers})

        worksheet.write(f'K{number}', '', structure.blank_12)
        worksheet.write(f'L{number}', '', structure.blank_12)
        worksheet.write(f'M{number}', '', structure.blank_12)
        worksheet.write(f'N{number}', '', structure.blank_12)


def create_inventory_template(basic_data: dict) -> None:
    """
    Create Dashboard here
    """
    current_setup: dict = create_inventory_transaction(basic_data)

    set_sheet_structure(current_setup)

    remove_excel_green_corners(current_setup)

    workbook: xlsxwriter = current_setup.get('workbook')
    structure: xlsxwriter = current_setup.get('structure')

    add_default_data(current_setup, structure)

    try:
        workbook.close()
    except xlsxwriter.workbook.FileCreateError:
        print(f'\n\tCannot create excel output because dashboard.xlsx is {Fore.RED}ALREADY OPEN{Style.RESET_ALL}')
        print(f'\tPlease close the dashboard.xlsx to create a new dashboard.xlsx')
        print(f'\tPress {Fore.BLUE}ENTER{Style.RESET_ALL} to try again...', end='')
        input()
        create_inventory_template()
