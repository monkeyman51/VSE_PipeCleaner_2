"""
9/16/2021 - Data Dump for Aws or others.
"""

from time import strftime

import xlsxwriter


def set_sheet_structure(current_setup: dict) -> None:
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
    add_header_items_under_testing(current_setup)


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
    worksheet: xlsxwriter = current_setup.get('worksheet')
    structure: xlsxwriter = current_setup.get('structure')

    worksheet.write('B5', f'  Console Server - Data Dump', structure.blue_font_22)


def set_excel_design(current_setup: dict) -> None:
    """
    Set up excel output design/parameters.
    """
    set_rows_and_columns_sizes(current_setup)

    add_column_titles(current_setup)
    add_freeze_panes(current_setup)
    add_vse_logo_top_right(current_setup)


def add_header_user_name(current_setup: dict):
    """
    Add clean user name to the top left corner.
    """
    clean_name: str = current_setup.get('clean_name')

    worksheet: xlsxwriter = current_setup.get('worksheet')
    structure: xlsxwriter = current_setup.get('structure')

    worksheet.write('B7', f'            {clean_name}', structure.bold_italic_blue_font)


def get_user_virtual_machines(user_info):
    return user_info['virtual_machines']


def get_user_pipe_total(user_info) -> int:
    user_systems: dict = get_user_systems(user_info)
    user_pipes: list = get_user_pipes(user_systems)
    user_unique_pipes: list = get_user_unique_pipes(user_pipes)
    return len(user_unique_pipes)


def get_user_unique_pipes(user_pipes):
    return sorted(list(set(user_pipes)))


def get_user_unique_pipes(user_pipes):
    return sorted(list(set(user_pipes)))


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


def add_header_items_under_testing(current_setup: dict) -> None:
    """
    These items under testing are meant to be components still not 100% confident
    """
    worksheet: xlsxwriter = current_setup.get('worksheet')
    structure: xlsxwriter = current_setup.get('structure')
    header_height: xlsxwriter = current_setup.get('header_height')
    upper_header: str = header_height - 1


def get_user_pipes(user_systems):
    all_pipes: list = []
    for item in user_systems:
        if 'VSE' in item and '-' in item:
            all_pipes.append(user_systems[item]['pipe_name'])
    return all_pipes


def get_user_systems(user_info) -> dict:
    return user_info['systems']


def get_user_info_alt(console_server_data, default_name) -> dict:
    default_name_underscore: str = default_name_period_to_underscore(default_name)
    # import json
    # foo = json.dumps(console_server_data['user_base'], sort_keys=True, indent=4)
    # print(foo)
    # input()

    try:
        for user_name in console_server_data['user_base']:
            alt_name = str(console_server_data['user_base'][user_name]['alt_name']).lower()

            if default_name.lower() in alt_name or default_name_underscore in user_name:
                return console_server_data['user_base'][user_name]
        else:
            return {}

    except KeyError:
        return {}


def default_name_period_to_underscore(default_name):
    if 'steph' in default_name and '.ak' in default_name:
        return 'steph_ak'
    else:
        return default_name.replace('.', '_').replace('-EXT', '')


def add_vse_logo_top_right(current_setup: dict) -> None:
    """
    Creates VSE Logo on the top left corner
    :param current_setup:
    """
    worksheet: xlsxwriter = current_setup.get('worksheet')

    worksheet.insert_image('A1', 'pipe_cleaner/img/vsei_logo.png')


def clean_pipe_cleaner_version(pipe_cleaner_version) -> str:
    """
    Version for documentation
    :param pipe_cleaner_version:
    :return: cleaner version
    """
    return f"v{pipe_cleaner_version.split(' ')[0]}"


def add_freeze_panes(current_setup: dict) -> None:
    """
    Allows information to the left to stay
    """
    header_height: int = current_setup.get('header_height')
    worksheet: xlsxwriter = current_setup.get('worksheet')

    worksheet.freeze_panes(header_height, 5)


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

    # worksheet.write('G5', 'TOTAL - Active Parts', structure.teal_middle_14)
    # worksheet.write('G6', 'DIMM - Count', structure.teal_middle_12)
    # worksheet.write('G7', 'NVMe - Count', structure.teal_middle_12)
    # worksheet.write('G8', 'Disk - Count', structure.teal_middle_12)
    # worksheet.write('G9', 'Total - Active / Inactive', structure.teal_middle_12)
    # worksheet.write('G10', 'Percentage - Active Parts', structure.teal_middle_12)


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


def add_white_cell(position: str, current_setup) -> None:
    """
    Account for empty cells that don't have column title.
    Meant for giving space between different groups of data.
    """
    worksheet: xlsxwriter = current_setup.get('worksheet')
    structure: xlsxwriter = current_setup.get('structure')

    worksheet.write(position, '', structure.white)


def get_letter_for_column_position(initial: int, left_padding: int) -> str:
    """
    For positioning the column title based on starting point of the left padding.
    :return: letter of excel column
    """
    return convert_index_to_letter(initial + left_padding)


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
    Establishes current worksheet row heights for the header.
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


def convert_index_to_letter(index: int) -> str:
    """

    :param index: Current index due to how many columns we care about in the excel output sheet
    """
    lower_character = chr(ord('a') + index)
    return str(lower_character).upper()


def remove_excel_green_corners(current_setup) -> None:
    """
    Excel sometimes have green corners within a cell. Removes to clear up look of excel output.
    :param current_setup: Current worksheet
    """
    worksheet = current_setup.get('worksheet')

    worksheet.ignore_errors({'number_stored_as_text': 'A1:XFD1048576'})


def create_personal_issues_sheet(excel_setup: dict) -> dict:
    """
    Current excel sheet design to setup the excel tab for data to fill in later.
    """
    workbook: xlsxwriter = excel_setup.get('workbook')

    excel_setup['host_group_hyperlink']: str = 'http://172.30.1.100/console/host_groups.php'
    excel_setup['sheet_title']: str = 'Data Dump'

    excel_setup['worksheet']: xlsxwriter = workbook.add_worksheet(excel_setup.get('sheet_title'))

    excel_setup['rows_height']: tuple = (15.75, 15.75, 15.75, 15.75, 21.00, 15.75, 15.75, 15.75, 15.75, 15.75,
                                         3.75, 3.75, 3.75)

    excel_setup['columns_width']: tuple = (0.50, 0.50, 27.0, 26.0, 18.0, 22.0, 21.0, 10.0, 24.0, 49.00, 18.0, 10.0,
                                           10.0, 10.0, 10.0, 10.0, 10.0, 10.0, 10.0, 10.0, 10.0, 10.0, 10.0, 10.0)

    excel_setup['column_names']: tuple = ('Pipe Number',
                                          'Machine Name',
                                          'Location',
                                          'Assigned To',
                                          'BIOS',
                                          'BMC',
                                          'Host IP',
                                          'DHCP IP',
                                          'TRR')

    return excel_setup


def get_last_active(part_number_data):
    last_found_alive: float = part_number_data['last_found_alive']
    days = last_found_alive / 86400.00
    # TODO
    if days < 1:
        # return 'Less than 1 Day'
        return '1'
    else:
        first_part = str(days).split('.')[0]
        if first_part == '1':
            # return f'{first_part} day last online'
            return f'{first_part}'
        else:
            return f'{first_part}'


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


def add_all_serial_data(console_server_data: dict, current_setup: dict):
    """
    Writes Pipe Name column in excel output
    """
    machines_data: list = get_all_machines_data(console_server_data)

    worksheet: xlsxwriter = current_setup.get('worksheet')
    structure: xlsxwriter = current_setup.get('structure')
    initial_position: int = current_setup.get('body_position')

    for index, machine_data in enumerate(machines_data, start=0):
        row_color: xlsxwriter = get_row_color(index, structure)
        current_position: int = index + initial_position

        worksheet.write(f'C{current_position}', machine_data["pipe_number"], row_color)
        worksheet.write(f'D{current_position}', machine_data["machine_name"], row_color)
        worksheet.write(f'E{current_position}', machine_data["location"], row_color)
        worksheet.write(f'F{current_position}', machine_data["assigned_to"], row_color)
        worksheet.write(f'G{current_position}', machine_data["bios"], row_color)
        worksheet.write(f'H{current_position}', machine_data["bmc"], row_color)
        worksheet.write(f'I{current_position}', machine_data["host_ip"], row_color)
        worksheet.write(f'J{current_position}', machine_data["dhcp_ip"], row_color)
        worksheet.write(f'K{current_position}', machine_data["ticket"], row_color)

        worksheet.set_row(current_position - 1, 18.00)


def get_dhcp_ip(console_server_data: dict, key_name: str) -> str:
    """
    Get correct DHCP IP based on naming convention.
    :param console_server_data:
    :param key_name:
    :return:
    """
    short_pipe_name = key_name[-7:-3]
    # print(f"short_pipe_name: {short_pipe_name}")
    dhcp_data: list = console_server_data.get('dhcp_data')

    possible_dhcp: list = []
    for dhcp in dhcp_data:
        dhcp_name: str = dhcp.get('name')
        if short_pipe_name in dhcp_name:
            dhcp_ip: str = dhcp.get('ip')
            possible_dhcp.append(dhcp_ip)

    if len(possible_dhcp) >= 2:
        return ", ".join(possible_dhcp)
    elif len(possible_dhcp) == 0:
        return "None"
    else:
        return str(possible_dhcp[0])


def get_clean_ticket(pipe_data: dict, machine_name: str) -> str:
    """
    Get clean ticket data.  Placed None if none.
    :param pipe_data:
    :param machine_name:
    :return:
    """
    ticket: str = pipe_data[machine_name]["ticket"]
    if ticket.upper() == "NONE" or ticket is None or ticket == "":
        return "None"
    else:
        return ticket


def get_machine_name(machine_name: str) -> str:
    """

    :param machine_name:
    :return:
    """
    machine_name: str = machine_name.strip().upper()

    if "-VM-" in machine_name:
        return "None"
    else:
        return machine_name


def get_assigned_to(pipe_data: dict, machine_name: str) -> str:
    """

    :param pipe_data:
    :param machine_name:
    :return:
    """
    checked_out_to: str = pipe_data[machine_name]["checked_out_to"]

    if checked_out_to == "None" or checked_out_to == "" or checked_out_to is None:
        return "None"
    else:
        return checked_out_to.replace(".", " ").title()


def get_all_machines_data(console_server_data: dict) -> list:
    """
    Console Server data
    :param console_server_data:
    :return:
    """
    all_machines: list = []
    for key_name in console_server_data:

        dhcp_ip: str = get_dhcp_ip(console_server_data, key_name)

        if "PIPE" in key_name.upper():
            pipe_data: dict = console_server_data[key_name]["pipe_data"]

            for machine_name in pipe_data:
                if "VSE" in machine_name:
                    clean_machine_name: str = get_machine_name(machine_name)
                    ticket: str = get_clean_ticket(pipe_data, machine_name)

                    if clean_machine_name != "None":
                        machine_data: dict = {"pipe_number": process_pipe_name(key_name),
                                              "machine_name": clean_machine_name,
                                              "location": pipe_data[machine_name]["location"],
                                              "assigned_to": get_assigned_to(pipe_data, machine_name),
                                              "bios": pipe_data[machine_name]["server_bios"],
                                              "bmc": pipe_data[machine_name]["server_bmc"],
                                              "host_ip": pipe_data[machine_name]["host_ip"],
                                              "dhcp_ip": dhcp_ip,
                                              "ticket": ticket}

                        all_machines.append(machine_data)

    return all_machines


def get_row_color(index: int, structure: xlsxwriter) -> xlsxwriter:
    """
    Contrast excel row colors to improve readability.
    :param index:
    :param structure:
    :return:
    """
    result: int = index % 2

    if result == 0:
        return structure.blue_middle

    elif result == 1:
        return structure.alt_blue_middle


def main_method(console_server_data: dict, excel_setup: dict) -> None:
    """
    Data dump.
    """
    current_setup: dict = create_personal_issues_sheet(excel_setup)

    set_sheet_structure(current_setup)

    add_all_serial_data(console_server_data, current_setup)

    remove_excel_green_corners(current_setup)
