import getpass
from time import strftime

import xlsxwriter
from pipe_cleaner.src.data_console_server import get_total_systems
from collections import OrderedDict

pipe_information: dict = {}


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


def process_issues_length(all_issues: list) -> dict:
    """
    Get the number of machine names and pipes names
    For merging cells later
    :param all_issues:
    :return:
    """
    issues_length: dict = {}
    for issue in all_issues:
        pipe_name = issue.get('pipe_name', 'None')
        machine_name = issue.get('machine_name', 'None')
        if pipe_name not in issues_length:
            issues_length[pipe_name] = 0

        if machine_name not in issues_length:
            issues_length[machine_name] = 0
        issues_length[pipe_name] += 1
        issues_length[machine_name] += 1

    return issues_length


def process_color(check_color: int, type_setting: str, structure):
    """

    :param check_color:
    :param type_setting:
    :param structure:
    :return:
    """
    if type_setting.upper() == 'MIDDLE':
        if check_color == 0:
            return structure.blue_middle
        elif check_color == 1:
            return structure.alt_blue_middle

    elif type_setting.upper() == 'LEFT':
        if check_color == 0:
            return structure.blue_left
        elif check_color == 1:
            return structure.alt_blue_left


def check_missing(data: str) -> str:
    """

    :param data:
    :return:
    """
    if data == 'None' or data == '' or data is None or not data or data.upper() == 'NONE':
        return 'None'
    else:
        return data


def get_hyperlinks(ticket_id: str, host_group_id: str, host_id: str, connection_status: str) -> dict:
    """
    Gather URL for hyperlinks in the excel output later for issues page.
    :param connection_status:
    :param ticket_id: TRR Number
    :param host_group_id: Pipe ID within Console Server associated as Host Group
    :param host_id: Individual Host ID known also as a machine or system
    :return:
    """
    ticket_url: str = f'https://azurecsi.visualstudio.com/' \
                      f'CSI%20Commodity%20Qualification/_workitems/edit/{ticket_id}'
    host_group_url: str = f'http://172.30.1.100/console/host_group_host_list.' \
                          f'php?host_group_id={host_group_id}'
    host_url: str = f'http://172.30.1.100/console/host_details.php?host_id={host_id}'

    return {'ticket_url': ticket_url, 'host_group_url': host_group_url, 'host_url': host_url,
            'connection_status': connection_status}


def get_unique_pipes(all_issues: list):
    """

    :param all_issues:
    :return:
    """
    all_pipes: list = []
    for issue in all_issues:
        pipe_name: str = issue.get('pipe_name')
        all_pipes.append(pipe_name)
    return sorted(list(set(all_pipes)))


def process_all_issues(all_issues: list):
    """
    Organize all issues into pipes and machine names dictionaries for easier parsing later.
    :param all_issues:
    :return:
    """
    new_all_issues: dict = get_new_all_issues(all_issues)

    for issue in all_issues:
        pipe_name: str = issue.get('pipe_name')
        machine_name: str = issue.get('machine_name')
        issue_state: str = issue.get('issue_state')
        system_component = str(issue.get('system_component')).lower().replace(' ', '_')
        if pipe_name in new_all_issues:
            new_all_issues[pipe_name][f'{machine_name}_{issue_state}_{system_component}'] = issue

    return new_all_issues


def get_new_all_issues(all_issues):
    new_all_issues: dict = {}
    unique_pipes: list = get_unique_pipes(all_issues)
    for unique_pipe in unique_pipes:
        new_all_issues[unique_pipe] = {}
    return new_all_issues


def write_pipe_name_column(letter: str, min_number: int, processed_issues, worksheet, structure,
                           console_server_data: dict) -> None:
    """
    Writes Pipe Name column in excel output
    :param letter: which column in excel
    :param min_number: starting point of writing column pipe names
    :param console_server_data: for host group IDs
    :param worksheet: which worksheet
    :param structure: color
    :param processed_issues: organized data
    """
    next_min_number: int = min_number

    for pipe_name in processed_issues:
        host_group_url: str = get_host_group_url(console_server_data, pipe_name)
        clean_pipe_name: str = process_pipe_name(pipe_name)
        total_issues: int = get_total_issues(pipe_name, processed_issues)

        pipe_section: int = next_min_number
        pipe_order: int = 0
        next_load: int = 0

        max_number: int = pipe_section + pipe_order + total_issues - 1
        minimal_number: int = pipe_order + pipe_section

        if total_issues == 1:
            worksheet.write_url(f'{letter}{minimal_number}', host_group_url, structure.blue_middle_24,
                                string=clean_pipe_name)
            next_load += 1

        elif total_issues >= 2:
            worksheet.merge_range(f'{letter}{minimal_number}:{letter}{max_number}',
                                  clean_pipe_name, structure.blue_middle_24)

            worksheet.write_url(f'{letter}{minimal_number}', host_group_url, structure.blue_middle_24,
                                string=clean_pipe_name)

            next_load += total_issues
            pipe_order += total_issues

        next_min_number: int = pipe_section + next_load + 1


def get_total_issues(pipe_name, processed_issues):
    total_issues: int = len(processed_issues[pipe_name])
    if total_issues == 1 or total_issues == 0:
        return 1
    else:
        return total_issues


def get_host_group_url(console_server_data, pipe_name):
    host_group_id: str = console_server_data.get(pipe_name, {}).get('host_id', 'None')
    return f'http://172.30.1.100/console/host_group_host_list.php?host_group_id={host_group_id}'


def write_hyperlink_cells(number_of_machines: int, letter: str, number_of_cell: str, max_number: int, hyperlink: str,
                          text: str, worksheet, structure, color):
    """
    Write uniform colors for excel output. Contains merge and hyperlinks
    :param structure:
    :param number_of_machines:
    :param letter:
    :param number_of_cell:
    :param max_number:
    :param hyperlink:
    :param color:
    :param text:
    :param worksheet:
    :return:
    """
    if check_missing(text) == 'None' and number_of_machines == 1:
        worksheet.write(f'{letter}{number_of_cell}', '', structure.missing_cell)

    elif check_missing(text) == 'None' and number_of_machines >= 2:
        worksheet.merge_range(f'{letter}{number_of_cell}:{letter}{max_number}',
                              '', structure.missing_cell)

    elif number_of_machines == 1:
        worksheet.write_url(f'{letter}{number_of_cell}', hyperlink, color,
                            string=text)

    elif number_of_machines >= 2:
        worksheet.merge_range(f'{letter}{number_of_cell}:{letter}{max_number}',
                              text, color)

        worksheet.write_url(f'{letter}{number_of_cell}', hyperlink, color,
                            string=text)


def write_machine_name_column(letter: str, minimal_header_number: int, pipe_structure: dict,
                              worksheet, structure, console_server_data: dict):
    """
    Write host group column in excel output

    :param letter: location of cell
    :param minimal_header_number: starting point of writing column pipe names
    :param worksheet:
    :param structure:
    :param pipe_structure: organized data
    :param console_server_data: for host group IDs
    :return:
    """
    next_min_number: int = minimal_header_number
    for index, pipe_name in enumerate(pipe_structure):
        all_unique_machines = pipe_structure.get(pipe_name, {}).get('unique_machines', 'None')

        pipe_order: int = 0
        end_pipe_section: int = next_min_number
        for start, machine_name in enumerate(all_unique_machines):

            # Fetches information
            host_id: str = console_server_data.get(pipe_name, {}). \
                get('pipe_data', {}).get(machine_name, {}).get('id', 'None')
            host_url: str = f'http://172.30.1.100/console/host_details.php?host_id={host_id}'
            connection_status: str = console_server_data.get(pipe_name, {}). \
                get('pipe_data', {}).get(machine_name, {}).get('connection_status', 'None')

            # Information for positioning
            number_of_machines: int = all_unique_machines.get(machine_name)
            max_number: int = end_pipe_section + number_of_machines - 1
            number_of_cell = str(pipe_order + next_min_number)

            # Console Server considers blade dead after 10 minutes
            if connection_status.upper() == 'DEAD':
                write_hyperlink_cells(number_of_machines, letter, number_of_cell, max_number, host_url,
                                      machine_name, worksheet, structure, structure.dark_grey_middle)

            # Note that mostly dead means dead for less than 10 minutes
            elif connection_status.upper() == 'MOSTLY_DEAD':
                write_hyperlink_cells(number_of_machines, letter, number_of_cell, max_number, host_url,
                                      machine_name, worksheet, structure, structure.light_grey_middle)

            elif connection_status.upper() == 'ALIVE':
                write_hyperlink_cells(number_of_machines, letter, number_of_cell, max_number, host_url,
                                      machine_name, worksheet, structure, structure.blue_middle)

            else:
                # In case connection status is other than dead, alive, or mostly dead
                write_hyperlink_cells(number_of_machines, letter, number_of_cell, max_number, host_url,
                                      machine_name, worksheet, structure, structure.bad_cell)

            # Increments
            pipe_order += number_of_machines
            end_pipe_section += number_of_machines

        next_min_number: int = end_pipe_section + 1


def write_sku_column(letter: str, pipe_structure: dict, min_number: int, worksheet, structure):
    """
    Write Pipe Name column in excel output
    :param min_number: starting point of writing column pipe names
    :param worksheet:
    :param letter:
    :param structure:
    :param pipe_structure: organized data
    :return:
    """
    # Initial minimum is the first data
    next_min_number: int = min_number
    for index, pipe_name in enumerate(pipe_structure):
        sku_type_issues: dict = pipe_structure.get(pipe_name, {}).get('sku_type_issues', 'None')

        # pipe section
        pipe_section: int = next_min_number
        pipe_order: int = 0
        next_load: int = 0

        # Adds to the pipe section
        for start, sku_type in enumerate(sku_type_issues):
            # number of issues associated to a SKU ie. UTL, XDR, VIZ within the pipe
            sku_issues: int = sku_type_issues.get(sku_type, 1)
            max_number: int = pipe_section + pipe_order + sku_issues - 1
            minimal_number: int = pipe_order + pipe_section

            if sku_issues == 1:
                worksheet.write_url(f'{letter}{minimal_number}', '', structure.blue_middle,
                                    string=sku_type)

            elif sku_issues >= 2:
                worksheet.merge_range(f'{letter}{minimal_number}:{letter}{max_number}',
                                      sku_type, structure.blue_middle_gigantic)

            next_load += sku_issues
            pipe_order += sku_issues

        next_min_number: int = pipe_section + next_load + 1


def write_trr_column(letter: str, min_number: int, pipe_structure: dict,
                     worksheet, structure, console_server_data: dict):
    """
    Write Pipe Name column in excel output
    :param min_number: starting point of writing column pipe names
    :param console_server_data: for host group IDs
    :param worksheet:
    :param letter:
    :param structure:
    :param pipe_structure: organized data
    :return:
    """
    next_min_number: int = min_number
    for index, pipe_name in enumerate(pipe_structure):
        all_unique_machines = pipe_structure.get(pipe_name, {}).get('unique_machines', 'None')

        end_section: int = next_min_number
        pipe_order: int = 0

        for start, machine_name in enumerate(all_unique_machines):
            ticket_id: str = console_server_data.get(pipe_name, {}). \
                get('pipe_data', {}).get(machine_name, {}).get('ticket', 'None')
            ticket_url: str = f'https://azurecsi.visualstudio.com/' \
                              f'CSI%20Commodity%20Qualification/_workitems/edit/{ticket_id}'

            number_of_machines: int = all_unique_machines.get(machine_name)
            max_number: int = end_section + number_of_machines - 1

            number_of_cell = str(pipe_order + next_min_number)
            pipe_order += number_of_machines

            write_hyperlink_cells(number_of_machines, letter, number_of_cell, max_number, ticket_url,
                                  ticket_id, worksheet, structure, structure.blue_middle)

            end_section += number_of_machines

        next_min_number: int = end_section + 1


def write_type_column(letter: str, min_number: int, pipe_structure: dict, ado_data,
                      worksheet, structure, console_server_data: dict):
    """
    Write Pipe Name column in excel output
    :param ado_data:
    :param min_number: starting point of writing column pipe names
    :param console_server_data: for host group IDs
    :param worksheet:
    :param letter:
    :param structure:
    :param pipe_structure: organized data
    :return:
    """
    next_min_number: int = min_number
    for index, pipe_name in enumerate(pipe_structure):
        all_unique_machines = pipe_structure.get(pipe_name, {}).get('unique_machines', 'None')

        end_section: int = next_min_number
        pipe_order: int = 0

        for start, machine_name in enumerate(all_unique_machines):

            ticket_id: str = console_server_data.get(pipe_name, {}). \
                get('pipe_data', {}).get(machine_name, {}).get('ticket', 'None')
            ticket_url: str = f'https://azurecsi.visualstudio.com/' \
                              f'CSI%20Commodity%20Qualification/_workitems/edit/{ticket_id}'
            request_type: str = ado_data.get(ticket_id, {}).get('table_data', {}).get('request_type')
            try:
                request_type = request_type.replace(' TEST', '').replace('TEST', '')

                number_of_machines: int = all_unique_machines.get(machine_name)
                max_number: int = end_section + number_of_machines - 1

                number_of_cell = str(pipe_order + next_min_number)
                pipe_order += number_of_machines

                write_hyperlink_cells(number_of_machines, letter, number_of_cell, max_number, ticket_url,
                                      request_type, worksheet, structure, structure.blue_middle)

            except AttributeError:
                number_of_machines: int = all_unique_machines.get(machine_name)
                max_number: int = end_section + number_of_machines - 1

                number_of_cell = str(pipe_order + next_min_number)
                pipe_order += number_of_machines

                write_hyperlink_cells(number_of_machines, letter, number_of_cell, max_number, ticket_url,
                                      'None', worksheet, structure, structure.bad_cell)
            except TypeError:
                number_of_machines: int = all_unique_machines.get(machine_name)
                max_number: int = end_section + number_of_machines - 1

                number_of_cell = str(pipe_order + next_min_number)
                pipe_order += number_of_machines

                write_hyperlink_cells(number_of_machines, letter, number_of_cell, max_number, ticket_url,
                                      'None', worksheet, structure, structure.bad_cell)

            end_section += number_of_machines

        next_min_number: int = end_section + 1


def write_tests_column(letter: str, min_number: int, pipe_structure: dict, ado_data,
                       worksheet, structure, console_server_data: dict):
    """
    Write Pipe Name column in excel output
    :param ado_data:
    :param min_number: starting point of writing column pipe names
    :param console_server_data: for host group IDs
    :param worksheet:
    :param letter:
    :param structure:
    :param pipe_structure: organized data
    :return:
    """
    next_min_number: int = min_number
    for index, pipe_name in enumerate(pipe_structure):
        all_unique_machines = pipe_structure.get(pipe_name, {}).get('unique_machines', 'None')

        end_section: int = next_min_number
        pipe_order: int = 0

        for start, machine_name in enumerate(all_unique_machines):
            ticket_id: str = console_server_data.get(pipe_name, {}). \
                get('pipe_data', {}).get(machine_name, {}).get('ticket', 'None')
            test_plan_hyperlink: str = ado_data.get(ticket_id, {}).get('test_plan_hyperlink', 'None')

            number_of_machines: int = all_unique_machines.get(machine_name)
            max_number: int = end_section + number_of_machines - 1

            number_of_cell = str(pipe_order + next_min_number)
            pipe_order += number_of_machines

            write_hyperlink_cells(number_of_machines, letter, number_of_cell, max_number, test_plan_hyperlink,
                                  'Test Plan', worksheet, structure, structure.blue_middle)

            end_section += number_of_machines

        next_min_number: int = end_section + 1


def write_state_column(letter: str, min_number: int, pipe_structure: dict, ado_data: dict,
                       worksheet, structure, console_server_data: dict):
    """
    Write Pipe Name column in excel output
    :param ado_data:
    :param min_number: starting point of writing column pipe names
    :param console_server_data: for host group IDs
    :param worksheet:
    :param letter:
    :param structure:
    :param pipe_structure: organized data
    :return:
    """
    next_min_number: int = min_number
    for index, pipe_name in enumerate(pipe_structure):
        all_unique_machines = pipe_structure.get(pipe_name, {}).get('unique_machines', 'None')

        end_section: int = next_min_number
        pipe_order: int = 0

        for start, machine_name in enumerate(all_unique_machines):
            ticket_id: str = console_server_data.get(pipe_name, {}). \
                get('pipe_data', {}).get(machine_name, {}).get('ticket', 'None')
            ticket_url: str = f'https://azurecsi.visualstudio.com/' \
                              f'CSI%20Commodity%20Qualification/_workitems/edit/{ticket_id}'
            ticket_state: str = ado_data.get(ticket_id, {}).get('state', {})

            # ticket state text sometimes sloppy and inconsistent
            ticket_state: str = ticket_state.replace('InProgress', 'In Progress'). \
                replace('Test completed', 'Test Completed'). \
                replace('Ready To Review', 'Ready to Review'). \
                replace('Ready to start', 'Ready to Start')

            number_of_machines: int = all_unique_machines.get(machine_name)
            max_number: int = end_section + number_of_machines - 1

            number_of_cell = str(pipe_order + next_min_number)
            pipe_order += number_of_machines

            if ticket_state == 'Done':
                write_hyperlink_cells(number_of_machines, letter, number_of_cell, max_number, ticket_url,
                                      ticket_state, worksheet, structure, structure.purple_middle)
            elif ticket_state == 'Blocked':
                write_hyperlink_cells(number_of_machines, letter, number_of_cell, max_number, ticket_url,
                                      ticket_state, worksheet, structure, structure.bad_cell)
            elif ticket_state == 'On Hold':
                write_hyperlink_cells(number_of_machines, letter, number_of_cell, max_number, ticket_url,
                                      ticket_state, worksheet, structure, structure.bad_cell)
            elif ticket_state == 'Test Completed':
                write_hyperlink_cells(number_of_machines, letter, number_of_cell, max_number, ticket_url,
                                      ticket_state, worksheet, structure, structure.purple_middle)
            elif ticket_state == 'Planning':
                write_hyperlink_cells(number_of_machines, letter, number_of_cell, max_number, ticket_url,
                                      ticket_state, worksheet, structure, structure.purple_middle)

            elif ticket_state == 'In Progress':
                write_hyperlink_cells(number_of_machines, letter, number_of_cell, max_number, ticket_url,
                                      ticket_state, worksheet, structure, structure.aqua_middle)

            elif ticket_state == 'Ready to Start':
                write_hyperlink_cells(number_of_machines, letter, number_of_cell, max_number, ticket_url,
                                      ticket_state, worksheet, structure, structure.aqua_middle)

            elif ticket_state == 'Ready to Review':
                write_hyperlink_cells(number_of_machines, letter, number_of_cell, max_number, ticket_url,
                                      ticket_state, worksheet, structure, structure.aqua_middle)

            else:
                write_hyperlink_cells(number_of_machines, letter, number_of_cell, max_number, ticket_url,
                                      ticket_state, worksheet, structure, structure.blue_middle)

            end_section += number_of_machines

        next_min_number: int = end_section + 1


def write_skudoc_column(letter: str, min_number: int, pipe_structure: dict, ado_data,
                        worksheet, structure, console_server_data: dict):
    """
    Write Pipe Name column in excel output
    :param min_number: starting point of writing column pipe names
    :param console_server_data: for host group IDs
    :param worksheet:
    :param letter:
    :param structure:
    :param pipe_structure: organized data
    :return:
    """
    next_min_number: int = min_number
    for index, pipe_name in enumerate(pipe_structure):
        all_unique_machines = pipe_structure.get(pipe_name, {}).get('unique_machines', 'None')

        end_section: int = next_min_number
        pipe_order: int = 0

        for start, machine_name in enumerate(all_unique_machines):
            ticket_id: str = console_server_data.get(pipe_name, {}). \
                get('pipe_data', {}).get(machine_name, {}).get('ticket', 'None')
            crd_path: str = ado_data.get(ticket_id, {}).get('attachment_file_paths', {}). \
                get('skudoc_drive_path', 'None')

            file_name: str = crd_path.split('\\')[-1]
            file_name = file_name[0:8]

            number_of_machines: int = all_unique_machines.get(machine_name)
            max_number: int = end_section + number_of_machines - 1

            number_of_cell = str(pipe_order + next_min_number)
            pipe_order += number_of_machines

            write_hyperlink_cells(number_of_machines, letter, number_of_cell, max_number, crd_path,
                                  file_name, worksheet, structure, structure.dark_grey_middle)

            end_section += number_of_machines

        next_min_number: int = end_section + 1

    # next_min_number: int = min_number
    # for index, pipe_name in enumerate(pipe_structure):
    #     all_unique_machines = pipe_structure.get(pipe_name, {}).get('unique_machines', 'None')
    #
    #     end_section: int = next_min_number
    #     pipe_order: int = 0
    #
    #     for start, machine_name in enumerate(all_unique_machines):
    #         ticket_id: str = console_server_data.get(pipe_name, {}). \
    #             get('pipe_data', {}).get(machine_name, {}).get('ticket', 'None')
    #         ticket_url: str = f'https://azurecsi.visualstudio.com/' \
    #                           f'CSI%20Commodity%20Qualification/_workitems/edit/{ticket_id}'
    #
    #         number_of_machines: int = all_unique_machines.get(machine_name)
    #         max_number: int = end_section + number_of_machines - 1
    #
    #         number_of_cell = str(pipe_order + next_min_number)
    #         pipe_order += number_of_machines
    #
    #         write_hyperlink_cells(number_of_machines, letter, number_of_cell, max_number, ticket_url,
    #                               'Z: Drive', worksheet, structure, structure.dark_grey_middle)
    #
    #         end_section += number_of_machines
    #
    #     next_min_number: int = end_section + 1


def write_crd_column(letter: str, min_number: int, pipe_structure: dict, ado_data: dict,
                     worksheet, structure, console_server_data: dict):
    """
    Write Pipe Name column in excel output
    :param ado_data:
    :param min_number: starting point of writing column pipe names
    :param console_server_data: for host group IDs
    :param worksheet:
    :param letter:
    :param structure:
    :param pipe_structure: organized data
    :return:
    """
    next_min_number: int = min_number
    for index, pipe_name in enumerate(pipe_structure):
        all_unique_machines = pipe_structure.get(pipe_name, {}).get('unique_machines', 'None')

        end_section: int = next_min_number
        pipe_order: int = 0

        for start, machine_name in enumerate(all_unique_machines):
            ticket_id: str = console_server_data.get(pipe_name, {}). \
                get('pipe_data', {}).get(machine_name, {}).get('ticket', 'None')
            crd_path: str = ado_data.get(ticket_id, {}).get('attachment_file_paths', {}).get('crd_drive_path', 'None')

            file_name: str = crd_path.split('\\')[-1]
            file_name = file_name[0:8]

            number_of_machines: int = all_unique_machines.get(machine_name)
            max_number: int = end_section + number_of_machines - 1

            number_of_cell = str(pipe_order + next_min_number)
            pipe_order += number_of_machines

            write_hyperlink_cells(number_of_machines, letter, number_of_cell, max_number, crd_path,
                                  file_name, worksheet, structure, structure.blue_middle)

            end_section += number_of_machines

        next_min_number: int = end_section + 1


def write_reason_column(letter: str, minimal_header_number: int, pipe_structure: dict, processed_issues,
                        ado_data: dict, worksheet: xlsxwriter, structure: xlsxwriter):
    """
    Writes Pipe Name column in excel output
    :param ado_data:
    :param processed_issues:
    :param minimal_header_number:
    :param pipe_structure:
    :param letter: which column in excel
    :param worksheet: which worksheet
    :param structure: color
    :return:
    """
    next_min_number: int = minimal_header_number
    for index, pipe_name in enumerate(pipe_structure):
        all_unique_machines = pipe_structure.get(pipe_name, {}).get('unique_machines', 'None')

        pipe_order: int = 0
        end_pipe_section: int = next_min_number
        for start, machine_name in enumerate(all_unique_machines):

            number_of_machines: int = all_unique_machines.get(machine_name)
            cell_row: int = pipe_order + next_min_number
            machine_issues: dict = pipe_structure.get(pipe_name, {}).get('machine_issues', {}).get(machine_name)
            issues_length: int = len(machine_issues)
            max_cell: int = cell_row + issues_length - 1

            for issues_index, current_issue in enumerate(machine_issues, start=1):
                current_reason: str = processed_issues.get(pipe_name, {}).get(current_issue, {}).get('reason')
                current_ticket_id: str = processed_issues.get(pipe_name, {}).get(current_issue, {}).get('ticket_id')
                ticket_state: str = ado_data.get(current_ticket_id, {}).get('state', {})

                if ticket_state == 'InProgress':

                    if issues_length == 1:
                        worksheet.write(f'{letter}{cell_row}', 'WAIVED', structure.aqua_middle)

                    elif issues_length >= 2:
                        temp_issues_length: int = issues_length - 1
                        max_temp = max_cell - temp_issues_length
                        max_temp_2 = max_temp + temp_issues_length
                        worksheet.merge_range(f'{letter}{max_temp}:{letter}{max_temp_2}', 'WAIVED',
                                              structure.aqua_middle)
                        break

                elif ticket_state == 'Test completed':
                    # worksheet.write(f'{letter}{cell_row}', 'WAIVED', structure.purple_middle)
                    if issues_length == 1:
                        worksheet.write(f'{letter}{cell_row}', 'WAIVED', structure.purple_middle)

                    elif issues_length >= 2:
                        temp_issues_length: int = issues_length - 1
                        max_temp = max_cell - temp_issues_length
                        max_temp_2 = max_temp + temp_issues_length
                        worksheet.merge_range(f'{letter}{max_temp}:{letter}{max_temp_2}', 'WAIVED',
                                              structure.purple_middle)
                        break

                else:
                    color_number = int(cell_row) % 2
                    if color_number == 0:
                        worksheet.write(f'{letter}{cell_row}', f'   {current_reason}', structure.blue_left)

                    elif color_number == 1:
                        worksheet.write(f'{letter}{cell_row}', f'   {current_reason}', structure.alt_blue_left)

                # Account for space breaks between pipe sections in the excel output
                if issues_index == issues_length:
                    cell_row += 1

                cell_row += 1

            # Increments
            pipe_order += number_of_machines
            end_pipe_section += number_of_machines

        next_min_number: int = end_pipe_section + 1

    # next_min_number: int = min_number
    # for index, pipe_name in enumerate(pipe_structure):
    #     all_unique_machines = pipe_structure.get(pipe_name, {}).get('unique_machines', 'None')
    #
    #     end_section: int = next_min_number
    #     pipe_order: int = 0
    #
    #     for start, machine_name in enumerate(all_unique_machines):
    #         ticket_id: str = console_server_data.get(pipe_name, {}). \
    #             get('pipe_data', {}).get(machine_name, {}).get('ticket', 'None')
    #         ticket_url: str = f'https://azurecsi.visualstudio.com/' \
    #                           f'CSI%20Commodity%20Qualification/_workitems/edit/{ticket_id}'
    #         ticket_state: str = ado_data.get(ticket_id, {}).get('state', {})
    #
    #         # ticket state text sometimes sloppy and inconsistent
    #         ticket_state: str = ticket_state.replace('InProgress', 'In Progress'). \
    #             replace('Test completed', 'Test Completed'). \
    #             replace('Ready To Review', 'Ready to Review'). \
    #             replace('Ready to start', 'Ready to Start')
    #
    #         number_of_machines: int = all_unique_machines.get(machine_name)
    #         max_number: int = end_section + number_of_machines - 1
    #
    #         number_of_cell = str(pipe_order + next_min_number)
    #         pipe_order += number_of_machines
    #
    #         if ticket_state == 'Done':
    #             write_hyperlink_cells(number_of_machines, letter, number_of_cell, max_number, ticket_url,
    #                                   ticket_state, worksheet, structure, structure.purple_middle)
    #         elif ticket_state == 'Blocked':
    #             write_hyperlink_cells(number_of_machines, letter, number_of_cell, max_number, ticket_url,
    #                                   ticket_state, worksheet, structure, structure.bad_cell)
    #         elif ticket_state == 'On Hold':
    #             write_hyperlink_cells(number_of_machines, letter, number_of_cell, max_number, ticket_url,
    #                                   ticket_state, worksheet, structure, structure.bad_cell)
    #         elif ticket_state == 'Test Completed':
    #             write_hyperlink_cells(number_of_machines, letter, number_of_cell, max_number, ticket_url,
    #                                   ticket_state, worksheet, structure, structure.purple_middle)
    #         elif ticket_state == 'Planning':
    #             write_hyperlink_cells(number_of_machines, letter, number_of_cell, max_number, ticket_url,
    #                                   ticket_state, worksheet, structure, structure.purple_middle)
    #
    #         elif ticket_state == 'In Progress':
    #             write_hyperlink_cells(number_of_machines, letter, number_of_cell, max_number, ticket_url,
    #                                   ticket_state, worksheet, structure, structure.aqua_middle)
    #
    #         elif ticket_state == 'Ready to Start':
    #             write_hyperlink_cells(number_of_machines, letter, number_of_cell, max_number, ticket_url,
    #                                   ticket_state, worksheet, structure, structure.aqua_middle)
    #
    #         elif ticket_state == 'Ready to Review':
    #             write_hyperlink_cells(number_of_machines, letter, number_of_cell, max_number, ticket_url,
    #                                   ticket_state, worksheet, structure, structure.aqua_middle)
    #
    #         else:
    #             write_hyperlink_cells(number_of_machines, letter, number_of_cell, max_number, ticket_url,
    #                                   ticket_state, worksheet, structure, structure.blue_middle)
    #
    #         end_section += number_of_machines
    #
    #     next_min_number: int = end_section + 1


def write_section_column(letter: str, minimal_header_number: int, pipe_structure: dict, processed_issues: dict,
                         worksheet, structure):
    """
    Writes section associated to the area of concern ex. Console Server, ADO
    :param pipe_structure:
    :param letter: which column in excel
    :param minimal_header_number: starting point of writing column pipe names
    :param worksheet: which worksheet
    :param structure: color
    :param processed_issues:
    :return:
    """
    next_min_number: int = minimal_header_number
    for index, pipe_name in enumerate(pipe_structure):
        all_unique_machines = pipe_structure.get(pipe_name, {}).get('unique_machines', 'None')

        pipe_order: int = 0
        end_pipe_section: int = next_min_number
        for start, machine_name in enumerate(all_unique_machines):

            # Information for positioning
            number_of_machines: int = all_unique_machines.get(machine_name)
            number_of_cell: int = pipe_order + next_min_number
            machine_issues: dict = pipe_structure.get(pipe_name, {}) \
                .get('machine_issues', {}).get(machine_name)
            issues_length: int = len(machine_issues)

            for issues_index, current_issue in enumerate(machine_issues, start=1):
                current_section: str = processed_issues.get(pipe_name, {}).get(current_issue, {}).get('section')

                # Spaces given before text to add readability
                # Makes sure text doesn't rub against border of cell too much
                color_number = int(number_of_cell) % 2
                if color_number == 0:
                    worksheet.write(f'{letter}{number_of_cell}', f'   {current_section}',
                                    structure.blue_left)

                elif color_number == 1:
                    worksheet.write(f'{letter}{number_of_cell}', f'   {current_section}',
                                    structure.alt_blue_left)

                # Account for space breaks between pipe sections in the excel output
                if issues_index == issues_length:
                    number_of_cell += 1
                number_of_cell += 1

            # Increments
            pipe_order += number_of_machines
            end_pipe_section += number_of_machines

        next_min_number: int = end_pipe_section + 1


def write_color_column(letter: str, minimal_header_number: int, pipe_structure, processed_issues, worksheet, structure):
    """
    Writes Pipe Name column in excel output
    :param pipe_structurre:
    :param letter: which column in excel
    :param min_number: starting point of writing column pipe names
    :param worksheet: which worksheet
    :param structure: color
    :param processed_issues: organized data
    :return:
    """
    next_min_number: int = minimal_header_number
    for index, pipe_name in enumerate(pipe_structure):
        all_unique_machines = pipe_structure.get(pipe_name, {}).get('unique_machines', 'None')

        pipe_order: int = 0
        end_pipe_section: int = next_min_number
        for start, machine_name in enumerate(all_unique_machines):

            # Information for positioning
            number_of_machines: int = all_unique_machines.get(machine_name)
            number_of_cell: int = pipe_order + next_min_number
            machine_issues: dict = pipe_structure.get(pipe_name, {}) \
                .get('machine_issues', {}).get(machine_name)
            issues_length: int = len(machine_issues)

            for issues_index, current_issue in enumerate(machine_issues, start=1):
                issue_state: str = processed_issues.get(pipe_name, {}).get(current_issue, {}).get('issue_state')

                # Spaces given before text to add readability
                # Makes sure text doesn't rub against border of cell too much
                if issue_state == 'VSE':
                    worksheet.write(f'{letter}{number_of_cell}', '', structure.aqua_middle)

                elif issue_state == 'MISMATCH':
                    worksheet.write(f'{letter}{number_of_cell}', '', structure.neutral_cell)

                elif issue_state == 'MISSING':
                    worksheet.write(f'{letter}{number_of_cell}', '', structure.purple_middle)

                # Account for space breaks between pipe sections in the excel output
                if issues_index == issues_length:
                    number_of_cell += 1
                number_of_cell += 1

            # Increments
            pipe_order += number_of_machines
            end_pipe_section += number_of_machines

        next_min_number: int = end_pipe_section + 1

    # next_min_number = 0
    # for index, pipe_name in enumerate(processed_issues):
    #     # Until ServiceNow becomes a reality, hyperlinking their website news section as a substitute for inventory
    #     pipe_issues: dict = processed_issues.get(pipe_name, {})
    #     pipe_order_count: int = 0
    #
    #     for pipe_order, issue in enumerate(pipe_issues):
    #         pipe_order_count += 1
    #         number_of_cell = str(min_number + pipe_order + next_min_number)
    #         issue_state = pipe_issues.get(issue, {}).get('issue_state', 'None')
    #
    #         if issue_state == 'VSE':
    #             worksheet.write(f'{letter}{number_of_cell}', '', structure.aqua_middle)
    #
    #         elif issue_state == 'MISMATCH':
    #             worksheet.write(f'{letter}{number_of_cell}', '', structure.neutral_cell)
    #
    #         elif issue_state == 'MISSING':
    #             worksheet.write(f'{letter}{number_of_cell}', '', structure.purple_middle)
    #
    #     # added 1 to account for spacing in between pipes
    #     next_min_number += pipe_order_count + 1


def write_item_column(letter: str, minimal_header_number: int, pipe_structure, processed_issues,
                      worksheet, structure):
    """
    Writes Pipe Name column in excel output
    :param pipe_structure:
    :param minimal_header_number:
    :param letter: which column in excel
    :param worksheet: which worksheet
    :param structure: color
    :param processed_issues: organized data
    :return:
    """
    next_min_number: int = minimal_header_number
    for index, pipe_name in enumerate(pipe_structure):
        all_unique_machines = pipe_structure.get(pipe_name, {}).get('unique_machines', 'None')

        pipe_order: int = 0
        end_pipe_section: int = next_min_number
        for start, machine_name in enumerate(all_unique_machines):

            # Information for positioning
            number_of_machines: int = all_unique_machines.get(machine_name)
            number_of_cell: int = pipe_order + next_min_number
            machine_issues: dict = pipe_structure.get(pipe_name, {}) \
                .get('machine_issues', {}).get(machine_name)
            issues_length: int = len(machine_issues)

            for issues_index, current_issue in enumerate(machine_issues, start=1):
                current_reason: str = processed_issues.get(pipe_name, {}).get(current_issue, {}).get('system_component')

                # Spaces given before text to add readability
                # Makes sure text doesn't rub against border of cell too much
                color_number = int(number_of_cell) % 2
                if color_number == 0:
                    worksheet.write(f'{letter}{number_of_cell}', f'   {current_reason}',
                                    structure.blue_left)

                elif color_number == 1:
                    worksheet.write(f'{letter}{number_of_cell}', f'   {current_reason}',
                                    structure.alt_blue_left)

                # Account for space breaks between pipe sections in the excel output
                if issues_index == issues_length:
                    number_of_cell += 1
                number_of_cell += 1

            # Increments
            pipe_order += number_of_machines
            end_pipe_section += number_of_machines

        next_min_number: int = end_pipe_section + 1

    # next_min_number = 0
    # for index, pipe_name in enumerate(processed_issues):
    #     # Until ServiceNow becomes a reality, hyperlinking their website news section as a substitute for inventory
    #     pipe_issues: dict = processed_issues.get(pipe_name, {})
    #     pipe_order_count: int = 0
    #
    #     for pipe_order, issue in enumerate(pipe_issues):
    #         pipe_order_count += 1
    #         number_of_cell = str(min_number + pipe_order + next_min_number)
    #         color_number = int(number_of_cell) % 2
    #         component = pipe_issues.get(issue, {}).get('system_component', 'None')
    #
    #         if color_number == 0:
    #             worksheet.write(f'{letter}{number_of_cell}', f'   {component}', structure.blue_left)
    #         elif color_number == 1:
    #             worksheet.write(f'{letter}{number_of_cell}', f'   {component}', structure.alt_blue_left)
    #
    #     # added 1 to account for spacing in between pipes
    #     next_min_number += pipe_order_count + 1


def calculate_time_statement(last_found_alive: float) -> str:
    """
    Write time stamp for statement
    :param last_found_alive:
    :return:
    """
    time = int(str(last_found_alive).split('.')[0])
    # Seconds
    if time < 60:
        return f'{time} seconds ago'

    # Minutes
    elif 60 <= time < 3_600:
        minutes = time / 60
        time_left = int(str(minutes).split('.')[0])
        if time_left == 1:
            return f'{time_left} minute ago'
        else:
            return f'{time_left} minutes ago'

    # Hours
    elif 3_600 <= time < 86_400:
        hours = time / 3_600
        time_left = int(str(hours).split('.')[0])
        if time_left == 1:
            return f'{time_left} hour ago'
        else:
            return f'{time_left} hours ago'

    # Days
    elif 86_400 <= time < 2_628_288:
        hours = time / 86_400
        time_left = int(str(hours).split('.')[0])
        if time_left == 1:
            return f'{time_left} day ago'
        else:
            return f'{time_left} days ago'


def write_data_columns(letter: str, data_type: str, minimal_header_number: int, pipe_structure, processed_issues,
                       worksheet, structure, console_server_data):
    """
    Writes Pipe Name column in excel output
    :param pipe_structure:
    :param data_type:
    :param console_server_data:
    :param letter: which column in excel
    :param minimal_header_number: starting point of writing column pipe names
    :param worksheet: which worksheet
    :param structure: color
    :param processed_issues: organized data
    :return:
    """
    next_min_number: int = minimal_header_number
    for index, pipe_name in enumerate(pipe_structure):
        all_unique_machines = pipe_structure.get(pipe_name, {}).get('unique_machines', 'None')

        pipe_order: int = 0
        end_pipe_section: int = next_min_number
        for start, machine_name in enumerate(all_unique_machines):

            # Information for positioning
            number_of_machines: int = all_unique_machines.get(machine_name)
            number_of_cell: int = pipe_order + next_min_number
            machine_issues: dict = pipe_structure.get(pipe_name, {}) \
                .get('machine_issues', {}).get(machine_name)
            issues_length: int = len(machine_issues)

            for issues_index, current_issue in enumerate(machine_issues, start=1):
                issue_data: str = processed_issues.get(pipe_name, {}).get(current_issue, {}). \
                    get(f'original_{data_type}_data')

                connection_status: str = console_server_data.get(pipe_name, {}). \
                    get('pipe_data', {}).get(machine_name, {}).get('connection_status', 'None')
                last_found_alive: float = console_server_data.get(pipe_name, {}). \
                    get('pipe_data', {}).get(machine_name, {}).get('last_found_alive', 'None')

                time_statement = calculate_time_statement(last_found_alive)
                issue_data = check_missing(issue_data)

                color_number = int(number_of_cell) % 2
                # if connection_status.upper() == 'DEAD' and letter == 'K':
                #     worksheet.write(f'{letter}{number_of_cell}', f'OFFLINE - {time_statement}',
                #                     structure.dark_grey_middle)
                #
                # elif connection_status.upper() == 'MOSTLY_DEAD' and letter == 'K':
                #     worksheet.write(f'{letter}{number_of_cell}', f'RECENTLY OFFLINE - {time_statement}',
                #                     structure.light_grey_middle)

                # else:
                if color_number == 0:
                    if issue_data == 'None':
                        worksheet.write(f'{letter}{number_of_cell}', '', structure.missing_cell)
                    else:
                        worksheet.write(f'{letter}{number_of_cell}', issue_data, structure.blue_middle)

                elif color_number == 1:
                    if issue_data == 'None':
                        worksheet.write(f'{letter}{number_of_cell}', '', structure.missing_cell)
                    else:
                        worksheet.write(f'{letter}{number_of_cell}', issue_data, structure.alt_blue_middle)

                # Account for space breaks between pipe sections in the excel output
                if issues_index == issues_length:
                    number_of_cell += 1
                number_of_cell += 1

            # Increments
            pipe_order += number_of_machines
            end_pipe_section += number_of_machines

        next_min_number: int = end_pipe_section + 1

    # next_min_number = 0
    # for index, pipe_name in enumerate(processed_issues):
    #     # Until ServiceNow becomes a reality, hyperlinking their website news section as a substitute for inventory
    #     pipe_issues: dict = processed_issues.get(pipe_name, {})
    #     # import json
    #     # print(json.dumps(pipe_issues, sort_keys=True, indent=4))
    #     # input()
    #     pipe_order_count: int = 0
    #
    #     for pipe_order, issue in enumerate(pipe_issues):
    #         pipe_order_count += 1
    #         number_of_cell = str(min_number + pipe_order + next_min_number)
    #         color_number = int(number_of_cell) % 2
    #
    #         original_system_data: str = pipe_issues.get(issue, {}).get(f'original_{data_type}_data', 'None')
    #         machine_name: str = pipe_issues.get(issue, {}).get(f'machine_name', 'None')
    #         connection_status: str = console_server_data.get(pipe_name, {}). \
    #             get('pipe_data', {}).get(machine_name, {}).get('connection_status', 'None')
    #         last_found_alive: float = console_server_data.get(pipe_name, {}). \
    #             get('pipe_data', {}).get(machine_name, {}).get('last_found_alive', 'None')
    #
    #         time_statement = calculate_time_statement(last_found_alive)
    #         system_data = check_missing(original_system_data)
    #
    #         if connection_status.upper() == 'DEAD' and letter == 'K':
    #             worksheet.write(f'{letter}{number_of_cell}', f'OFFLINE - {time_statement}',
    #                             structure.dark_grey_middle)
    #
    #         elif connection_status.upper() == 'MOSTLY_DEAD' and letter == 'K':
    #             worksheet.write(f'{letter}{number_of_cell}', f'RECENTLY OFFLINE - {time_statement}',
    #                             structure.light_grey_middle)
    #
    #         else:
    #             if color_number == 0:
    #                 if system_data == 'None':
    #                     worksheet.write(f'{letter}{number_of_cell}', '', structure.missing_cell)
    #                 else:
    #                     worksheet.write(f'{letter}{number_of_cell}', system_data, structure.blue_middle)
    #
    #             elif color_number == 1:
    #                 if system_data == 'None':
    #                     worksheet.write(f'{letter}{number_of_cell}', '', structure.missing_cell)
    #                 else:
    #                     worksheet.write(f'{letter}{number_of_cell}', system_data, structure.alt_blue_middle)
    #
    #     # added 1 to account for spacing in between pipes
    #     next_min_number += pipe_order_count + 1


def process_machine_names(all_machines_in_pipes: list):
    """

    :param all_machines_in_pipes:
    :return:
    """
    all_unique_machines: dict = {}
    unique_machines = sorted(list(set(all_machines_in_pipes)))
    for machine_name in unique_machines:
        all_unique_machines[machine_name] = 0

    for unique_machine in unique_machines:
        for machine_name in all_machines_in_pipes:
            if unique_machine in machine_name:
                all_unique_machines[machine_name] += 1

    return all_unique_machines


def process_sku_issues(all_machines_in_pipes: list):
    """

    :param all_machines_in_pipes:
    :return:
    """

    all_unique_machines: dict = {}
    unique_machines = sorted(list(set(all_machines_in_pipes)))
    for machine_name in unique_machines:
        all_unique_machines[machine_name] = 0

    all_machine_names: list = []
    for unique_machine in unique_machines:
        for machine_name in all_machines_in_pipes:
            if unique_machine in machine_name:
                all_unique_machines[machine_name] += 1
                machine_name = machine_name[8:11]
                all_machine_names.append(machine_name)
    unique_skus = sorted(list(set(all_machine_names)))

    skus_count: dict = {}
    for sku in unique_skus:
        skus_count[sku] = 0

    for machine_tag in all_unique_machines:
        current_count = all_unique_machines.get(machine_tag)
        current_sku = machine_tag[8:11]
        skus_count[current_sku] += current_count

    return skus_count


def process_unique_types(all_machines_in_pipes: list):
    """

    :param all_machines_in_pipes:
    :return:
    """
    machine_types = sorted(list(set(all_machines_in_pipes)))
    unique_machine_types: dict = {}
    all_machine_types: list = []

    for machine_name in machine_types:
        machine_type = machine_name[8:11]
        unique_machine_types[machine_type] = 0
        all_machine_types.append(machine_type)

    for machine_type in all_machine_types:
        for machine in unique_machine_types:
            if machine_type in machine:
                unique_machine_types[machine_type] += 1

    return unique_machine_types


def get_machine_issues(all_machines_in_pipes: list, issues_dict: dict):
    """

    :param issues_dict:
    :param all_machines_in_pipes:
    :return:
    """
    all_unique_machines: dict = {}
    unique_machines = sorted(list(set(all_machines_in_pipes)))
    for machine_name in unique_machines:
        all_unique_machines[machine_name] = []

    for unique_machine in unique_machines:
        for machine_name in issues_dict:
            if unique_machine in machine_name:
                all_unique_machines.get(unique_machine, []).append(machine_name)

    # import json
    # foo = json.dumps(all_unique_machines, sort_keys=True, indent=4)
    # print(foo)
    # input()

    return all_unique_machines


def process_data(processed_issues: dict, issues_dict: dict) -> dict:
    """

    :param issues_dict:
    :param processed_issues:
    :return:
    """
    new_all_issues: dict = {}

    all_machines_in_pipe: list = []
    for unique_pipe in processed_issues:
        new_all_issues[unique_pipe] = {}
        pipe_issues = processed_issues.get(unique_pipe)
        new_all_issues[unique_pipe]['total_issues'] = len(pipe_issues)

        all_machines_in_pipe.clear()
        for item in processed_issues.get(unique_pipe):
            machine_name = processed_issues.get(unique_pipe, {}).get(item, {}).get('machine_name')
            all_machines_in_pipe.append(machine_name)

        new_all_issues[unique_pipe]['total_machines'] = len(list(set(all_machines_in_pipe)))
        new_all_issues[unique_pipe]['unique_machines'] = process_machine_names(all_machines_in_pipe)
        new_all_issues[unique_pipe]['machine_issues'] = get_machine_issues(all_machines_in_pipe, issues_dict)
        new_all_issues[unique_pipe]['machine_types'] = process_unique_types(all_machines_in_pipe)
        new_all_issues[unique_pipe]['sku_type_issues'] = process_sku_issues(all_machines_in_pipe)

    return new_all_issues


def sort_processed_issues(processed_issues: dict) -> dict:
    """
    Organize the processed issues alphabetically to make excel output consistent
    :param processed_issues:
    :return:
    """
    import json
    print(json.dumps(processed_issues, sort_keys=True, indent=4))
    input()

    return processed_issues


def process_issues_into_list(all_issues: list) -> dict:
    """

    :param all_issues:
    :return:
    """
    all_issues_dict: dict = {}
    for issue in all_issues:
        machine_name: str = issue.get('machine_name')
        issue_state: str = issue.get('issue_state')
        system_component = str(issue.get('system_component')).lower().replace(' ', '_')
        all_issues_dict[f'{machine_name}_{issue_state}_{system_component}'] = issue
    return all_issues_dict


def add_issue_data(ado_data: dict, console_server_data: dict, all_issues: list, worksheet, structure):
    """

    :param console_server_data:
    :param ado_data:
    :param all_issues:
    :param worksheet:
    :param structure:
    :return:
    """
    minimal_header_number: int = 14

    processed_issues: dict = process_all_issues(all_issues)
    issues_dict: dict = process_issues_into_list(all_issues)
    pipe_structure: dict = process_data(processed_issues, issues_dict)

    write_pipe_name_column('C', minimal_header_number, processed_issues, worksheet, structure, console_server_data)

    write_sku_column('D', pipe_structure, minimal_header_number, worksheet, structure)

    write_machine_name_column('E', minimal_header_number, pipe_structure, worksheet, structure, console_server_data)

    write_section_column('F', minimal_header_number, pipe_structure, processed_issues, worksheet, structure)

    write_reason_column('G', minimal_header_number, pipe_structure, processed_issues, ado_data, worksheet, structure)

    write_item_column('H', minimal_header_number, pipe_structure, processed_issues, worksheet, structure)

    write_color_column('I', minimal_header_number, pipe_structure, processed_issues, worksheet, structure)

    write_data_columns('K', 'system', minimal_header_number, pipe_structure, processed_issues, worksheet,
                       structure, console_server_data)

    write_data_columns('L', 'ticket', minimal_header_number, pipe_structure, processed_issues, worksheet,
                       structure, console_server_data)

    write_trr_column('N', minimal_header_number, pipe_structure, worksheet, structure,
                     console_server_data)

    write_type_column('O', minimal_header_number, pipe_structure, ado_data, worksheet, structure,
                      console_server_data)

    write_tests_column('P', minimal_header_number, pipe_structure, ado_data, worksheet, structure,
                       console_server_data)

    write_state_column('Q', minimal_header_number, pipe_structure, ado_data, worksheet, structure,
                       console_server_data)

    write_crd_column('R', minimal_header_number, pipe_structure, ado_data, worksheet, structure,
                     console_server_data)

    write_skudoc_column('S', minimal_header_number, pipe_structure, ado_data, worksheet, structure,
                        console_server_data)


def write_data(column_name, letter_cell: str, number_cell: str, data: str, worksheet, structure,
               info_package, color, issues_length: dict):
    """
    Create output
    :param column_name: column title
    :param letter_cell:
    :param number_cell:
    :param issues_length:
    :param color:
    :param info_package:
    :param data:
    :param worksheet:
    :param structure:
    :return:
    """
    color_middle = process_color(color, 'MIDDLE', structure)
    color_left = process_color(color, 'LEFT', structure)

    # Assure uppercase for matching purposes
    column_name = column_name.upper()
    letter_cell = letter_cell.upper()

    clean_data = str(data).replace('[', '').replace(']', '').replace("'", '')

    url_container: list = []
    if letter_cell == 'C':
        url_container.append(info_package.get('host_group_url', 'None'))
    elif letter_cell == 'D':
        url_container.append(info_package.get('host_url', 'None'))
    elif letter_cell == 'E':
        url_container.append(info_package.get('ticket_url', 'None'))

    if letter_cell == 'C' and \
            column_name == 'PIPE NAME':

        # Check if information is there
        pipe_length = str(issues_length.get(data, 'None'))
        clean_pipe_name: str = process_pipe_name(clean_data)

        # First One
        if len(pipe_information) == 0:
            next_line = str(int(pipe_length) + int(number_cell) + 1)
            group_length = str(int(pipe_length) + int(number_cell) - 1)
            pipe_information['pipe_name'] = data
            pipe_information['next_line'] = next_line
            pipe_information['pipe_number'] = 0

            worksheet.set_row(int(group_length), 12, structure.white)

            if pipe_length == 1:
                worksheet.write_url(f'{letter_cell}{number_cell}', url_container[0],
                                    color_middle, string=clean_pipe_name)

            elif int(pipe_length) > 1:
                worksheet.merge_range(f'{letter_cell}{number_cell}:{letter_cell}{group_length}',
                                      data, structure.blue_middle_huge)
                worksheet.write_url(f'{letter_cell}{number_cell}', url_container[0],
                                    structure.blue_middle_huge, string=clean_pipe_name)

        elif pipe_information.get('next_line') == number_cell:

            next_line = str(int(pipe_length) + int(number_cell) + 1)
            group_length = str(int(pipe_length) + int(number_cell) - 1)
            pipe_information['pipe_name'] = data
            pipe_information['next_line'] = next_line
            pipe_information['pipe_number'] += 1

            worksheet.set_row(int(group_length), 12, structure.white)

            if int(pipe_length) == 1:
                worksheet.write_url(f'{letter_cell}{number_cell}', url_container[0],
                                    structure.blue_middle, string=clean_pipe_name)

            elif int(pipe_length) > 1:
                worksheet.merge_range(f'{letter_cell}{number_cell}:{letter_cell}{group_length}',
                                      data, structure.blue_middle_huge)
                worksheet.write_url(f'{letter_cell}{number_cell}', url_container[0],
                                    structure.blue_middle_huge, string=clean_pipe_name)

    elif letter_cell == 'D' and \
            column_name == 'MACHINE NAME':

        pipe_number = pipe_information.get('pipe_number')
        number_true = str(int(pipe_number) + int(number_cell))
        pipe_name: str = pipe_information.get('pipe_name')

        if pipe_name:

            initial = issues_length.get('initial')

            if info_package.get('connection_status') == 'DEAD':
                worksheet.write_url(f'{letter_cell}{initial}', url_container[0],
                                    structure.dark_grey_middle, string=clean_data)

            elif info_package.get('connection_status') == 'ALIVE':
                worksheet.write_url(f'{letter_cell}{initial}', url_container[0],
                                    color_middle, string=clean_data)

        else:
            pass
            # Don't know what mostly dead means...
            # Will pop as dead data
            # worksheet.write_url(f'{letter_cell}{number_true}', url_container[0],
            #                     structure.dark_grey_middle, string=clean_data)

    elif letter_cell == 'E' and \
            column_name == 'INVENTORY':

        worksheet.write_url(f'{letter_cell}{number_cell}', 'https://www.servicenow.com',
                            structure.dark_grey_middle, string='ServiceNow')

    elif letter_cell == 'F' and \
            column_name == 'ITEM STATE':

        if data.upper() == 'MISMATCH':
            worksheet.write(f'{letter_cell}{number_cell}', '', structure.neutral_cell)
        elif data.upper() == 'MISSING':
            worksheet.write(f'{letter_cell}{number_cell}', '', structure.bad_cell)

    elif letter_cell == 'G' and \
            column_name == 'ITEM':

        if check_missing(data) == 'None':
            worksheet.write(f'{letter_cell}{number_cell}', '', structure.missing_cell)
        else:
            worksheet.write(f'{letter_cell}{number_cell}', data, color_middle)

    elif letter_cell == 'I' and \
            column_name == 'CONSOLE SERVER':

        if check_missing(data) == 'None':
            worksheet.write(f'{letter_cell}{number_cell}', '', structure.missing_cell)
        else:
            worksheet.write(f'{letter_cell}{number_cell}', data, color_middle)

    elif letter_cell == 'J' and \
            column_name == 'AZURE DEVOPS':

        if check_missing(data) == 'None':
            worksheet.write(f'{letter_cell}{number_cell}', '', structure.missing_cell)
        else:
            worksheet.write(f'{letter_cell}{number_cell}', data, color_middle)

    elif letter_cell == 'M' and \
            column_name == 'REQUEST TYPE':

        if check_missing(data) == 'None':
            worksheet.write(f'{letter_cell}{number_cell}', '', structure.missing_cell)
        else:
            worksheet.write(f'{letter_cell}{number_cell}', data, color_middle)

    elif letter_cell == 'N' and \
            column_name == 'TESTS':

        if check_missing(data) == 'None':
            worksheet.write(f'{letter_cell}{number_cell}', 'TBA', structure.dark_grey_middle)
        else:
            worksheet.write(f'{letter_cell}{number_cell}', data, color_middle)

    elif letter_cell == 'O' and \
            column_name == 'TRR STATE':

        if check_missing(data) == 'None':
            worksheet.write(f'{letter_cell}{number_cell}', '', structure.missing_cell)
        elif data == 'Test completed':
            worksheet.write(f'{letter_cell}{number_cell}', data, structure.dark_grey_middle)
        else:
            worksheet.write(f'{letter_cell}{number_cell}', data, color_middle)

    elif letter_cell == 'P' and \
            column_name == 'CRD':

        worksheet.write(f'{letter_cell}{number_cell}', 'Z:Drive', structure.dark_grey_middle)

    elif letter_cell == 'Q' and \
            column_name == 'SKUDOC':
        worksheet.write(f'{letter_cell}{number_cell}', 'Z:Drive', structure.dark_grey_middle)

    elif len(url_container) == 1:
        if check_missing(data) == 'None':
            worksheet.write_url(f'{letter_cell}{number_cell}', '', structure.missing_cell, string=clean_data)

        elif data == 'MISMATCH':
            worksheet.write_url(f'{letter_cell}{number_cell}', url_container[0],
                                structure.neutral_cell, string=clean_data)
        elif data == 'MISSING':
            worksheet.write_url(f'{letter_cell}{number_cell}', url_container[0],
                                structure.bad_cell, string=clean_data)
        elif 'VSE' in data or 'Pipe-' in data:
            worksheet.write_url(f'{letter_cell}{number_cell}', url_container[0],
                                color_left, string=clean_data)
        else:
            worksheet.write_url(f'{letter_cell}{number_cell}', url_container[0],
                                color_middle, string=clean_data)


def create_breakdown_graph(console_server_data: dict, workbook: xlsxwriter, worksheet: xlsxwriter, sheet_name: str,
                           mismatch_tally: str, missing_tally: str):
    """
    Create Graph for Issues
    :param console_server_data:
    :param workbook:
    :param worksheet:
    :param sheet_name:
    :param mismatch_tally:
    :param missing_tally:
    :return:
    """
    # import json
    # print(json.dumps(console_server_data, sort_keys=True, indent=4))
    # input()
    vse_log: int = console_server_data.get('host_groups_data', {}).get('vse_log', 0)

    # Add the worksheet data that the charts will refer to.
    headings: list = ['Number', 'Tallies']
    data = [
        ['Azure DevOps (TRRs)', 'Comparison', 'Veritas Engineering & Services'],
        [int(missing_tally), int(mismatch_tally), int(vse_log)],
    ]

    # Write to excel output to hold data for graph, bolded Title
    bold = workbook.add_format({'bold': 1})
    worksheet.write_row('A1', headings, bold)
    worksheet.write_column('A2', data[0])
    worksheet.write_column('B2', data[1])

    # Type of Graph
    chart_structure = workbook.add_chart({'type': 'bar'})

    # Structure Graph
    chart_structure.add_series({
        'name': "='" + sheet_name + "'!$B$1",
        'categories': "='" + sheet_name + "'!$A$2:$A$4",
        'values': "='" + sheet_name + "'!$B$2:$B$4",
        'points': [
            {'fill': {'color': '#7030A0'}},
            # {'fill': {'color': '#FF0000'}},
            {'fill': {'color': '#DCAA1B'}},
            {'fill': {'color': '#31869B'}},  # Aqua Color
        ],
    })

    # Configure a second series. Note use of alternative syntax to define ranges.
    chart_structure.add_series({
        'name': [f"{sheet_name}", 0, 2],
        'categories': [f"{sheet_name}", 1, 0, 3, 0],
        'values': [f"{sheet_name}", 1, 2, 3, 2],
    })

    # Add a chart title and some axis labels.
    chart_structure.set_title({'name': 'Breakdown of Issues'})

    # Chart Style of Graph
    chart_structure.set_style(11)
    chart_structure.set_legend({'none': True})

    # Size of Chart
    worksheet.insert_chart('J1', chart_structure, {'x_scale': 2.350, 'y_scale': 0.84})


def set_issue_structure(worksheet, structure, sheet_title, site_location, total_issues, total_checks,
                        pipe_numbers, pipe_cleaner_version: str):
    """
    Create dashboard structure
    :param pipe_cleaner_version:
    :param pipe_numbers:
    :param total_checks:
    :param total_issues:
    :param worksheet:
    :param structure:
    :param sheet_title:
    :param site_location:
    :return:
    """
    time = strftime('%I:%M %p')
    date = strftime('%m/%d/%Y')
    default_name = str(getpass.getuser()).replace('.', ' ').title().replace('-Ext', '')

    # Set Top Plane of Excel Sheet
    top_plane_height = 13

    # Structure of the Excel Sheet
    set_issue_layout(worksheet, structure)
    set_issue_columns(top_plane_height, worksheet, structure)

    # Freeze Planes
    worksheet.freeze_panes(top_plane_height, 9)

    # while top_plane_height < 500:
    #     worksheet.set_row(top_plane_height, 16.5, structure.white)
    #     top_plane_height += 1

    correct_total = int(total_checks) - int(total_issues)
    total_pipes = str(pipe_numbers.get('pipes'))
    total_systems = str(pipe_numbers.get('systems'))

    pipe_cleaner_version: str = pipe_cleaner_version.split(' ')[0]

    # Top Left Plane
    worksheet.insert_image('A1', 'pipe_cleaner/img/vse_logo.png')
    worksheet.write('B5', f' Pipe Cleaner - {sheet_title}', structure.big_blue_font)
    worksheet.write('B6', f'       {site_location}', structure.bold_italic_blue_font)
    worksheet.write('B7', f'            Pipes - {total_pipes}', structure.bold_italic_blue_font)
    worksheet.write('B8', f'            Checks - {total_checks}', structure.bold_italic_blue_font)
    # worksheet.write('B9', f'            Online - ', structure.bold_italic_blue_font)
    worksheet.write('D7', f'Blades - {total_systems}', structure.bold_italic_blue_font)
    worksheet.write('D8', f'Issues - {total_issues}', structure.bold_italic_blue_font)
    # worksheet.write('D9', f'Offline - ', structure.bold_italic_blue_font)
    worksheet.write('B10', f'            Percentage Correct - N/A %', structure.bold_italic_green_font)
    worksheet.write('B11', f'       {date} - {time} - {default_name} - v{pipe_cleaner_version}',
                    structure.italic_blue_font)

    worksheet.merge_range('F6:H6', f'Checks being Done', structure.red_middle_18)
    worksheet.write('F7', f'Console Server [BIOS, BMC, CPLD, OS]', structure.bold_italic_blue_font)
    worksheet.write('F8', f'Chipset Driver, NIC [FM, PXE, UEFI, Driver]', structure.bold_italic_blue_font)
    worksheet.write('F9', f'Server PSU FW, Rack Manager FW, Boot Drive', structure.bold_italic_blue_font)
    worksheet.write('F10', f'Request Type, Target Config, Part Number', structure.bold_italic_blue_font)
    worksheet.write('F11', f'Supplier, Toolkit, CRD [BIOS, BMC], TRR Title', structure.bold_italic_blue_font)


def set_issue_columns(top_plane_height, worksheet, structure):
    """
    Set up Column Names in the Excel table for adding data later
    :param top_plane_height:
    :param worksheet:
    :param structure:
    :return:
    """
    name_to_number: dict = {}

    column_names: list = [
        'Pipe',
        'SKU',
        'Machine',
        'Section',
        'Reason',
        'Issue Item',
        '',
        '',
        'Console Server',
        'Client',
        '',
        'TRR',
        'Type',
        'Tests',
        'State',
        'CRD',
        'Skudoc',
        '']

    # Number part of the excel position
    number = str(top_plane_height)

    initial = 0
    while initial < len(column_names):
        little = chr(ord('c') + initial)
        letter = str(little).upper()

        # Pipe Column
        if letter == 'C':
            worksheet.write_url(f'{letter}{top_plane_height}', 'http://172.30.1.100/console/host_groups.php',
                                structure.teal_middle, f'{column_names[initial]}')

        # SKU Column
        elif letter == 'D':
            worksheet.write_url(f'{letter}{top_plane_height}', '',
                                structure.teal_middle, f' {column_names[initial]}')

        # Machine Column
        elif letter == 'E':
            worksheet.write_url(f'{letter}{top_plane_height}', 'http://172.30.1.100/console/all_hosts.php',
                                structure.teal_middle, f'{column_names[initial]}')

        # Section Column
        elif letter == 'F':
            worksheet.write_url(f'{letter}{top_plane_height}', '',
                                structure.teal_left, f' {column_names[initial]}')

        # Reason Column
        elif letter == 'G':
            worksheet.write_url(f'{letter}{top_plane_height}', 'https://www.servicenow.com',
                                structure.teal_left, f' {column_names[initial]}')
        # Item Column
        elif letter == 'H':
            bigger = chr(ord('c') + initial + 1)
            bigger = str(bigger).upper()
            worksheet.merge_range(f'{letter}{top_plane_height}:{bigger}{top_plane_height}', f' {column_names[initial]}',
                                  structure.teal_left)
            worksheet.write_url(f'{letter}{top_plane_height}:{bigger}{top_plane_height}',
                                'http://172.30.1.100/console/query.php',
                                structure.teal_left, f' {column_names[initial]}')

        # Item Column
        elif letter == 'N':
            worksheet.write_url(f'{letter}{top_plane_height}',
                                'https://azurecsi.visualstudio.com/CSI%20Commodity%20Qualification/',
                                structure.teal_middle, f'{column_names[initial]}')

        elif letter == 'J' or letter == 'M' or letter == 'T':
            worksheet.write(f'{letter}{top_plane_height}', '', structure.white)

        else:
            worksheet.write(f'{letter}{top_plane_height}', f'{column_names[initial]}', structure.teal_middle)

        # Create key for dictionary
        name = str(column_names[initial]).lower().replace(' ', '_')
        number = initial + 1

        name_to_number[name] = str(number)

        initial += 1

    return name_to_number


def set_issue_layout(worksheet, structure):
    """
    Beginning of the Excel Structure
    :return:
    """
    worksheet.set_row(0, 12, structure.white)
    worksheet.set_row(1, 20, structure.white)
    worksheet.set_row(2, 15, structure.white)
    worksheet.set_row(3, 15, structure.white)
    worksheet.set_row(4, 15, structure.white)
    worksheet.set_row(5, 15, structure.white)
    worksheet.set_row(6, 15, structure.white)
    worksheet.set_row(7, 15, structure.white)
    worksheet.set_row(8, 15, structure.white)
    worksheet.set_row(9, 15, structure.white)
    worksheet.set_row(10, 15, structure.white)
    worksheet.set_row(11, 15, structure.white)

    worksheet.set_column('A:A', 0.67, structure.white)
    worksheet.set_column('B:B', 0.67, structure.white)
    worksheet.set_column('C:C', 24, structure.white)
    worksheet.set_column('D:D', 11, structure.white)
    worksheet.set_column('E:E', 19, structure.white)
    worksheet.set_column('F:F', 16, structure.white)
    worksheet.set_column('G:G', 14, structure.white)
    worksheet.set_column('H:H', 16, structure.white)
    worksheet.set_column('I:I', 1.57, structure.white)
    worksheet.set_column('J:J', 1.57, structure.white)
    worksheet.set_column('K:K', 46, structure.white)
    worksheet.set_column('L:L', 46, structure.white)
    worksheet.set_column('M:M', 1.57, structure.white)
    worksheet.set_column('N:N', 9, structure.white)
    worksheet.set_column('O:O', 9, structure.white)
    worksheet.set_column('P:P', 10, structure.white)
    worksheet.set_column('Q:Q', 17, structure.white)
    worksheet.set_column('R:R', 12, structure.white)
    worksheet.set_column('S:S', 13, structure.white)
    worksheet.set_column('T:T', 1.57, structure.white)
    worksheet.set_column('U:U', 37, structure.white)
    worksheet.set_column('V:V', 37, structure.white)


def main_method(ado_data: dict, console_server_data: dict, workbook, structure, site_location: str, all_issues,
                all_checks, mismatch_tally: str, missing_tally: str, pipe_numbers: dict, pipe_cleaner_version: str):
    """

    :param pipe_cleaner_version:
    :param pipe_numbers:
    :param console_server_data:
    :param ado_data:
    :param workbook:
    :param structure:
    :param site_location:
    :param all_issues:
    :param all_checks:
    :param mismatch_tally:
    :param missing_tally:
    :return:
    """
    sheet_name: str = 'All Issues'

    worksheet_issues = workbook.add_worksheet(sheet_name)

    set_issue_structure(worksheet_issues, structure, sheet_name, site_location, len(all_issues), all_checks,
                        pipe_numbers, pipe_cleaner_version)

    add_issue_data(ado_data, console_server_data, all_issues, worksheet_issues, structure)

    create_breakdown_graph(console_server_data, workbook, worksheet_issues, sheet_name, mismatch_tally, missing_tally)

    worksheet_issues.ignore_errors({'number_stored_as_text': 'A1:XFD1048576'})
