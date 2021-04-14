from time import strftime

import xlsxwriter

from pipe_cleaner.src.dashboard_main_virtual_machine import get_virtual_machines


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


def check_missing(data: str) -> str:
    """

    :param data:
    :return:
    """
    if data == 'None' or data == '' or data is None:
        return 'None'
    else:
        return data


def write_pipe_name_column(initial_point: int, user_sorted_pipes: dict, worksheet, structure,
                           console_server_data: dict, user_virtual_machines: dict):
    """
    Writes Pipe Name column in excel output
    :param user_virtual_machines:
    :param initial_point: starting point of writing column pipe names
    :param console_server_data: for host group IDs
    :param worksheet: which worksheet
    :param structure: color
    :param user_sorted_pipes: organized data
    :return:
    """
    column: str = 'C'

    current_pipe_point = initial_point
    for index, pipe_name in enumerate(user_sorted_pipes):

        pipe_size: int = get_pipe_size(pipe_name, user_sorted_pipes)
        pipe_max_size: int = current_pipe_point + pipe_size - 1

        clean_pipe_name: str = process_pipe_name(pipe_name)
        pipe_hyperlink: str = get_pipe_hyperlink(console_server_data, pipe_name)

        if pipe_size >= 2:
            worksheet.merge_range(f'{column}{current_pipe_point}:{column}{pipe_max_size}', clean_pipe_name,
                                  structure.blue_middle_huge)

            worksheet.write_url(f'{column}{current_pipe_point}', pipe_hyperlink, structure.blue_middle_huge,
                                string=clean_pipe_name)

        else:
            worksheet.write_url(f'{column}{current_pipe_point}', pipe_hyperlink, structure.blue_middle,
                                string=clean_pipe_name)

        worksheet.set_row(pipe_max_size, 13.5, structure.white)
        current_pipe_point += pipe_size + 1

    virtual_machine_size: int = len(user_virtual_machines)
    current_pipe_point += 1

    if virtual_machine_size >= 2:
        proper_size = current_pipe_point + virtual_machine_size - 1
        worksheet.merge_range(f'{column}{current_pipe_point}:{column}{proper_size}', 'VMs',
                              structure.blue_middle_huge)

    else:
        worksheet.write(f'{column}{current_pipe_point}', 'VM', structure.blue_middle_huge)


def get_pipe_size(pipe_name, user_sorted_pipes):
    current_pipe = user_sorted_pipes.get(pipe_name)
    return get_current_pipe_size(current_pipe)


def get_current_pipe_size(current_pipe):
    tally: int = 0
    for item in current_pipe:
        current_machine: dict = current_pipe[item]
        if len(current_machine) == 0:
            tally += 1
        else:
            tally += len(current_machine)
    return tally


def get_pipe_hyperlink(console_server_data, pipe_name):
    host_group_id: str = console_server_data.get(pipe_name, {}).get('host_id', 'None')
    host_group_url: str = f'http://172.30.1.100/console/host_group_host_list.php?host_group_id={host_group_id}'
    return host_group_url


def count_issues_in_user_pipe(pipe_name, user_sorted_pipes):
    systems_issues: dict = user_sorted_pipes[pipe_name]

    count: int = 0
    for system_issue in systems_issues:
        bar = len(user_sorted_pipes[pipe_name][system_issue])
        count += bar
    return count


def count_systems_in_user_pipe(pipe_name, user_sorted_pipes):
    return len(user_sorted_pipes[pipe_name])


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
        try:
            worksheet.write_url(f'{letter}{number_of_cell}', hyperlink, color,
                                string=text)
        except AttributeError:
            worksheet.write_url(f'{letter}{number_of_cell}', hyperlink, structure.missing_cell,
                                string='')

    elif number_of_machines >= 2:
        worksheet.merge_range(f'{letter}{number_of_cell}:{letter}{max_number}',
                              text, color)

        worksheet.write_url(f'{letter}{number_of_cell}', hyperlink, color,
                            string=text)


def add_machine_name_column(initial_point, user_sorted_pipes, worksheet, structure,
                            console_server_data, user_virtual_machines, virtual_machines_data, tally_storage):
    """
    Write host group column in excel output

    :param tally_storage:
    :param virtual_machines_data:
    :param user_virtual_machines:
    :param user_sorted_pipes:
    :param initial_point: starting point of writing column pipe names
    :param worksheet:
    :param structure:
    :param console_server_data: for host group IDs
    :return:
    """
    column: str = 'D'

    color_change: int = 0
    current_pipe_point: int = initial_point
    for index, pipe_name in enumerate(user_sorted_pipes):
        user_machines = sorted(list(user_sorted_pipes.get(pipe_name, {}).keys()))
        tally_storage['pipes'] += 1

        machine_order: int = 0
        pipe_max_size: int = current_pipe_point
        for start, machine_name in enumerate(user_machines):

            machine_issues: dict = user_sorted_pipes.get(pipe_name, {}).get(machine_name)
            machine_issues_total: int = get_machine_issues_size(machine_issues)

            max_number: int = pipe_max_size + machine_issues_total - 1
            number_of_cell = str(machine_order + current_pipe_point)

            machine_hyperlink: str = get_machine_hyperlink(console_server_data, machine_name, pipe_name)
            connection_status: str = get_machine_connection_status(console_server_data, machine_name, pipe_name)

            contrast_color = get_contrast_color(color_change, structure)

            if connection_status.upper() == 'DEAD':
                write_hyperlink_cells(machine_issues_total, column, number_of_cell, max_number, machine_hyperlink,
                                      machine_name, worksheet, structure, structure.dark_grey_middle)
                tally_storage['dead'] += 1
                tally_storage['total_machines'] += 1

            elif connection_status.upper() == 'MOSTLY_DEAD':
                write_hyperlink_cells(machine_issues_total, column, number_of_cell, max_number, machine_hyperlink,
                                      machine_name, worksheet, structure, structure.light_grey_middle)
                tally_storage['dead'] += 1
                tally_storage['total_machines'] += 1

            elif connection_status.upper() == 'ALIVE':
                write_hyperlink_cells(machine_issues_total, column, number_of_cell, max_number, machine_hyperlink,
                                      machine_name, worksheet, structure, contrast_color)
                tally_storage['alive'] += 1
                tally_storage['total_machines'] += 1

            else:
                write_hyperlink_cells(machine_issues_total, column, number_of_cell, max_number, machine_hyperlink,
                                      machine_name, worksheet, structure, structure.light_red_middle_11)

            if machine_issues_total == 1:
                actual_max_number = max_number - 1
                worksheet.set_row(actual_max_number, 22.50)

            machine_order += machine_issues_total
            pipe_max_size += machine_issues_total
            color_change += 1

        current_pipe_point: int = pipe_max_size + 1

    current_pipe_point += 1

    for index, item in enumerate(user_virtual_machines):
        contrast_color = get_contrast_color(color_change, structure)
        current_row = current_pipe_point + index
        virtual_machine_hyperlink: str = get_rdp_connection_hyperlink(item, virtual_machines_data)

        worksheet.write_url(f'{column}{current_row}', virtual_machine_hyperlink, contrast_color, string=item)
        color_change += 1
        tally_storage['virtual_machines'] += 1

    return tally_storage


def get_rdp_connection_hyperlink(item, virtual_machines):
    rdp_connection_string: str = virtual_machines.get(item, {}).get('rdp_connection_string')
    virtual_machine_hyperlink = f'http://172.30.1.100/guacamole/#/client/{rdp_connection_string}'
    return virtual_machine_hyperlink


def get_contrast_color(initial_point, structure):
    contrast_color = initial_point % 2
    if contrast_color == 0:
        return structure.blue_middle
    else:
        return structure.alt_blue_middle


def write_status_field_column(initial_point, user_sorted_pipes, worksheet, structure,
                              console_server_data, user_virtual_machines, virtual_machines_data):
    """
    Write host group column in excel output

    :param virtual_machines_data:
    :param user_virtual_machines:
    :param user_sorted_pipes:
    :param initial_point: starting point of writing column pipe names
    :param worksheet:
    :param structure:
    :param console_server_data: for host group IDs
    :return:
    """
    column: str = 'E'

    color_change: int = 0
    current_pipe_point: int = initial_point
    for index, pipe_name in enumerate(user_sorted_pipes):
        user_machines = sorted(list(user_sorted_pipes.get(pipe_name, {}).keys()))

        machine_order: int = 0
        pipe_max_size: int = current_pipe_point
        for start, machine_name in enumerate(user_machines):
            machine_issues: dict = user_sorted_pipes.get(pipe_name, {}).get(machine_name)
            machine_issues_total: int = get_machine_issues_size(machine_issues)

            machine_comment: str = get_machine_comment(console_server_data, machine_name, pipe_name)

            max_number: int = pipe_max_size + machine_issues_total - 1
            number_of_cell = str(machine_order + current_pipe_point)

            contrast_color = get_contrast_color(color_change, structure)

            write_hyperlink_cells(machine_issues_total, column, number_of_cell, max_number, '',
                                  machine_comment, worksheet, structure, contrast_color)

            machine_order += machine_issues_total
            pipe_max_size += machine_issues_total
            color_change += 1

        current_pipe_point: int = pipe_max_size + 1

    current_pipe_point += 1
    for index, item in enumerate(user_virtual_machines):
        contrast_color = get_contrast_color(color_change, structure)
        current_row = current_pipe_point + index

        virtual_machine_comment: str = virtual_machines_data.get(item, {}).get('comment')

        if not virtual_machine_comment:
            worksheet.write(f'{column}{current_row}', virtual_machine_comment, structure.missing_cell)
        else:
            worksheet.write(f'{column}{current_row}', virtual_machine_comment, contrast_color)
        color_change += 1


def get_machine_comment(console_server_data, machine_name, pipe_name):
    machine_comment: str = console_server_data.get(pipe_name, {}). \
        get('pipe_data', {}).get(machine_name, {}).get('comment', 'None')
    return machine_comment


def get_machine_issues_size(machine_issues):
    machine_issues_total: int = len(machine_issues)

    if machine_issues_total == 0:
        return 1
    else:
        return machine_issues_total


def get_machine_connection_status(console_server_data, machine_name, pipe_name):
    connection_status: str = console_server_data.get(pipe_name, {}). \
        get('pipe_data', {}).get(machine_name, {}).get('connection_status', 'None')
    return connection_status


def get_machine_hyperlink(console_server_data, machine_name, pipe_name):
    host_id: str = console_server_data.get(pipe_name, {}). \
        get('pipe_data', {}).get(machine_name, {}).get('id', 'None')
    host_url: str = f'http://172.30.1.100/console/host_details.php?host_id={host_id}'
    return host_url


def add_trr_column(initial_point, user_sorted_pipes, worksheet, structure, console_server_data):
    """
    Write Pipe Name column in excel output
    :param user_sorted_pipes:
    :param initial_point:
    :param console_server_data: for host group IDs
    :param worksheet:
    :param structure:
    :return:
    """
    column: str = 'L'

    color_change: int = 0
    current_pipe_point: int = initial_point
    for index, pipe_name in enumerate(user_sorted_pipes):
        user_machines = sorted(list(user_sorted_pipes.get(pipe_name, {}).keys()))

        machine_order: int = 0
        pipe_max_size: int = current_pipe_point
        for start, machine_name in enumerate(user_machines):
            machine_issues: dict = user_sorted_pipes.get(pipe_name, {}).get(machine_name)
            machine_issues_total: int = get_machine_issues_size(machine_issues)

            ticket_id: str = get_machine_ticket_id(console_server_data, machine_name, pipe_name)
            ticket_hyperlink: str = get_machine_ticket_hyperlink(console_server_data, machine_name, pipe_name)

            max_number: int = pipe_max_size + machine_issues_total - 1
            number_of_cell = str(machine_order + current_pipe_point)

            contrast_color = get_contrast_color(color_change, structure)

            write_hyperlink_cells(machine_issues_total, column, number_of_cell, max_number, ticket_hyperlink,
                                  ticket_id, worksheet, structure, contrast_color)

            machine_order += machine_issues_total
            pipe_max_size += machine_issues_total
            color_change += 1

        current_pipe_point: int = pipe_max_size + 1


def get_machine_ticket_id(console_server_data, machine_name, pipe_name):
    return console_server_data.get(pipe_name, {}). \
        get('pipe_data', {}).get(machine_name, {}).get('ticket', 'None')


def get_machine_ticket_hyperlink(console_server_data, machine_name, pipe_name):
    ticket_id: str = console_server_data.get(pipe_name, {}). \
        get('pipe_data', {}).get(machine_name, {}).get('ticket', 'None')
    return f'https://azurecsi.visualstudio.com/CSI%20Commodity%20Qualification/_workitems/edit/{ticket_id}'


def add_type_column(initial_point, user_sorted_pipes, worksheet, structure, console_server_data, azure_devops_data):
    """
    Write Pipe Name column in excel output
    :param azure_devops_data:
    :param user_sorted_pipes:
    :param initial_point:
    :param console_server_data: for host group IDs
    :param worksheet:
    :param structure:
    :return:
    """
    column: str = 'M'

    color_change: int = 0
    current_pipe_point: int = initial_point
    for index, pipe_name in enumerate(user_sorted_pipes):
        user_machines = sorted(list(user_sorted_pipes.get(pipe_name, {}).keys()))

        machine_order: int = 0
        pipe_max_size: int = current_pipe_point
        for start, machine_name in enumerate(user_machines):

            machine_issues: dict = user_sorted_pipes.get(pipe_name, {}).get(machine_name)
            machine_issues_total: int = get_machine_issues_size(machine_issues)

            ticket_id: str = get_machine_ticket_id(console_server_data, machine_name, pipe_name)
            ticket_hyperlink: str = get_machine_ticket_hyperlink(console_server_data, machine_name, pipe_name)

            ticket_type: str = get_ticket_type(azure_devops_data, ticket_id)

            max_number: int = pipe_max_size + machine_issues_total - 1
            number_of_cell = str(machine_order + current_pipe_point)

            contrast_color = get_contrast_color(color_change, structure)

            try:
                ticket_type: str = ticket_type.upper().replace(' TEST', '').replace('TEST', '')

                write_hyperlink_cells(machine_issues_total, column, number_of_cell, max_number, ticket_hyperlink,
                                      ticket_type, worksheet, structure, contrast_color)

                machine_order += machine_issues_total
                pipe_max_size += machine_issues_total
                color_change += 1
            except AttributeError:
                write_hyperlink_cells(machine_issues_total, column, number_of_cell, max_number, ticket_hyperlink,
                                      'None', worksheet, structure, contrast_color)

                machine_order += machine_issues_total
                pipe_max_size += machine_issues_total
                color_change += 1

            except TypeError:
                write_hyperlink_cells(machine_issues_total, column, number_of_cell, max_number, ticket_hyperlink,
                                      'None', worksheet, structure, contrast_color)

                machine_order += machine_issues_total
                pipe_max_size += machine_issues_total
                color_change += 1

        current_pipe_point: int = pipe_max_size + 1


def get_ticket_type(azure_devops_data, ticket_id):
    return azure_devops_data.get(ticket_id, {}).get('table_data', {}).get('request_type')


def write_state_column(initial_point, user_sorted_pipes, worksheet, structure, console_server_data, azure_devops_data):
    """
    Write Pipe Name column in excel output
    :param azure_devops_data:
    :param user_sorted_pipes:
    :param initial_point:
    :param console_server_data: for host group IDs
    :param worksheet:
    :param structure:
    :return:
    """
    column: str = 'N'

    current_pipe_point: int = initial_point
    for index, pipe_name in enumerate(user_sorted_pipes):
        user_machines = sorted(list(user_sorted_pipes.get(pipe_name, {}).keys()))

        machine_order: int = 0
        pipe_max_size: int = current_pipe_point
        for start, machine_name in enumerate(user_machines):

            machine_issues: dict = user_sorted_pipes.get(pipe_name, {}).get(machine_name)
            machine_issues_total: int = get_machine_issues_size(machine_issues)

            ticket_id: str = get_machine_ticket_id(console_server_data, machine_name, pipe_name)
            ticket_hyperlink: str = get_machine_ticket_hyperlink(console_server_data, machine_name, pipe_name)
            raw_ticket_state: str = azure_devops_data.get(ticket_id, {}).get('state', {})

            ticket_state = clean_ticket_state(raw_ticket_state)

            max_number: int = pipe_max_size + machine_issues_total - 1
            number_of_cell = str(machine_order + current_pipe_point)

            if ticket_state == 'Done':
                write_hyperlink_cells(machine_issues_total, column, number_of_cell, max_number, ticket_hyperlink,
                                      ticket_state, worksheet, structure, structure.purple_middle)

            elif ticket_state == 'Signed Off':
                write_hyperlink_cells(machine_issues_total, column, number_of_cell, max_number, ticket_hyperlink,
                                      ticket_state, worksheet, structure, structure.purple_middle)

            elif ticket_state == 'Test Completed':
                write_hyperlink_cells(machine_issues_total, column, number_of_cell, max_number, ticket_hyperlink,
                                      ticket_state, worksheet, structure, structure.purple_middle)
            elif ticket_state == 'Planning':
                write_hyperlink_cells(machine_issues_total, column, number_of_cell, max_number, ticket_hyperlink,
                                      ticket_state, worksheet, structure, structure.purple_middle)

            elif ticket_state == 'Blocked':
                write_hyperlink_cells(machine_issues_total, column, number_of_cell, max_number, ticket_hyperlink,
                                      ticket_state, worksheet, structure, structure.light_red_middle_11)

            elif ticket_state == 'On Hold':
                write_hyperlink_cells(machine_issues_total, column, number_of_cell, max_number, ticket_hyperlink,
                                      ticket_state, worksheet, structure, structure.light_red_middle_11)

            elif ticket_state == 'In Progress':
                write_hyperlink_cells(machine_issues_total, column, number_of_cell, max_number, ticket_hyperlink,
                                      ticket_state, worksheet, structure, structure.aqua_middle)

            elif ticket_state == 'Ready to Start':
                write_hyperlink_cells(machine_issues_total, column, number_of_cell, max_number, ticket_hyperlink,
                                      ticket_state, worksheet, structure, structure.aqua_middle)

            elif ticket_state == 'Ready to Review':
                write_hyperlink_cells(machine_issues_total, column, number_of_cell, max_number, ticket_hyperlink,
                                      ticket_state, worksheet, structure, structure.aqua_middle)

            else:
                write_hyperlink_cells(machine_issues_total, column, number_of_cell, max_number, ticket_hyperlink,
                                      ticket_state, worksheet, structure, structure.blue_middle)

            machine_order += machine_issues_total
            pipe_max_size += machine_issues_total

        current_pipe_point: int = pipe_max_size + 1


def clean_ticket_state(ticket_state):
    if not ticket_state:
        return ticket_state
    else:
        return ticket_state.replace('InProgress', 'In Progress'). \
            replace('Test completed', 'Test Completed'). \
            replace('Ready To Review', 'Ready to Review'). \
            replace('Ready to start', 'Ready to Start')


def write_skudoc_column(initial_point, user_sorted_pipes, worksheet, structure, console_server_data, azure_devops_data):
    """
    Write Pipe Name column in excel output
    :param user_sorted_pipes:
    :param initial_point:
    :param console_server_data: for host group IDs
    :param worksheet:
    :param structure:
    :return:
    """

    letter: str = 'P'

    color_change: int = 0
    current_pipe_point: int = initial_point
    for index, pipe_name in enumerate(user_sorted_pipes):
        user_machines = sorted(list(user_sorted_pipes.get(pipe_name, {}).keys()))

        machine_order: int = 0
        pipe_max_size: int = current_pipe_point
        for start, machine_name in enumerate(user_machines):
            machine_issues: dict = user_sorted_pipes.get(pipe_name, {}).get(machine_name)
            machine_issues_total: int = get_machine_issues_size(machine_issues)

            skudoc_path = get_skudoc_path(azure_devops_data, console_server_data, machine_name, pipe_name)

            crd_file_name: str = get_crd_file_name(skudoc_path)

            max_number: int = pipe_max_size + machine_issues_total - 1
            number_of_cell = str(machine_order + current_pipe_point)

            contrast_color = get_contrast_color(color_change, structure)

            # if 'SKUDOC' not in crd_file_name:
            #     worksheet.write(f'{letter}{number_of_cell}', 'Broken', contrast_color)

            if machine_issues_total >= 2:
                if skudoc_path == 'None':
                    worksheet.merge_range(f'{letter}{number_of_cell}:{letter}{max_number}', '', structure.missing_cell)
                else:
                    worksheet.merge_range(f'{letter}{number_of_cell}:{letter}{max_number}', crd_file_name,
                                          structure.missing_cell)
                    worksheet.write_url(f'{letter}{number_of_cell}', skudoc_path, contrast_color,
                                        string=crd_file_name)
            else:
                if skudoc_path == 'None':
                    worksheet.write(f'{letter}{number_of_cell}', '', structure.missing_cell)
                else:
                    worksheet.write(f'{letter}{number_of_cell}', crd_file_name,
                                    structure.missing_cell)
                    worksheet.write_url(f'{letter}{number_of_cell}', skudoc_path, contrast_color,
                                        string=crd_file_name)

            machine_order += machine_issues_total
            pipe_max_size += machine_issues_total
            color_change += 1

        current_pipe_point: int = pipe_max_size + 1


def write_crd_column(initial_point, user_sorted_pipes, worksheet, structure, console_server_data, azure_devops_data):
    """
    Write Pipe Name column in excel output
    """
    letter: str = 'O'
    color_change: int = 0
    current_pipe_point: int = initial_point
    for index, pipe_name in enumerate(user_sorted_pipes, start=0):
        user_machines = sorted(list(user_sorted_pipes.get(pipe_name, {}).keys()))

        machine_order: int = 0
        pipe_max_size: int = current_pipe_point
        for start, machine_name in enumerate(user_machines):
            machine_issues: dict = user_sorted_pipes.get(pipe_name, {}).get(machine_name)
            machine_issues_total: int = get_machine_issues_size(machine_issues)
            max_number: int = pipe_max_size + machine_issues_total - 1
            number_of_cell = str(machine_order + current_pipe_point)
            contrast_color: xlsxwriter = get_contrast_color(color_change, structure)

            crd_path: str = get_crd_path(azure_devops_data, console_server_data, machine_name, pipe_name)
            crd_file_name: str = get_crd_file_name(crd_path)

            if machine_issues_total >= 2:
                if crd_path == 'None':
                    worksheet.merge_range(f'{letter}{number_of_cell}:{letter}{max_number}', '', structure.missing_cell)
                else:
                    worksheet.merge_range(f'{letter}{number_of_cell}:{letter}{max_number}', crd_file_name,
                                          structure.missing_cell)
                    worksheet.write_url(f'{letter}{number_of_cell}', crd_path, contrast_color,
                                        string=crd_file_name)
            else:
                if crd_path == 'None':
                    worksheet.write(f'{letter}{number_of_cell}', '', structure.missing_cell)
                else:
                    worksheet.write(f'{letter}{number_of_cell}', crd_file_name,
                                          structure.missing_cell)
                    worksheet.write_url(f'{letter}{number_of_cell}', crd_path, contrast_color,
                                        string=crd_file_name)

            machine_order += machine_issues_total
            pipe_max_size += machine_issues_total
            color_change += 1

        current_pipe_point: int = pipe_max_size + 1


def get_crd_file_name(crd_path):
    file_name: str = crd_path.split('\\')[-1]
    return file_name[0:8]


def get_crd_path(azure_devops_data, console_server_data, machine_name, pipe_name):
    ticket_id: str = console_server_data.get(pipe_name, {}). \
        get('pipe_data', {}).get(machine_name, {}).get('ticket', 'None')
    return azure_devops_data.get(ticket_id, {}).get('attachment_file_paths', {}).get('crd_drive_path', 'None')


def get_skudoc_path(azure_devops_data, console_server_data, machine_name, pipe_name):
    ticket_id: str = console_server_data.get(pipe_name, {}). \
        get('pipe_data', {}).get(machine_name, {}).get('ticket', 'None')
    return azure_devops_data.get(ticket_id, {}).get('attachment_file_paths', {}).get('skudoc_drive_path', 'None')


def write_item_column(initial_point, user_sorted_pipes, worksheet, structure, console_server_data,
                      user_virtual_machines) -> None:
    """
    Writes Pipe Name column in excel output
    """
    column: str = 'G'

    color_change: int = 0
    current_pipe_point: int = initial_point
    for pipe_name in user_sorted_pipes:
        user_machines: list = get_user_machines(pipe_name, user_sorted_pipes)

        machine_order: int = 0
        pipe_max_size: int = current_pipe_point
        for machine_name in user_machines:

            machine_issues: dict = user_sorted_pipes.get(pipe_name, {}).get(machine_name, '')
            contrast_color: int = get_contrast_color(color_change, structure)
            row = str(machine_order + current_pipe_point)

            machine_issues_count: int = get_machine_issues_count(machine_issues)
            machine_ticket_id: str = get_machine_ticket_id(console_server_data, machine_name, pipe_name)

            if machine_issues_count == 1 and check_missing(machine_ticket_id) == 'None':
                worksheet.write(f'{column}{row}', 'Ticket', contrast_color)

            elif machine_issues_count == 1 and machine_issues == {}:
                worksheet.write(f'{column}{row}', 'System', contrast_color)

            else:

                for issue_index, current_issue in enumerate(machine_issues):
                    original_ticket_data: str = machine_issues.get(current_issue, {}).get('original_ticket_data')
                    machine_issue: str = get_machine_issue(current_issue, machine_issues)
                    actual_row = str(int(row) + issue_index)

                    if 'TRR' in original_ticket_data and 'DIMM' in original_ticket_data and \
                            'QCL' in original_ticket_data:
                        worksheet.write(f'{column}{actual_row}', 'DIMM P/N', contrast_color)

                    elif 'TRR' in original_ticket_data and 'NVMe' in original_ticket_data and \
                            'QCL' in original_ticket_data and 'P/N' in original_ticket_data:
                        worksheet.write(f'{column}{actual_row}', 'NVMe P/N', contrast_color)

                    elif 'TRR' in original_ticket_data and 'NVMe' in original_ticket_data and \
                            'QCL' in original_ticket_data and 'F/W' in original_ticket_data:
                        worksheet.write(f'{column}{actual_row}', 'NVMe F/W', contrast_color)

                    elif 'TRR' in original_ticket_data and 'Disk' in original_ticket_data and \
                            'QCL' in original_ticket_data and 'P/N' in original_ticket_data:
                        worksheet.write(f'{column}{actual_row}', 'Disk P/N', contrast_color)

                    elif 'TRR' in original_ticket_data and 'Disk' in original_ticket_data and \
                            'QCL' in original_ticket_data and 'F/W' in original_ticket_data:
                        worksheet.write(f'{column}{actual_row}', 'Disk F/W', contrast_color)


                    else:
                        add_machine_issue_item(column, contrast_color, machine_issue, actual_row, worksheet)

            machine_issues_total: int = get_machine_issues_size(machine_issues)
            machine_order += machine_issues_total
            pipe_max_size += machine_issues_total
            color_change += 1

        current_pipe_point: int = pipe_max_size + 1

    current_pipe_point += 1
    for index, item in enumerate(user_virtual_machines):
        current_row = current_pipe_point + index
        worksheet.write(f'{column}{current_row}', '', structure.missing_cell)
        color_change += 1


def get_machine_issues_count(machine_issues):
    if machine_issues == {} or len(machine_issues) == 1:
        return 1


def write_category_column(initial_point, user_sorted_pipes, worksheet, structure, tally_storage,
                          console_server_data, user_virtual_machines) -> dict:
    """
    Writes Pipe Name column in excel output
    :param console_server_data:
    :param tally_storage:
    :param user_sorted_pipes:
    :param initial_point:
    :param worksheet: which worksheet
    :param structure: color
    :return:
    """
    column: str = 'F'

    color_change: int = 0
    current_pipe_point: int = initial_point
    for pipe_name in user_sorted_pipes:
        user_machines: list = get_user_machines(pipe_name, user_sorted_pipes)

        machine_order: int = 0
        pipe_max_size: int = current_pipe_point
        for machine_name in user_machines:

            machine_issues: dict = user_sorted_pipes.get(pipe_name, {}).get(machine_name, '')
            contrast_color: int = get_contrast_color(color_change, structure)
            row = str(machine_order + current_pipe_point)

            machine_issues_count: int = get_machine_issues_count(machine_issues)
            machine_ticket_id: str = get_machine_ticket_id(console_server_data, machine_name, pipe_name)

            if machine_issues_count == 1 and check_missing(machine_ticket_id) == 'None':
                worksheet.write(f'{column}{row}', 'EMPTY', structure.light_red_middle_11)
                tally_storage['empty'] += 1

            elif machine_issues_count == 1 and machine_issues == {}:
                worksheet.write(f'{column}{row}', 'MATCH', structure.middle_green_11)
                tally_storage['match'] += 1

            else:

                for issue_index, current_issue in enumerate(machine_issues):
                    issue_state: str = machine_issues.get(current_issue, {}).get('issue_state')
                    system_component: str = machine_issues.get(current_issue, {}).get('system_component')
                    actual_row = str(int(row) + issue_index)

                    result = add_machine_issue_state(column, issue_state, system_component,
                                                     actual_row, worksheet, structure)
                    tally_storage[result] += 1

            machine_issues_total: int = get_machine_issues_size(machine_issues)
            machine_order += machine_issues_total
            pipe_max_size += machine_issues_total
            color_change += 1

        current_pipe_point: int = pipe_max_size + 1

    current_pipe_point += 1
    for index, item in enumerate(user_virtual_machines):
        current_row = current_pipe_point + index
        worksheet.write(f'{column}{current_row}', '', structure.missing_cell)
        color_change += 1

    return tally_storage


def get_machine_issue(current_issue, machine_issues):
    return machine_issues.get(current_issue, {}).get('system_component')


def add_machine_issue_item(column, contrast_color, issue_item, row, worksheet):
    if issue_item == 'BIOS' or issue_item == 'BMC' or issue_item == 'OS' or issue_item == 'System' or \
            issue_item == 'MATCH' or issue_item == 'CPLD':
        worksheet.write(f'{column}{row}', issue_item, contrast_color)

    else:
        worksheet.write(f'{column}{row}', 'Ticket', contrast_color)


def add_machine_issue_state(column, issue_item, system_component, row, worksheet, structure):
    if 'MISMATCH' in issue_item:
        worksheet.write(f'{column}{row}', 'MISMATCH', structure.neutral_cell)
        return 'mismatch'
    elif 'Title' in system_component:
        worksheet.write(f'{column}{row}', 'INCOMPLETE', structure.light_red_middle_11)
        return 'incomplete'
    else:
        worksheet.write(f'{column}{row}', 'INCOMPLETE', structure.light_red_middle_11)
        return 'incomplete'


def get_user_machines(pipe_name, user_sorted_pipes) -> list:
    return sorted(list(user_sorted_pipes.get(pipe_name, {}).keys()))


def write_console_server_column(initial_point, user_clean_pipes, worksheet, structure,
                                console_server_data, user_virtual_machines):
    """
    Writes Pipe Name column in excel output
    :param console_server_data:
    :param user_clean_pipes:
    :param initial_point:
    :param worksheet: which worksheet
    :param structure: color
    :return:
    """
    column: str = 'I'

    color_change: int = 0
    current_pipe_point: int = initial_point
    for pipe_name in user_clean_pipes:
        user_machines: list = get_user_machines(pipe_name, user_clean_pipes)

        machine_order: int = 0
        pipe_max_size: int = current_pipe_point
        for machine_name in user_machines:

            machine_issues: dict = user_clean_pipes.get(pipe_name, {}).get(machine_name, '')
            contrast_color: int = get_contrast_color(color_change, structure)
            row = str(machine_order + current_pipe_point)

            machine_issues_count: int = get_machine_issues_count(machine_issues)
            machine_ticket_id: str = get_machine_ticket_id(console_server_data, machine_name, pipe_name)

            if machine_issues_count == 1 and check_missing(machine_ticket_id) == 'None':
                worksheet.write(f'{column}{row}', 'Ticket Field - No TRR ID', contrast_color)

            elif machine_issues_count == 1 and machine_issues == {}:
                worksheet.write(f'{column}{row}', '', structure.missing_cell)

            else:

                for issue_index, current_issue in enumerate(machine_issues):
                    original_system_data: str = machine_issues.get(current_issue, {}).get('original_system_data')
                    actual_row = str(int(row) + issue_index)

                    if not original_system_data or 'NONE' in original_system_data.upper():
                        worksheet.write(f'{column}{actual_row}', '', structure.missing_cell)
                    else:
                        worksheet.write(f'{column}{actual_row}', original_system_data, contrast_color)

            machine_issues_total: int = get_machine_issues_size(machine_issues)
            machine_order += machine_issues_total
            pipe_max_size += machine_issues_total
            color_change += 1

        current_pipe_point: int = pipe_max_size + 1

    current_pipe_point += 1
    for index, item in enumerate(user_virtual_machines):
        contrast_color = get_contrast_color(color_change, structure)
        current_row = current_pipe_point + index

        pipe_name: str = user_virtual_machines.get(item, {}).get('pipe_name', '')

        if not pipe_name:
            worksheet.write(f'{column}{current_row}', '', structure.missing_cell)
        else:
            worksheet.write(f'{column}{current_row}', pipe_name, contrast_color)
        color_change += 1


def write_azure_column(initial_point, user_clean_pipes, worksheet, structure, console_server_data):
    """
    Writes Pipe Name column in excel output
    :param user_clean_pipes:
    :param initial_point:
    :param worksheet: which worksheet
    :param structure: color
    :return:
    """
    column: str = 'J'

    color_change: int = 0
    current_pipe_point: int = initial_point
    for pipe_name in user_clean_pipes:
        user_machines: list = get_user_machines(pipe_name, user_clean_pipes)

        machine_order: int = 0
        pipe_max_size: int = current_pipe_point
        for machine_name in user_machines:

            machine_issues: dict = user_clean_pipes.get(pipe_name, {}).get(machine_name, '')
            contrast_color: int = get_contrast_color(color_change, structure)
            row = str(machine_order + current_pipe_point)

            machine_issues_count: int = get_machine_issues_count(machine_issues)
            machine_ticket_id: str = get_machine_ticket_id(console_server_data, machine_name, pipe_name)

            if machine_issues_count == 1 and check_missing(machine_ticket_id) == 'None':
                worksheet.write(f'{column}{row}', '', structure.missing_cell)

            elif machine_issues_count == 1 and machine_issues == {}:
                worksheet.write(f'{column}{row}', '', structure.missing_cell)

            else:

                for issue_index, current_issue in enumerate(machine_issues):
                    original_ticket_data: str = machine_issues.get(current_issue, {}).get('original_ticket_data')
                    actual_row = str(int(row) + issue_index)

                    if not original_ticket_data:
                        worksheet.write(f'{column}{actual_row}', '', structure.missing_cell)
                    else:
                        worksheet.write(f'{column}{actual_row}', original_ticket_data, contrast_color)

            machine_issues_total: int = get_machine_issues_size(machine_issues)
            machine_order += machine_issues_total
            pipe_max_size += machine_issues_total
            color_change += 1

        current_pipe_point: int = pipe_max_size + 1


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
    return all_unique_machines


def add_issue_data(azure_devops_data: dict, console_server_data: dict, all_issues: list, current_setup: dict) -> None:
    """

    :param current_setup:
    :param console_server_data:
    :param azure_devops_data:
    :param all_issues:
    """
    default_user_name: str = current_setup.get('default_user_name')
    initial_point: int = current_setup.get('header_height') + 1

    worksheet: xlsxwriter = current_setup.get('worksheet')
    structure: xlsxwriter = current_setup.get('structure')

    user_info: dict = get_user_info(console_server_data, default_user_name)

    if user_info == {}:
        pass

    else:
        user_systems: dict = get_user_systems(user_info)
        user_pipes: list = get_user_pipes(user_systems)
        user_unique_pipes: list = get_user_unique_pipes(user_pipes)
        user_virtual_machines: dict = get_user_virtual_machines(user_info)
        user_sorted_pipes: dict = get_sorted_pipes_and_systems(user_systems, user_unique_pipes)
        virtual_machines_data: dict = get_virtual_machines()
        tally_storage: dict = get_tally_storage()
        user_clean_pipes: dict = clean_user_sorted_pipes(all_issues, user_sorted_pipes)

        write_pipe_name_column(initial_point, user_clean_pipes, worksheet, structure, console_server_data,
                               user_virtual_machines)

        tally_storage: dict = add_machine_name_column(initial_point, user_clean_pipes, worksheet, structure,
                                                      console_server_data, user_virtual_machines,
                                                      virtual_machines_data, tally_storage)

        write_status_field_column(initial_point, user_clean_pipes, worksheet, structure,
                                  console_server_data, user_virtual_machines, virtual_machines_data)

        add_trr_column(initial_point, user_clean_pipes, worksheet, structure, console_server_data)

        add_type_column(initial_point, user_clean_pipes, worksheet, structure, console_server_data, azure_devops_data)

        write_state_column(initial_point, user_clean_pipes, worksheet, structure, console_server_data,
                           azure_devops_data)

        write_crd_column(initial_point, user_clean_pipes, worksheet, structure, console_server_data,
                         azure_devops_data)

        write_skudoc_column(initial_point, user_clean_pipes, worksheet, structure, console_server_data,
                            azure_devops_data)

        write_item_column(initial_point, user_clean_pipes, worksheet, structure, console_server_data,
                          user_virtual_machines)

        write_category_column(initial_point, user_clean_pipes, worksheet, structure, tally_storage,
                              console_server_data, user_virtual_machines)

        write_console_server_column(initial_point, user_clean_pipes, worksheet, structure,
                                    console_server_data, user_virtual_machines)

        write_azure_column(initial_point, user_clean_pipes, worksheet, structure, console_server_data)


def get_tally_storage():
    tally_storage: dict = {'dead': 0,
                           'alive': 0,
                           'match': 0,
                           'mismatch': 0,
                           'incomplete': 0,
                           'empty': 0,
                           'total_machines': 0,
                           'virtual_machines': 0,
                           'pipes': 0}
    return tally_storage


def clean_user_sorted_pipes(all_issues, user_sorted_pipes) -> dict:
    count: int = 0
    for user_pipe in user_sorted_pipes:
        for user_machine in user_sorted_pipes[user_pipe]:
            for issue_system in all_issues:
                issue_machine: str = issue_system.get('machine_name')

                if issue_machine in user_machine:
                    user_sorted_pipes[user_pipe][issue_machine][count] = issue_system
                    count += 1
    return user_sorted_pipes


def get_sorted_pipes_and_systems(user_systems, user_unique_pipes):
    """

    :param user_systems:
    :param user_unique_pipes:
    :return:
    """
    user_sorted_pipes: dict = get_user_sorted_pipes(user_unique_pipes)
    for user_system in user_systems:
        pipe_name: str = user_systems.get(user_system, {}).get('pipe_name')
        user_sorted_pipes[pipe_name][user_system] = {}

    return user_sorted_pipes


def get_user_sorted_pipes(user_unique_pipes) -> dict:
    user_sorted_pipes: dict = {}
    for unique_pipe in user_unique_pipes:
        user_sorted_pipes[unique_pipe] = {}
    return user_sorted_pipes


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


class ExcelStructure:
    def __init__(self, current_setup: dict, console_server_data: dict):
        self.current_setup: dict = current_setup
        self.console_server_data: dict = console_server_data

        self.worksheet: xlsxwriter = self.current_setup.get('worksheet')
        self.structure: xlsxwriter = self.current_setup.get('structure')

        self.sheet_title: str = self.current_setup.get('sheet_title')
        self.site_location: str = self.current_setup.get('site_location')
        self.default_user_name: str = self.current_setup.get('default_user_name')
        self.version: str = self.current_setup.get('version')
        self.header_height: str = self.current_setup.get('header_height')


def set_sheet_structure(current_setup: dict, console_server_data: dict) -> None:
    """
    Create dashboard structure
    """
    set_excel_design(current_setup)
    add_header_data(console_server_data, current_setup)


def add_header_keys(current_setup: dict) -> None:
    """

    """
    worksheet: xlsxwriter = current_setup.get('worksheet')
    structure: xlsxwriter = current_setup.get('structure')

    worksheet.merge_range('F6:G6', 'KEY', structure.teal_middle_14)

    worksheet.write('F7', f'GOOD', structure.middle_green_12)
    worksheet.write('F8', f'CONFLICT', structure.light_red_middle_12)
    worksheet.write('F9', f'CAUTION', structure.neutral_cell_12)

    worksheet.write('G7', f'VSE', structure.aqua_middle_12)
    worksheet.write('G8', f'CLIENT', structure.purple_middle_12)
    worksheet.write('G9', f'OFFLINE', structure.light_grey_middle_12_bold)


def add_header_data(console_server_data: dict, current_setup: dict) -> None:
    """
    Add header data on ex. username, date, version, etc.
    """
    add_header_user_name(current_setup)
    add_header_sheet_title(current_setup)
    add_header_site_location(current_setup)
    add_header_date_and_version(current_setup)
    add_header_items_under_testing(current_setup)
    add_header_keys(current_setup)
    add_header_user_info(console_server_data, current_setup)


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
    sheet_title: str = current_setup.get('sheet_title')

    worksheet: xlsxwriter = current_setup.get('worksheet')
    structure: xlsxwriter = current_setup.get('structure')

    worksheet.write('B5', f'  Pipe Cleaner - {sheet_title}', structure.big_blue_font)


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


def add_header_user_info(console_server_data, current_setup) -> None:
    """
    
    :param console_server_data: 
    :param current_setup: 
    :return: 
    """
    add_user_info_titles(current_setup)
    add_user_info_totals(console_server_data, current_setup)


def add_user_info_totals(console_server_data: dict, current_setup: dict):
    default_user_name: str = current_setup.get('default_user_name')

    user_info: dict = get_user_info_alt(console_server_data, default_user_name)

    add_user_pipes_total(user_info, current_setup)
    add_user_hosts_total(user_info, current_setup)
    add_user_virtual_machines_total(user_info, current_setup)


def add_user_virtual_machines_total(user_info: dict, current_setup: dict):
    worksheet: xlsxwriter = current_setup.get('worksheet')
    structure: xlsxwriter = current_setup.get('structure')

    try:
        hosts_vms: int = len(get_user_virtual_machines(user_info))
        worksheet.write('G4', hosts_vms, structure.pale_teal_middle_12)

    except KeyError:
        worksheet.write('G4', f'None', structure.pale_teal_middle_12)


def get_user_virtual_machines(user_info):
    return user_info['virtual_machines']


def add_user_hosts_total(user_info: dict, current_setup: dict):
    worksheet: xlsxwriter = current_setup.get('worksheet')
    structure: xlsxwriter = current_setup.get('structure')

    try:
        hosts_total: int = len(user_info['systems'])
        worksheet.write('G3', hosts_total, structure.pale_teal_middle_12)

    except KeyError:
        worksheet.write('G3', 'None', structure.pale_teal_middle_12)


def add_user_pipes_total(user_info: dict, current_setup: dict):
    worksheet: xlsxwriter = current_setup.get('worksheet')
    structure: xlsxwriter = current_setup.get('structure')

    try:
        user_pipes_total: int = get_user_pipe_total(user_info)

        worksheet.write('G2', user_pipes_total, structure.pale_teal_middle_12)

    except KeyError:
        worksheet.write('G2', f'None', structure.pale_teal_middle_12)


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


def add_user_info_titles(current_setup: dict):
    worksheet: xlsxwriter = current_setup.get('worksheet')
    structure: xlsxwriter = current_setup.get('structure')

    worksheet.write('F2', f'Pipes', structure.teal_middle_14)
    worksheet.write('F3', f'Hosts', structure.teal_middle_14)
    worksheet.write('F4', f'VMs', structure.teal_middle_14)


def add_header_items_under_testing(current_setup: dict) -> None:
    """
    These items under testing are meant to be components still not 100% confident
    """
    worksheet: xlsxwriter = current_setup.get('worksheet')
    structure: xlsxwriter = current_setup.get('structure')
    header_height: xlsxwriter = current_setup.get('header_height')
    upper_header: str = header_height - 1

    worksheet.write('E2', f'ITEMS UNDER TESTING', structure.light_red_middle_14)
    worksheet.write('E3', f'BIOS, BMC, CPLD, OS, Ticket', structure.pale_red_middle_12)
    worksheet.write('E4', f'Configured Systems, Virtual Machines', structure.pale_red_middle_12)
    worksheet.write('E5', f'DIMM - Part Number', structure.pale_red_middle_12)
    worksheet.write('E6', f'SSD/HDD - Part Number / Firmware', structure.pale_red_middle_12)
    worksheet.write('E7', f'NVMe - Part Number / Firmware', structure.pale_red_middle_12)

    worksheet.merge_range(f'C{upper_header}:G{upper_header}', f'Veritas Services & Engineering - Console Server',
                          structure.teal_middle_14)
    worksheet.merge_range(f'I{upper_header}:J{upper_header}', f'Compared Data', structure.teal_middle_14)
    worksheet.merge_range(f'L{upper_header}:P{upper_header}', f'Client - TRRs - ADO', structure.teal_middle_14)


def get_user_pipes(user_systems):
    all_pipes: list = []
    for item in user_systems:
        if 'VSE' in item and '-' in item:
            all_pipes.append(user_systems[item]['pipe_name'])
    return all_pipes


def get_user_systems(user_info) -> dict:
    return user_info['systems']


def get_user_info(console_server_data, default_name) -> dict:
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
        import sys
        print(f'\n')
        print(f'\tDear {default_name.replace(".", " ").title()},')
        print(f'\tPipe Cleaner did not detect any machines checked out under {default_name} within Console Server.')
        print(f'\tPlease checkout a system in order to use Personal Issues page.')
        print(f'\n\tPress enter to continue...')
        input()
        return {}


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

    worksheet.insert_image('A1', 'pipe_cleaner/img/vse_logo.png')


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

    worksheet.freeze_panes(header_height, 7)


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
    excel_setup['sheet_title']: str = 'Machine Ownership'

    excel_setup['worksheet']: xlsxwriter = workbook.add_worksheet(excel_setup.get('sheet_title'))

    excel_setup['rows_height']: tuple = (12.0, 19.5, 19.5, 18.0, 20.25, 20.25, 20.25, 20.25, 20.25, 3.75, 3.75, 18.75)

    excel_setup['columns_width']: tuple = (0.5, 0.5, 24.0, 20.0, 36.0, 16.0, 12.0, 1.86, 49.0, 49.0, 1.86, 9.0,
                                           11.0, 16.0, 13.0, 13.0, 10.0, 10.0, 10.0, 10.0, 10.0, 10.0)

    excel_setup['column_names']: tuple = ('Pipe',
                                          'Machine',
                                          'Status Field',
                                          'Category',
                                          'Item',
                                          '',
                                          'Console Server',
                                          'Client',
                                          '',
                                          'TRR',
                                          'Qual',
                                          'State',
                                          'CRD',
                                          'Skudoc')

    return excel_setup


def main_method(azure_devops_data: dict, console_server_data: dict, excel_setup: dict, all_issues: list) -> None:
    """
    Create Personal Issues
    """
    current_setup: dict = create_personal_issues_sheet(excel_setup)

    set_sheet_structure(current_setup, console_server_data)

    add_issue_data(azure_devops_data, console_server_data, all_issues, current_setup)

    remove_excel_green_corners(current_setup)
