import getpass
from time import strftime, localtime

import xlsxwriter

from pipe_cleaner.src.dashboard_main_setup import process_pipe_name
from pipe_cleaner.src.data_access import get_all_hosts_console_server

pipe_information: dict = {}


def check_missing(data: str) -> str:
    """

    :param data:
    :return:
    """
    try:
        clean_data = data.replace(' ', '')
        if clean_data == 'None' or clean_data == '' or clean_data is None:
            return 'None'
        else:
            return data
    except AttributeError:
        return 'None'


def write_pipe_column(letter: str, min_number: int, worksheet, structure, console_server_data: dict,
                      available_virtual_machines: dict):
    """
    Write host group column in excel output

    :param available_virtual_machines:
    :param min_number:
    :param letter: location of cell
    :param worksheet:
    :param structure:
    :param console_server_data: for host group IDs
    :return:
    """
    new_min_number: int = min_number
    virtual_machine_data: list = console_server_data.get('virtual_machine_data')

    # Increments if pipe name changes
    color_number: int = 0
    previous_pipe_name: str = ''
    for virtual_machine in virtual_machine_data:
        current_pipe_name: str = virtual_machine.get('pipe_name')
        host_id: str = virtual_machine.get('host_id')
        host_group_url: str = f'http://172.30.1.100/console/host_details.php?host_id={host_id}'
        hold_previous_pipe_name: str = previous_pipe_name

        max_number: int = new_min_number - 1
        adjust_height: int = new_min_number - 1

        if previous_pipe_name != current_pipe_name:
            color_number += 1

        previous_pipe_name = current_pipe_name

        if check_missing(current_pipe_name) == 'None':
            worksheet.write_url(f'{letter}{new_min_number}', host_group_url, structure.dark_grey_missing, string='')

        else:
            processed_name: str = process_pipe_name(current_pipe_name)
            color_code: int = color_number % 2

            if color_code == 1:

                worksheet.write_url(f'{letter}{new_min_number}', host_group_url, structure.blue_left_18,
                                    string=f'   {processed_name}')
            else:
                worksheet.write_url(f'{letter}{new_min_number}', host_group_url, structure.alt_blue_left_18,
                                    string=f'   {processed_name}')

        if current_pipe_name == hold_previous_pipe_name:
            worksheet.set_row(adjust_height, 30)
            worksheet.set_row(new_min_number, 0)
            new_min_number: int = max_number + 4

        elif hold_previous_pipe_name == '':
            worksheet.set_row(adjust_height, 30)
            worksheet.set_row(new_min_number, 5.25)
            new_min_number: int = max_number + 3
        else:
            new_adjust_height: int = adjust_height - 1
            worksheet.set_row(new_adjust_height, 5.25)
            worksheet.set_row(adjust_height, 30)
            worksheet.set_row(new_min_number, 0)
            new_min_number: int = max_number + 3

    new_min_number += 1

    for available_machine in available_virtual_machines:
        max_number: int = new_min_number - 1
        host_id: str = available_virtual_machines.get(available_machine, {}).get('host_id')
        host_group_url: str = f'http://172.30.1.100/console/host_details.php?host_id={host_id}'

        color_code: int = max_number % 2
        if color_code == 1:
            worksheet.write_url(f'{letter}{new_min_number}', host_group_url, structure.blue_left_11,
                                string=f'    AVAILABLE')
        else:
            worksheet.write_url(f'{letter}{new_min_number}', host_group_url, structure.alt_blue_left_11,
                                string=f'    AVAILABLE')

        new_min_number: int = max_number + 2


def write_pipe_name_column(letter: str, min_number: int, worksheet, structure, console_server_data: dict,
                           available_virtual_machines: dict):
    """
    Write host group column in excel output

    :param available_virtual_machines:
    :param min_number:
    :param letter: location of cell
    :param worksheet:
    :param structure:
    :param console_server_data: for host group IDs
    :return:
    """
    new_min_number: int = min_number
    virtual_machine_data: list = console_server_data.get('virtual_machine_data')

    # Increments if pipe name changes
    color_number: int = 0
    previous_pipe_name: str = ''
    for virtual_machine in virtual_machine_data:
        current_pipe_name: str = virtual_machine.get('pipe_name')
        host_id: str = virtual_machine.get('host_id')
        host_group_url: str = f'http://172.30.1.100/console/host_details.php?host_id={host_id}'
        hold_previous_pipe_name: str = previous_pipe_name

        max_number: int = new_min_number - 1
        adjust_height: int = new_min_number - 1

        if previous_pipe_name != current_pipe_name:
            color_number += 1

        previous_pipe_name = current_pipe_name

        if check_missing(current_pipe_name) == 'None':
            worksheet.write_url(f'{letter}{new_min_number}', host_group_url, structure.dark_grey_missing, string='')

        else:
            color_code: int = color_number % 2

            if color_code == 1:

                worksheet.write_url(f'{letter}{new_min_number}', host_group_url, structure.blue_left_12,
                                    string=f'   {current_pipe_name}')
            else:
                worksheet.write_url(f'{letter}{new_min_number}', host_group_url, structure.alt_blue_left_12,
                                    string=f'   {current_pipe_name}')

        if current_pipe_name == hold_previous_pipe_name:
            worksheet.set_row(adjust_height, 30)
            worksheet.set_row(new_min_number, 0)
            new_min_number: int = max_number + 4

        elif hold_previous_pipe_name == '':
            worksheet.set_row(adjust_height, 30)
            worksheet.set_row(new_min_number, 5.25)
            new_min_number: int = max_number + 3
        else:
            new_adjust_height: int = adjust_height - 1
            worksheet.set_row(new_adjust_height, 5.25)
            worksheet.set_row(adjust_height, 30)
            worksheet.set_row(new_min_number, 0)
            new_min_number: int = max_number + 3

    new_min_number += 1

    for available_machine in available_virtual_machines:
        max_number: int = new_min_number - 1
        host_id: str = available_virtual_machines.get(available_machine, {}).get('host_id')
        host_group_url: str = f'http://172.30.1.100/console/host_details.php?host_id={host_id}'

        color_code: int = max_number % 2
        if color_code == 1:
            worksheet.write_url(f'{letter}{new_min_number}', host_group_url, structure.blue_left_11,
                                string=f'    AVAILABLE')
        else:
            worksheet.write_url(f'{letter}{new_min_number}', host_group_url, structure.alt_blue_left_11,
                                string=f'    AVAILABLE')

        new_min_number: int = max_number + 2


def write_vm_name_column(letter: str, min_number: int, worksheet, structure, console_server_data: dict,
                         available_virtual_machines: dict):
    """
    Write host group column in excel output

    :param available_virtual_machines:
    :param min_number:
    :param letter: location of cell
    :param worksheet:
    :param structure:
    :param console_server_data: for host group IDs
    :return:
    """
    new_min_number: int = min_number
    virtual_machine_data: list = console_server_data.get('virtual_machine_data')

    # Increments if pipe name changes
    color_number: int = 0
    previous_pipe_name: str = ''
    for virtual_machine in virtual_machine_data:
        current_pipe_name: str = virtual_machine.get('pipe_name')
        machine_name: str = virtual_machine.get('machine_name')
        host_id: str = virtual_machine.get('host_id')
        connection_status: str = virtual_machine.get('connection_status')
        host_group_url: str = f'http://172.30.1.100/console/host_details.php?host_id={host_id}'
        hold_previous_pipe_name: str = previous_pipe_name

        max_number: int = new_min_number - 1
        adjust_height: int = new_min_number - 1

        if previous_pipe_name != current_pipe_name:
            color_number += 1

        previous_pipe_name = current_pipe_name

        if check_missing(current_pipe_name) == 'None':
            worksheet.write_url(f'{letter}{new_min_number}', host_group_url, structure.dark_grey_missing, string='')

        else:
            color_code: int = color_number % 2
            if connection_status == 'dead':
                worksheet.write_url(f'{letter}{new_min_number}', host_group_url, structure.dark_grey_left_11,
                                    string=f'   {machine_name}')
            elif color_code == 1:

                worksheet.write_url(f'{letter}{new_min_number}', host_group_url, structure.blue_left_11,
                                    string=f'   {machine_name}')
            else:
                worksheet.write_url(f'{letter}{new_min_number}', host_group_url, structure.alt_blue_left_11,
                                    string=f'   {machine_name}')

        if current_pipe_name == hold_previous_pipe_name:
            worksheet.set_row(adjust_height, 30)
            worksheet.set_row(new_min_number, 0)
            new_min_number: int = max_number + 4

        elif hold_previous_pipe_name == '':
            worksheet.set_row(adjust_height, 30)
            worksheet.set_row(new_min_number, 5.25)
            new_min_number: int = max_number + 3
        else:
            new_adjust_height: int = adjust_height - 1
            worksheet.set_row(new_adjust_height, 5.25)
            worksheet.set_row(adjust_height, 30)
            worksheet.set_row(new_min_number, 0)
            new_min_number: int = max_number + 3

    new_min_number += 1

    for available_machine in available_virtual_machines:
        max_number: int = new_min_number - 1
        host_id: str = available_virtual_machines.get(available_machine, {}).get('host_id')
        machine_name: str = available_virtual_machines.get(available_machine, {}).get('machine_name')
        connection_status: str = available_virtual_machines.get(available_machine, {}).get('connection_status')
        host_group_url: str = f'http://172.30.1.100/console/host_details.php?host_id={host_id}'

        color_code: int = max_number % 2
        if connection_status == 'dead':
            worksheet.write_url(f'{letter}{new_min_number}', host_group_url, structure.dark_grey_left_11,
                                string=f'   {machine_name}')
        elif color_code == 1:
            worksheet.write_url(f'{letter}{new_min_number}', host_group_url, structure.blue_left_11,
                                string=f'   {machine_name}')
        else:
            worksheet.write_url(f'{letter}{new_min_number}', host_group_url, structure.alt_blue_left_11,
                                string=f'   {machine_name}')

        new_min_number: int = max_number + 2


def write_checkout_column(letter: str, min_number: int, worksheet, structure, console_server_data: dict,
                          available_virtual_machines: dict):
    """
    Write host group column in excel output

    :param available_virtual_machines:
    :param min_number:
    :param letter: location of cell
    :param worksheet:
    :param structure:
    :param console_server_data: for host group IDs
    :return:
    """
    new_min_number: int = min_number
    virtual_machine_data: list = console_server_data.get('virtual_machine_data')

    # Increments if pipe name changes
    color_number: int = 0
    previous_pipe_name: str = ''
    for virtual_machine in virtual_machine_data:
        current_pipe_name: str = virtual_machine.get('pipe_name')
        checked_out_to: str = virtual_machine.get('checked_out_to')
        checked_out_to: str = checked_out_to.lower().replace('.', ' ').title()
        host_id: str = virtual_machine.get('host_id')
        host_group_url: str = f'http://172.30.1.100/console/host_details.php?host_id={host_id}'
        hold_previous_pipe_name: str = previous_pipe_name

        max_number: int = new_min_number - 1
        adjust_height: int = new_min_number - 1

        if previous_pipe_name != current_pipe_name:
            color_number += 1

        previous_pipe_name = current_pipe_name

        if check_missing(current_pipe_name) == 'None':
            worksheet.write_url(f'{letter}{new_min_number}', host_group_url, structure.dark_grey_missing, string='')

        else:
            color_code: int = color_number % 2
            if check_missing(checked_out_to) == 'None':
                worksheet.write_url(f'{letter}{new_min_number}', host_group_url, structure.aqua_missing,
                                    string=f'')
            elif color_code == 1:

                worksheet.write_url(f'{letter}{new_min_number}', host_group_url, structure.blue_left_11,
                                    string=f'   {checked_out_to}')
            else:
                worksheet.write_url(f'{letter}{new_min_number}', host_group_url, structure.alt_blue_left_11,
                                    string=f'   {checked_out_to}')

        if current_pipe_name == hold_previous_pipe_name:
            worksheet.set_row(adjust_height, 30)
            worksheet.set_row(new_min_number, 0)
            new_min_number: int = max_number + 4

        elif hold_previous_pipe_name == '':
            worksheet.set_row(adjust_height, 30)
            worksheet.set_row(new_min_number, 5.25)
            new_min_number: int = max_number + 3
        else:
            new_adjust_height: int = adjust_height - 1
            worksheet.set_row(new_adjust_height, 5.25)
            worksheet.set_row(adjust_height, 30)
            worksheet.set_row(new_min_number, 0)
            new_min_number: int = max_number + 3

    new_min_number += 1

    for available_machine in available_virtual_machines:
        max_number: int = new_min_number - 1
        host_id: str = available_virtual_machines.get(available_machine, {}).get('host_id')
        checked_out_to: str = available_virtual_machines.get(available_machine, {}).get('checked_out_to')
        checked_out_to: str = checked_out_to.lower().replace('.', ' ').title()
        host_group_url: str = f'http://172.30.1.100/console/host_details.php?host_id={host_id}'

        color_code: int = max_number % 2
        if check_missing(checked_out_to) == 'None':
            worksheet.write_url(f'{letter}{new_min_number}', host_group_url, structure.aqua_missing,
                                string=f'')
        elif color_code == 1:
            worksheet.write_url(f'{letter}{new_min_number}', host_group_url, structure.blue_left_11,
                                string=f'   {checked_out_to}')
        else:
            worksheet.write_url(f'{letter}{new_min_number}', host_group_url, structure.alt_blue_left_11,
                                string=f'   {checked_out_to}')

        new_min_number: int = max_number + 2


def write_comment_column(letter: str, min_number: int, worksheet, structure, console_server_data: dict,
                         available_virtual_machines: dict):
    """
    Write host group column in excel output

    :param available_virtual_machines:
    :param min_number:
    :param letter: location of cell
    :param worksheet:
    :param structure:
    :param console_server_data: for host group IDs
    :return:
    """
    new_min_number: int = min_number
    virtual_machine_data: list = console_server_data.get('virtual_machine_data')

    # Increments if pipe name changes
    color_number: int = 0
    previous_pipe_name: str = ''
    for virtual_machine in virtual_machine_data:
        current_pipe_name: str = virtual_machine.get('pipe_name')
        comment: str = virtual_machine.get('comment')
        host_id: str = virtual_machine.get('host_id')
        host_group_url: str = f'http://172.30.1.100/console/host_details.php?host_id={host_id}'
        hold_previous_pipe_name: str = previous_pipe_name

        max_number: int = new_min_number - 1
        adjust_height: int = new_min_number - 1

        if previous_pipe_name != current_pipe_name:
            color_number += 1

        previous_pipe_name = current_pipe_name

        if check_missing(current_pipe_name) == 'None':
            worksheet.write_url(f'{letter}{new_min_number}', host_group_url, structure.dark_grey_missing, string='')

        else:
            color_code: int = color_number % 2
            if check_missing(comment) == 'None':
                worksheet.write_url(f'{letter}{new_min_number}', host_group_url, structure.aqua_missing,
                                    string=f'')
            elif color_code == 1:

                worksheet.write_url(f'{letter}{new_min_number}', host_group_url, structure.blue_left_11,
                                    string=f'   {comment}')
            else:
                worksheet.write_url(f'{letter}{new_min_number}', host_group_url, structure.alt_blue_left_11,
                                    string=f'   {comment}')

        if current_pipe_name == hold_previous_pipe_name:
            worksheet.set_row(adjust_height, 30)
            worksheet.set_row(new_min_number, 0)
            new_min_number: int = max_number + 4

        elif hold_previous_pipe_name == '':
            worksheet.set_row(adjust_height, 30)
            worksheet.set_row(new_min_number, 5.25)
            new_min_number: int = max_number + 3
        else:
            new_adjust_height: int = adjust_height - 1
            worksheet.set_row(new_adjust_height, 5.25)
            worksheet.set_row(adjust_height, 30)
            worksheet.set_row(new_min_number, 0)
            new_min_number: int = max_number + 3

    new_min_number += 1

    for available_machine in available_virtual_machines:
        max_number: int = new_min_number - 1
        host_id: str = available_virtual_machines.get(available_machine, {}).get('host_id')
        comment: str = available_virtual_machines.get(available_machine, {}).get('comment')
        host_group_url: str = f'http://172.30.1.100/console/host_details.php?host_id={host_id}'

        color_code: int = max_number % 2
        if check_missing(comment) == 'None':
            worksheet.write_url(f'{letter}{new_min_number}', host_group_url, structure.aqua_missing,
                                string=f'')
        elif color_code == 1:
            worksheet.write_url(f'{letter}{new_min_number}', host_group_url, structure.blue_left_11,
                                string=f'   {comment}')
        else:
            worksheet.write_url(f'{letter}{new_min_number}', host_group_url, structure.alt_blue_left_11,
                                string=f'   {comment}')

        new_min_number: int = max_number + 2


def write_host_ip_column(letter: str, min_number: int, worksheet, structure, console_server_data: dict,
                         available_virtual_machines: dict):
    """
    Write host group column in excel output

    :param available_virtual_machines:
    :param min_number:
    :param letter: location of cell
    :param worksheet:
    :param structure:
    :param console_server_data: for host group IDs
    :return:
    """
    new_min_number: int = min_number
    virtual_machine_data: list = console_server_data.get('virtual_machine_data')

    # Increments if pipe name changes
    color_number: int = 0
    previous_pipe_name: str = ''
    for virtual_machine in virtual_machine_data:
        current_pipe_name: str = virtual_machine.get('pipe_name')
        host_ip: str = virtual_machine.get('host_ip')
        hold_previous_pipe_name: str = previous_pipe_name

        max_number: int = new_min_number - 1
        adjust_height: int = new_min_number - 1

        if previous_pipe_name != current_pipe_name:
            color_number += 1

        previous_pipe_name = current_pipe_name

        if check_missing(current_pipe_name) == 'None':
            worksheet.write(f'{letter}{new_min_number}', '', structure.dark_grey_missing)

        else:
            color_code: int = color_number % 2
            if check_missing(host_ip) == 'None':
                worksheet.write(f'{letter}{new_min_number}', '', structure.aqua_missing)
            elif color_code == 1:

                worksheet.write(f'{letter}{new_min_number}', host_ip, structure.blue_middle)
            else:
                worksheet.write(f'{letter}{new_min_number}', host_ip, structure.alt_blue_middle)

        if current_pipe_name == hold_previous_pipe_name:
            worksheet.set_row(adjust_height, 30)
            worksheet.set_row(new_min_number, 0)
            new_min_number: int = max_number + 4

        elif hold_previous_pipe_name == '':
            worksheet.set_row(adjust_height, 30)
            worksheet.set_row(new_min_number, 5.25)
            new_min_number: int = max_number + 3
        else:
            new_adjust_height: int = adjust_height - 1
            worksheet.set_row(new_adjust_height, 5.25)
            worksheet.set_row(adjust_height, 30)
            worksheet.set_row(new_min_number, 0)
            new_min_number: int = max_number + 3

    new_min_number += 1

    for available_machine in available_virtual_machines:
        max_number: int = new_min_number - 1
        color_code: int = max_number % 2

        if check_missing(available_machine) == 'None':
            worksheet.write(f'{letter}{new_min_number}', '', structure.aqua_missing)

        elif color_code == 1:
            worksheet.write(f'{letter}{new_min_number}', available_machine, structure.blue_middle)

        else:
            worksheet.write(f'{letter}{new_min_number}', available_machine, structure.alt_blue_middle)
        new_min_number: int = max_number + 2


def write_status_column(letter: str, min_number: int, worksheet, structure, console_server_data: dict,
                        available_virtual_machines: dict):
    """
    Write host group column in excel output

    :param available_virtual_machines:
    :param min_number:
    :param letter: location of cell
    :param worksheet:
    :param structure:
    :param console_server_data: for host group IDs
    :return:
    """
    new_min_number: int = min_number
    virtual_machine_data: list = console_server_data.get('virtual_machine_data')

    # Increments if pipe name changes
    color_number: int = 0
    previous_pipe_name: str = ''
    for virtual_machine in virtual_machine_data:
        current_pipe_name: str = virtual_machine.get('pipe_name')
        connection_status: str = virtual_machine.get('connection_status').upper()
        hold_previous_pipe_name: str = previous_pipe_name

        max_number: int = new_min_number - 1
        adjust_height: int = new_min_number - 1

        if previous_pipe_name != current_pipe_name:
            color_number += 1

        previous_pipe_name = current_pipe_name

        if check_missing(current_pipe_name) == 'None':
            worksheet.write(f'{letter}{new_min_number}', '', structure.dark_grey_missing)

        else:
            color_code: int = color_number % 2
            if check_missing(connection_status) == 'None':
                worksheet.write(f'{letter}{new_min_number}', '', structure.aqua_missing)

            elif connection_status == 'DEAD':
                worksheet.write(f'{letter}{new_min_number}', 'OFFLINE', structure.dark_grey_middle)

            elif color_code == 1:
                worksheet.write(f'{letter}{new_min_number}', 'ONLINE', structure.blue_middle)

            else:
                worksheet.write(f'{letter}{new_min_number}', 'ONLINE', structure.alt_blue_middle)

        if current_pipe_name == hold_previous_pipe_name:
            worksheet.set_row(adjust_height, 30)
            worksheet.set_row(new_min_number, 0)
            new_min_number: int = max_number + 4

        elif hold_previous_pipe_name == '':
            worksheet.set_row(adjust_height, 30)
            worksheet.set_row(new_min_number, 5.25)
            new_min_number: int = max_number + 3
        else:
            new_adjust_height: int = adjust_height - 1
            worksheet.set_row(new_adjust_height, 5.25)
            worksheet.set_row(adjust_height, 30)
            worksheet.set_row(new_min_number, 0)
            new_min_number: int = max_number + 3

    new_min_number += 1

    for available_machine in available_virtual_machines:
        connection_status: str = available_virtual_machines.get(available_machine, {}).get('connection_status').upper()
        max_number: int = new_min_number - 1
        color_code: int = max_number % 2

        if check_missing(available_machine) == 'None':
            worksheet.write(f'{letter}{new_min_number}', '', structure.aqua_missing)

        elif connection_status == 'DEAD':
            worksheet.write(f'{letter}{new_min_number}', 'OFFLINE', structure.dark_grey_middle)

        elif color_code == 1:
            worksheet.write(f'{letter}{new_min_number}', 'ONLINE', structure.blue_middle)

        else:
            worksheet.write(f'{letter}{new_min_number}', 'ONLINE', structure.alt_blue_middle)
        new_min_number: int = max_number + 2


def write_last_online_column(letter: str, min_number: int, worksheet, structure, console_server_data: dict,
                             available_virtual_machines: dict):
    """
    Write host group column in excel output

    :param available_virtual_machines:
    :param min_number:
    :param letter: location of cell
    :param worksheet:
    :param structure:
    :param console_server_data: for host group IDs
    :return:
    """
    new_min_number: int = min_number
    virtual_machine_data: list = console_server_data.get('virtual_machine_data')

    # Increments if pipe name changes
    color_number: int = 0
    previous_pipe_name: str = ''
    for virtual_machine in virtual_machine_data:
        current_pipe_name: str = virtual_machine.get('pipe_name')
        last_found_alive: float = virtual_machine.get('last_found_alive')
        last_found_alive: str = strftime('%Y-%m-%d %H:%M:%S', localtime(last_found_alive))
        hold_previous_pipe_name: str = previous_pipe_name

        max_number: int = new_min_number - 1
        adjust_height: int = new_min_number - 1

        if previous_pipe_name != current_pipe_name:
            color_number += 1

        previous_pipe_name = current_pipe_name

        if check_missing(last_found_alive) == 'None':
            worksheet.write(f'{letter}{new_min_number}', last_found_alive, structure.dark_grey_middle)

        else:
            color_code: int = color_number % 2
            if color_code == 1:
                worksheet.write(f'{letter}{new_min_number}', last_found_alive, structure.blue_middle)

            else:
                worksheet.write(f'{letter}{new_min_number}', last_found_alive, structure.alt_blue_middle)

        if current_pipe_name == hold_previous_pipe_name:
            worksheet.set_row(adjust_height, 30)
            worksheet.set_row(new_min_number, 0)
            new_min_number: int = max_number + 4

        elif hold_previous_pipe_name == '':
            worksheet.set_row(adjust_height, 30)
            worksheet.set_row(new_min_number, 5.25)
            new_min_number: int = max_number + 3
        else:
            new_adjust_height: int = adjust_height - 1
            worksheet.set_row(new_adjust_height, 5.25)
            worksheet.set_row(adjust_height, 30)
            worksheet.set_row(new_min_number, 0)
            new_min_number: int = max_number + 3

    new_min_number += 1

    for available_machine in available_virtual_machines:
        last_found_alive: str = available_virtual_machines.get(available_machine, {}).get('last_found_alive')
        max_number: int = new_min_number - 1
        color_code: int = max_number % 2

        if check_missing(available_machine) == 'None':
            worksheet.write(f'{letter}{new_min_number}', '', structure.aqua_missing)

        elif color_code == 1:
            worksheet.write(f'{letter}{new_min_number}', last_found_alive, structure.blue_middle)

        else:
            worksheet.write(f'{letter}{new_min_number}', last_found_alive, structure.alt_blue_middle)
        new_min_number: int = max_number + 2


def write_vm_host_column(letter: str, min_number: int, worksheet, structure, console_server_data: dict,
                         available_virtual_machines: dict):
    """
    Write host group column in excel output

    :param available_virtual_machines:
    :param min_number:
    :param letter: location of cell
    :param worksheet:
    :param structure:
    :param console_server_data: for host group IDs
    :return:
    """
    new_min_number: int = min_number
    virtual_machine_data: list = console_server_data.get('virtual_machine_data')

    # Increments if pipe name changes
    color_number: int = 0
    previous_pipe_name: str = ''
    for virtual_machine in virtual_machine_data:
        current_pipe_name: str = virtual_machine.get('pipe_name')
        location: str = virtual_machine.get('location')
        hold_previous_pipe_name: str = previous_pipe_name

        max_number: int = new_min_number - 1
        adjust_height: int = new_min_number - 1

        if previous_pipe_name != current_pipe_name:
            color_number += 1

        previous_pipe_name = current_pipe_name

        if check_missing(location) == 'None':
            worksheet.write(f'{letter}{new_min_number}', location, structure.dark_grey_middle)

        else:
            color_code: int = color_number % 2
            if color_code == 1:
                worksheet.write(f'{letter}{new_min_number}', location, structure.blue_middle)

            else:
                worksheet.write(f'{letter}{new_min_number}', location, structure.alt_blue_middle)

        if current_pipe_name == hold_previous_pipe_name:
            worksheet.set_row(adjust_height, 30)
            worksheet.set_row(new_min_number, 0)
            new_min_number: int = max_number + 4

        elif hold_previous_pipe_name == '':
            worksheet.set_row(adjust_height, 30)
            worksheet.set_row(new_min_number, 5.25)
            new_min_number: int = max_number + 3
        else:
            new_adjust_height: int = adjust_height - 1
            worksheet.set_row(new_adjust_height, 5.25)
            worksheet.set_row(adjust_height, 30)
            worksheet.set_row(new_min_number, 0)
            new_min_number: int = max_number + 3

    new_min_number += 1

    for available_machine in available_virtual_machines:
        location: str = available_virtual_machines.get(available_machine, {}).get('location')
        max_number: int = new_min_number - 1
        color_code: int = max_number % 2

        if check_missing(available_machine) == 'None':
            worksheet.write(f'{letter}{new_min_number}', '', structure.aqua_missing)

        elif color_code == 1:
            worksheet.write(f'{letter}{new_min_number}', location, structure.blue_middle)

        else:
            worksheet.write(f'{letter}{new_min_number}', location, structure.alt_blue_middle)
        new_min_number: int = max_number + 2


def write_rdp_column(letter: str, min_number: int, worksheet, structure, console_server_data: dict,
                     available_virtual_machines: dict):
    """
    Write host group column in excel output

    :param available_virtual_machines:
    :param min_number:
    :param letter: location of cell
    :param worksheet:
    :param structure:
    :param console_server_data: for host group IDs
    :return:
    """
    new_min_number: int = min_number
    virtual_machine_data: list = console_server_data.get('virtual_machine_data')

    # Increments if pipe name changes
    color_number: int = 0
    previous_pipe_name: str = ''
    for virtual_machine in virtual_machine_data:
        current_pipe_name: str = virtual_machine.get('pipe_name')
        rdp_connection_string: str = virtual_machine.get('rdp_connection_string')
        rdp_url: str = f'http://172.30.1.100/guacamole/#/client/{rdp_connection_string}'
        hold_previous_pipe_name: str = previous_pipe_name

        max_number: int = new_min_number - 1
        adjust_height: int = new_min_number - 1

        if previous_pipe_name != current_pipe_name:
            color_number += 1

        previous_pipe_name = current_pipe_name

        if check_missing(rdp_connection_string) == 'None':
            worksheet.write(f'{letter}{new_min_number}', 'None', structure.dark_grey_middle)

        else:
            color_code: int = color_number % 2
            if color_code == 1:
                worksheet.write_url(f'{letter}{new_min_number}', rdp_url, structure.blue_middle,
                                    string=f'{rdp_connection_string}')

            else:
                worksheet.write_url(f'{letter}{new_min_number}', rdp_url, structure.alt_blue_middle,
                                    string=f'{rdp_connection_string}')

        if current_pipe_name == hold_previous_pipe_name:
            worksheet.set_row(adjust_height, 30)
            worksheet.set_row(new_min_number, 0)
            new_min_number: int = max_number + 4

        elif hold_previous_pipe_name == '':
            worksheet.set_row(adjust_height, 30)
            worksheet.set_row(new_min_number, 5.25)
            new_min_number: int = max_number + 3
        else:
            new_adjust_height: int = adjust_height - 1
            worksheet.set_row(new_adjust_height, 5.25)
            worksheet.set_row(adjust_height, 30)
            worksheet.set_row(new_min_number, 0)
            new_min_number: int = max_number + 3

    new_min_number += 1

    for available_machine in available_virtual_machines:
        rdp_connection_string: str = available_virtual_machines.get(available_machine, {}).get('rdp_connection_string')
        rdp_url: str = f'http://172.30.1.100/guacamole/#/client/{rdp_connection_string}'

        max_number: int = new_min_number - 1
        color_code: int = max_number % 2

        if check_missing(available_machine) == 'None':
            worksheet.write(f'{letter}{new_min_number}', 'None', structure.aqua_missing)

        elif color_code == 1:
            worksheet.write_url(f'{letter}{new_min_number}', rdp_url, structure.blue_middle,
                                string=f'{rdp_connection_string}')

        else:
            worksheet.write_url(f'{letter}{new_min_number}', rdp_url, structure.alt_blue_middle,
                                string=f'{rdp_connection_string}')
        new_min_number: int = max_number + 2


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


def get_virtual_machines():
    """
    Get virtual machines that do not show up in the host groups. These would be available virtual machines
    :return:
    """
    all_hosts_data: dict = get_all_hosts_console_server()

    virtual_machines: dict = {}
    for host in all_hosts_data:
        machine_name: str = host.get('machine_name')
        if '-VM-' in machine_name:
            host_ip: str = host.get('host_ip')

            virtual_machines[host_ip] = {}
            virtual_machines[host_ip]['checked_out_to'] = host.get('checked_out_to')
            virtual_machines[host_ip]['host_id'] = host.get('id')
            virtual_machines[host_ip]['last_found_alive'] = host.get('last_found_alive')
            virtual_machines[host_ip]['location'] = host.get('location')
            virtual_machines[host_ip]['ticket'] = host.get('ticket')
            virtual_machines[host_ip]['sku_name'] = host.get('sku_name')
            virtual_machines[host_ip]['serial'] = host.get('serial')
            virtual_machines[host_ip]['connection_status'] = host.get('connection_status')
            virtual_machines[host_ip]['machine_name'] = host.get('machine_name')
            virtual_machines[host_ip]['comment'] = host.get('comment')
            virtual_machines[host_ip]['bmc_ip'] = host.get('bmc_ip')
            virtual_machines[host_ip]['vnc_connection_string'] = host.get('vnc_connection_string')
            virtual_machines[host_ip]['rdp_connection_string'] = host.get('rdp_connection_string')
            virtual_machines[host_ip]['ssh_connection_string'] = host.get('ssh_connection_string')
            virtual_machines[host_ip]['sensor_status'] = host.get('sensor_status')

    return virtual_machines


def get_available_virtual_machines(all_virtual_machines: dict, console_server_data: dict):
    """

    :param all_virtual_machines:
    :param console_server_data:
    :return:
    """
    console_server_virtual_machines: list = []
    for virtual_machine_ip in all_virtual_machines:
        console_server_virtual_machines.append(virtual_machine_ip)

    pipe_virtual_machines: dict = console_server_data.get('virtual_machine_data')

    host_group_virtual_machines: list = []
    for virtual_machine in pipe_virtual_machines:
        host_ip: str = virtual_machine.get('host_ip')
        host_group_virtual_machines.append(host_ip)

    available_virtual_machines = list(set(console_server_virtual_machines) - set(host_group_virtual_machines))

    available_virtual_machine_data: dict = {}
    for available_machine_ip in available_virtual_machines:
        for virtual_machine_ip in all_virtual_machines:
            if available_machine_ip in virtual_machine_ip:
                available_virtual_machine_data[available_machine_ip] = {}
                checked_out_to: str = all_virtual_machines.get(available_machine_ip, {}).get('checked_out_to')
                host_id: str = all_virtual_machines.get(available_machine_ip, {}).get('host_id')
                last_found_alive: str = all_virtual_machines.get(available_machine_ip, {}).get('last_found_alive')
                location: str = all_virtual_machines.get(available_machine_ip, {}).get('location')
                ticket: str = all_virtual_machines.get(available_machine_ip, {}).get('ticket')
                sku_name: str = all_virtual_machines.get(available_machine_ip, {}).get('sku_name')
                serial: str = all_virtual_machines.get(available_machine_ip, {}).get('serial')
                connection_status: str = all_virtual_machines.get(available_machine_ip, {}).get('connection_status')
                machine_name: str = all_virtual_machines.get(available_machine_ip, {}).get('machine_name')
                comment: str = all_virtual_machines.get(available_machine_ip, {}).get('comment')
                bmc_ip: str = all_virtual_machines.get(available_machine_ip, {}).get('bmc_ip')
                sensor_status: str = all_virtual_machines.get(available_machine_ip, {}).get('sensor_status')
                vnc_connection_string: str = all_virtual_machines.get(available_machine_ip, {}). \
                    get('vnc_connection_string')
                rdp_connection_string: str = all_virtual_machines.get(available_machine_ip, {}). \
                    get('rdp_connection_string')
                ssh_connection_string: str = all_virtual_machines.get(available_machine_ip, {}). \
                    get('ssh_connection_string')

                available_virtual_machine_data[available_machine_ip]['checked_out_to'] = checked_out_to
                available_virtual_machine_data[available_machine_ip]['host_id'] = host_id
                available_virtual_machine_data[available_machine_ip]['last_found_alive'] = last_found_alive
                available_virtual_machine_data[available_machine_ip]['location'] = location
                available_virtual_machine_data[available_machine_ip]['ticket'] = ticket
                available_virtual_machine_data[available_machine_ip]['sku_name'] = sku_name
                available_virtual_machine_data[available_machine_ip]['serial'] = serial
                available_virtual_machine_data[available_machine_ip]['connection_status'] = connection_status
                available_virtual_machine_data[available_machine_ip]['machine_name'] = machine_name
                available_virtual_machine_data[available_machine_ip]['comment'] = comment
                available_virtual_machine_data[available_machine_ip]['bmc_ip'] = bmc_ip
                available_virtual_machine_data[available_machine_ip]['sensor_status'] = sensor_status
                available_virtual_machine_data[available_machine_ip]['machine_ip'] = available_machine_ip
                available_virtual_machine_data[available_machine_ip]['vnc_connection_string'] = vnc_connection_string
                available_virtual_machine_data[available_machine_ip]['rdp_connection_string'] = rdp_connection_string
                available_virtual_machine_data[available_machine_ip]['ssh_connection_string'] = ssh_connection_string

    all_machine_name: list = []
    for virtual_machine_ip in available_virtual_machine_data:
        machine_name: str = available_virtual_machine_data.get(virtual_machine_ip, {}).get('machine_name')
        all_machine_name.append(machine_name)

    ordered_machine_name: list = sorted(all_machine_name)

    available_vm_data: dict = {}
    for machine_name in ordered_machine_name:
        for virtual_machine_ip in available_virtual_machine_data:
            virtual_machine_name: str = available_virtual_machine_data.get(virtual_machine_ip, {}).get('machine_name')
            if machine_name in virtual_machine_name:
                available_vm_data[virtual_machine_ip] = available_virtual_machine_data.get(virtual_machine_ip, {})

    return available_vm_data


def add_issue_data(available_virtual_machines: dict, console_server_data: dict, worksheet, structure):
    """

    :param available_virtual_machines:
    :param console_server_data:
    :param worksheet:
    :param structure:
    :return:
    """
    # Initial accounts for starting point of the dashboard data
    minimal_header_number: int = 14

    write_pipe_column('C', minimal_header_number, worksheet, structure, console_server_data,
                      available_virtual_machines)

    write_vm_name_column('D', minimal_header_number, worksheet, structure, console_server_data,
                         available_virtual_machines)

    write_checkout_column('E', minimal_header_number, worksheet, structure, console_server_data,
                          available_virtual_machines)

    write_comment_column('F', minimal_header_number, worksheet, structure, console_server_data,
                         available_virtual_machines)

    write_pipe_name_column('H', minimal_header_number, worksheet, structure, console_server_data,
                           available_virtual_machines)

    write_host_ip_column('I', minimal_header_number, worksheet, structure, console_server_data,
                         available_virtual_machines)

    write_status_column('J', minimal_header_number, worksheet, structure, console_server_data,
                        available_virtual_machines)

    write_last_online_column('K', minimal_header_number, worksheet, structure, console_server_data,
                             available_virtual_machines)

    write_vm_host_column('L', minimal_header_number, worksheet, structure, console_server_data,
                         available_virtual_machines)

    write_rdp_column('M', minimal_header_number, worksheet, structure, console_server_data,
                     available_virtual_machines)


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
    worksheet.insert_chart('H1', chart_structure, {'x_scale': 2.42, 'y_scale': 0.84})


def set_issue_structure(worksheet, structure, sheet_title, site_location, total_issues,
                        total_checks, pipe_numbers, pipe_cleaner_version, all_virtual_machines,
                        available_virtual_machines):
    """
    Create dashboard structure
    :param available_virtual_machines:
    :param all_virtual_machines:
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
    worksheet.freeze_panes(top_plane_height, 7)

    while top_plane_height < 500:
        worksheet.set_row(top_plane_height, 16.5, structure.white)
        top_plane_height += 1

    correct_total = int(total_checks) - int(total_issues)
    total_pipes = str(pipe_numbers.get('pipes'))
    total_systems = str(pipe_numbers.get('systems'))

    percentage_correct = str((correct_total / int(total_checks)) * 100)[0:4]

    pipe_cleaner_version = pipe_cleaner_version.split(' ')[0]

    # Top Left Plane
    worksheet.insert_image('A1', 'pipe_cleaner/img/vse_logo.png')
    worksheet.write('B5', f' Pipe Cleaner - {sheet_title}', structure.big_blue_font)
    worksheet.write('B6', f'       {site_location}', structure.bold_italic_blue_font)
    worksheet.write('B7', f'            Pipes - {total_pipes}', structure.bold_italic_blue_font)
    worksheet.write('B8', f'            Checks - {total_checks}', structure.bold_italic_blue_font)
    worksheet.write('B9', f'            Total VMs - {len(all_virtual_machines)}', structure.bold_italic_blue_font)
    worksheet.write('D7', f'Blades - {total_systems}', structure.bold_italic_blue_font)
    worksheet.write('D8', f'Issues - {total_issues}', structure.bold_italic_blue_font)
    worksheet.write('D9', f'Available - {len(available_virtual_machines)}', structure.bold_italic_blue_font)
    worksheet.write('B10', f'            Percentage Correct - {percentage_correct} %', structure.bold_italic_green_font)
    worksheet.write('B11', f'       {date} - {time} - {default_name} - v{pipe_cleaner_version}',
                    structure.italic_blue_font)

    worksheet.merge_range('E6:F6', f'Items Under Testing', structure.red_middle_18)
    worksheet.write('E7', f'RDP Column - Have not been Beta Tested', structure.bold_italic_blue_font)


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
        'VM Name',
        'Checkout',
        'Comment',
        '',
        'Pipe Name',
        'VM IP',
        'Status',
        'Last Online',
        'VM Host',
        'RDP']

    initial = 0
    while initial < len(column_names):
        little = chr(ord('c') + initial)
        letter = str(little).upper()

        # Pipe Column
        if letter == 'C':
            worksheet.write_url(f'{letter}{top_plane_height}', 'http://172.30.1.100/console/host_groups.php',
                                structure.teal_left, f'  {column_names[initial]}')

        elif letter == 'N':
            worksheet.write_url(f'{letter}{top_plane_height}', 'http://172.30.1.100/console/host_groups.php',
                                structure.teal_left, f'  {column_names[initial]}')

        # DHCP Information
        elif letter == 'H' or letter == 'I':
            worksheet.write_url(f'{letter}{top_plane_height}', 'http://172.30.1.100/console/reservations.php',
                                structure.teal_middle, f'{column_names[initial]}')

        # SKU Column
        elif letter == 'D':
            worksheet.write_url(f'{letter}{top_plane_height}', 'http://172.30.1.100/console/host_groups.php',
                                structure.teal_left, f'  {column_names[initial]}')

        # Machine Column
        elif letter == 'E':
            worksheet.write_url(f'{letter}{top_plane_height}', 'http://172.30.1.100/console/host_groups.php',
                                structure.teal_left, f'  {column_names[initial]}')

        # Section Column
        elif letter == 'F':
            worksheet.write_url(f'{letter}{top_plane_height}', '',
                                structure.teal_left, f'  {column_names[initial]}')

        elif letter == 'G':
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

    worksheet.set_column('A:A', 1, structure.white)
    worksheet.set_column('B:B', 1, structure.white)
    worksheet.set_column('C:C', 28, structure.white)
    worksheet.set_column('D:D', 22, structure.white)
    worksheet.set_column('E:E', 25, structure.white)
    worksheet.set_column('F:F', 45, structure.white)
    worksheet.set_column('G:G', 0.75, structure.white)
    worksheet.set_column('H:H', 35, structure.white)
    worksheet.set_column('I:I', 18, structure.white)
    worksheet.set_column('J:J', 12, structure.white)
    worksheet.set_column('K:K', 22, structure.white)
    worksheet.set_column('L:L', 18, structure.white)
    worksheet.set_column('M:M', 24, structure.white)
    worksheet.set_column('N:N', 18, structure.white)
    worksheet.set_column('O:O', 18, structure.white)
    worksheet.set_column('P:P', 18, structure.white)
    worksheet.set_column('Q:Q', 18, structure.white)
    worksheet.set_column('R:R', 27, structure.white)
    worksheet.set_column('S:S', 25, structure.white)
    worksheet.set_column('T:T', 25, structure.white)
    worksheet.set_column('U:U', 25, structure.white)
    worksheet.set_column('V:V', 25, structure.white)
    worksheet.set_column('W:W', 25, structure.white)
    worksheet.set_column('X:X', 25, structure.white)
    worksheet.set_column('Y:Y', 25, structure.white)
    worksheet.set_column('Z:Z', 25, structure.white)


def main_method(console_server_data: dict, workbook, structure, site_location: str, all_issues,
                all_checks, mismatch_tally: str, missing_tally: str, pipe_numbers: dict, pipe_cleaner_version: str):
    """

    :param pipe_cleaner_version:
    :param pipe_numbers:
    :param console_server_data:
    :param workbook:
    :param structure:
    :param site_location:
    :param all_issues:
    :param all_checks:
    :param mismatch_tally:
    :param missing_tally:
    :return:
    """
    sheet_name: str = 'Virtual Machines'
    worksheet_issues = workbook.add_worksheet(sheet_name)
    all_virtual_machines: dict = get_virtual_machines()
    available_virtual_machines: dict = get_available_virtual_machines(all_virtual_machines, console_server_data)

    set_issue_structure(worksheet_issues, structure, sheet_name, site_location, len(all_issues),
                        all_checks, pipe_numbers, pipe_cleaner_version, all_virtual_machines,
                        available_virtual_machines)

    add_issue_data(available_virtual_machines, console_server_data, worksheet_issues, structure)

    create_breakdown_graph(console_server_data, workbook, worksheet_issues, 'All Issues', mismatch_tally, missing_tally)

    # Get rid of all errors showing up in excel cells
    worksheet_issues.ignore_errors({'number_stored_as_text': 'A1:XFD1048576'})
