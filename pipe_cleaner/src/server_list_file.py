from pipe_cleaner.src.data_access import get_reservations
from colorama import Fore, Style
import os

from pipe_cleaner.src.terminal_properties import terminal_header_section


def distribute_data(pipe_info: list) -> list:
    """
    Gather
    :param pipe_info: Bundles of Machine Names, U-Slot, IP Address
    :return: Remote Blade Info
    """
    pipe_systems = []

    for blade in pipe_info:

        # Note: Index 0 = Name, 1 = Slot, 2 = Address
        system = remote_blade_info(blade[0], blade[1], blade[2])

        pipe_systems.append(system)

    return pipe_systems


def remote_blade_info(name: str, slot: str, address: str) -> str:
    """
    WCS Test Remote Blade Info
    :param slot: slot in console server ex. 03, 05, etc.
    :param address: ip address ex. 192.168.239.189
    :param name: machine name ex. VSE0GIWCPT-037
    :return:
    """
    # In order to do f-strings, must substitute curly brackets
    open_curly = '{'
    close_curly = '}'

    blade = f"@{open_curly} slot = {slot};    Address = '{address}'  ;  Name = '{name}'{close_curly}"

    return blade


def get_name(console_server_json) -> list:
    """
    Remove VM/ CMA/ Other
    :return: Machine Name List
    """
    correct_machine_name: list = []

    for blade in console_server_json:

        # Checks Proper Machine Name ie. 15 characters-long, 5th character is G, 12th character is -
        if len(blade['machine_name']) == 15 and blade['machine_name'][4] == 'G' and blade['machine_name'][11] == '-':
            correct_machine_name.append(blade['machine_name'])

    return correct_machine_name


def get_slot(console_server_json, machine_names) -> dict:
    """
    Get U Height of Location from Console Server
    :return: Machine Name to U Slot
    """
    name_to_slot: dict = {}

    for name in machine_names:
        for blade in console_server_json:
            if name in blade['machine_name']:

                # Get U Height of Location
                name_to_slot[name] = blade['location'][4:6]

    return name_to_slot


def get_address(console_server_json, machine_names) -> dict:
    """
    Get U Height of Location from Console Server
    :return: Machine Name to U Slot
    """
    name_to_address: dict = {}

    for name in machine_names:
        for blade in console_server_json:
            if name in blade['machine_name']:
                # Get U Height of Location
                name_to_address[name] = blade['host_ip']

    return name_to_address


def merge_data(machine_names: list, name_to_slot: dict, name_to_address: dict):
    """

    :param machine_names:
    :param name_to_slot:
    :param name_to_address:
    :return:
    """
    # Index 0 = Name, 1 = Slot, 2 = Address
    server_list: list = []

    for name in machine_names:

        # Contains Machine Name, Slot, Address
        bundle: list = []

        #  Add Machine Name
        bundle.append(name)

        # Add Slot
        slot = name_to_slot.get(name)
        bundle.append(slot)

        # Add Address
        address = name_to_address.get(name)
        bundle.append(address)

        server_list.append(bundle)

    return server_list


def get_rack_manager(machine_name: str) -> str:
    """
    Get Generation based from Machine Name Naming Convention
    :param machine_name: one machine name needed
    :return: 5 = CMA, 6 >= RM
    """
    generation_number = machine_name[5]

    if generation_number == 5:
        return 'CMA'
    else:
        return 'RM'


def process_pipe_name(host_group_name: str):
    """
    Help find Test DHCP Information
    :param host_group_name:
    :return:
    """
    potential_components: list = []

    # 5 Characters
    position_00 = '('
    position_01 = 'R'
    position_02 = []
    position_03 = []
    position_04 = []

    host_group = host_group_name.upper()

    if '(R' in host_group and 'B' not in host_group:
        slice_line = [character for character in host_group_name]
        for index, character in enumerate(slice_line):

            if '(' in character:

                index_00 = index
                index_01 = index + 1
                index_02 = index + 2
                index_03 = index + 3

                if slice_line[index_00] != '(':
                    continue
                if slice_line[index_01] != 'R':
                    continue

                if index_00 < 0:
                    continue

                position_02.append(slice_line[index_02])
                position_03.append(slice_line[index_03])

                potential_components.append(f'{position_02[0]}'
                                            f'{position_03[0]}')
        return potential_components

    elif '(RB' in host_group:
        slice_line = [character for character in host_group_name]
        for index, character in enumerate(slice_line):

            if '(' in character:

                index_00 = index
                index_01 = index + 1
                index_02 = index + 2
                index_03 = index + 3
                index_04 = index + 4

                if slice_line[index_00] != '(':
                    continue
                if slice_line[index_01] != 'R':
                    continue
                if slice_line[index_02] != 'B':
                    continue

                if index_00 < 0:
                    continue

                position_02.append(slice_line[index_02])
                position_03.append(slice_line[index_03])
                position_04.append(slice_line[index_04])

                potential_components.append(f'{position_02[0]}'
                                            f'{position_03[0]}'
                                            f'{position_04[0]}')

        return potential_components[0]

    else:
        return None


def parse_b_number(rack_number: str, pipe_name: str) -> str:
    """
    Extract the B Row from the Pipe Name in order to find the DHCP Reservation and IP Address associated with it.
    :param pipe_name:
    :param rack_number:
    :return:
    """

    if rack_number is None:
        print(f'WARNING: {Fore.RED}Rack Number is None{Style.RESET_ALL} for {pipe_name}')
        return rack_number

    elif 'B' in rack_number and len(rack_number) == 4:
        slice_number = rack_number[-2:]
        number_1 = slice_number[0]
        number_2 = slice_number[1]
        number_3 = slice_number[2]
        add_numbers = number_1 + number_2 + number_3
        return add_numbers

    elif 'B' in rack_number and len(rack_number) == 3:
        slice_number = rack_number[-2:]
        number_1 = slice_number[0]
        number_2 = slice_number[1]
        add_numbers = '0' + number_1 + number_2
        return add_numbers

    else:
        return rack_number


def find_dhcp_name(dhcp_json, rack_manager, rack_number, check_b_number) -> list:
    """

    :param check_b_number:
    :param dhcp_json:
    :param rack_manager:
    :param rack_number:
    :return:
    """
    potential_name: list = []

    if 'B' in rack_number:
        for dhcp_name in dhcp_json:
            if 'B' in dhcp_name['name'] and check_b_number in dhcp_name['name'] and rack_manager in dhcp_name['name']:
                print(dhcp_name['name'])
                potential_name.append(dhcp_name['name'])
    else:
        for dhcp_name in dhcp_json:
            if rack_manager in dhcp_name['name'] and rack_number[0] in dhcp_name['name']:
                potential_name.append(dhcp_name['name'])

    if len(potential_name) > 1:
        print(f'\tWARNING: Cannot find {Fore.RED}UNIQUE DHCP Name{Style.RESET_ALL}...')
        print(f'\tAcceptable Naming Conventions from Host Groups:')
        print(f'\t\t- RB04-Low')
        print(f'\t\t- R33-Low')
    else:
        return potential_name


def find_dhcp_address(dhcp_json: list, dhcp_name: str):
    """
    Finds DHCP IP Address from Console Server
    :param dhcp_json:
    :param dhcp_name:
    :return:
    """
    for name in dhcp_json:
        if dhcp_name in name['name']:
            return name['ip']


def merge_dhcp_info(name, address):
    """

    :param name:
    :param address:
    :return:
    """
    dhcp_info: list = [name, address]
    return dhcp_info


def check_server_list_folders(users_path: str, user_info: dict):
    """
    Checks if server_list_folder is in the Z_Drive
    :param users_path:
    :param user_info:
    :return:
    """
    user_name = user_info['default_name']

    # Check if root Pipe Cleaner Folder in Z Drive for users is still there
    try:
        os.mkdir(users_path)
    except FileExistsError:
        pass

    # Creates folder for user if no t already
    try:
        os.mkdir(fr'{users_path}\{user_name}')
    except FileExistsError:
        pass

    return fr'{users_path}\{user_name}'


def warn_dhcp(user_info, blade_info, pipe_name, dhcp_names, dhcp_json, machine_name):
    """
    If DHCP doesn't have one reservation or more than 2, Warn user
    :return:
    """
    if len(dhcp_names) > 1:
        print(f'\tWARNING: {Fore.RED}More than one{Style.RESET_ALL} DHCP reservation:'
              f'\n\t\t- Based on rack location and pipe number...')
    elif len(dhcp_names) == 0:
        print(f'\tWARNING: {Fore.RED}No{Style.RESET_ALL} DHCP reservation:'
              f'\n\t\t- Based on rack location and pipe number...')
    else:
        create_power_shell(user_info, blade_info, pipe_name, dhcp_names, dhcp_json, machine_name)


def create_power_shell(user_info: dict, server_list: list, pipe_name: str,  dhcp_names: list, dhcp_json: list,
                       machine_name: str):
    """
    Create PowerShell Script
    :return:
    """
    server_list_template: str = "settings/server_list_template.ps1"
    z_drive_users: str = r'Z:\Kirkland_Lab\PipeCleaner_Users'

    user_path: str = check_server_list_folders(z_drive_users, user_info)

    # Create Server List info
    insert_server_list: str = ''
    dhcp_information: str = f"'{find_dhcp_address(dhcp_json, dhcp_names[0])}' # {dhcp_names[0]}"

    for server in server_list:
        insert_server_list += f'\t{server}\n'

    # Had to use 3 Context Managers because information will only write once with context manager for some reason
    with open(server_list_template, 'r') as file:
        file_data_1 = file.read()

    with open(fr'{user_path}\Pipe-{pipe_name}_ServerList.ps1', 'w') as file:
        server_list = file_data_1.replace('# {{insert_server_list}}', insert_server_list)
        file.write(server_list)

    with open(fr'{user_path}\Pipe-{pipe_name}_ServerList.ps1', 'r') as file:
        file_data_2 = file.read()

    with open(fr'{user_path}\Pipe-{pipe_name}_ServerList.ps1', 'w') as file:
        server_dhcp = file_data_2.replace('# {{insert_server_dhcp}}', dhcp_information)
        file.write(server_dhcp)

    with open(fr'{user_path}\Pipe-{pipe_name}_ServerList.ps1', 'r') as file:
        file_data_3 = file.read()

    with open(fr'{user_path}\Pipe-{pipe_name}_ServerList.ps1', 'w') as file:

        if machine_name == 'RM':
            request_type = 'M2010'
        else:
            request_type = 'Legacy'
        server_dhcp = file_data_3.replace('# {{insert_type}}', request_type)
        file.write(server_dhcp)


def main_method(console_server_json: list, host_group_name: str, user_info: dict, pipe_name: str) -> None:
    """
    Main method for Server List
    :param pipe_name:
    :param user_info: person using the PowerShell Script
    :param host_group_name:
    :param console_server_json:
    :return:
    """
    terminal_header_section("Automated Server List (Beta)", "Preparing for Engineer Testing")

    print(f'\tGathering {Fore.GREEN}Systems and DHCP Info{Style.RESET_ALL} from Console Server...\n')

    # Gather Machine Name, Slot, and Address for Blade Info
    machine_names: list = get_name(console_server_json)
    name_to_slot: dict = get_slot(console_server_json, machine_names)
    name_to_address: dict = get_address(console_server_json, machine_names)

    # Merge Machine Name, Slot, and Address
    pipe_info = merge_data(machine_names, name_to_slot, name_to_address)

    # Get WCS Test Remote Blade Info
    blade_info: list = distribute_data(pipe_info)

    # Gather DHCP JSON, RM/CMA, DHCP Name, and IPV4 Address for WCS Test Remote Mgr
    dhcp_json: list = get_reservations('pipe_cleaner/data/console_server_dhcp.json')
    rack_manager: str = get_rack_manager(machine_names[0])
    rack_number: str = process_pipe_name(host_group_name)
    check_b_number = parse_b_number(rack_number, pipe_name)

    # Find Correct DHCP Name and IP Address
    dhcp_names: list = find_dhcp_name(dhcp_json, rack_manager, rack_number, check_b_number)

    print(f'\n\tWCS Test Remote Blade Info:\n')
    for system in blade_info:
        print(f'\t\t{system}')

    print(f'\n\tPotential WCS Test Remote Mgr:\n')
    try:
        for dhcp_name in dhcp_names:
            print(f'\t\t- {dhcp_name} | {find_dhcp_address(dhcp_json, dhcp_name)}')
    except TypeError:
        print(f'\tWARNING: No DHCP found in Console Server... Press Enter to close down.')
        input()

    warn_dhcp(user_info, blade_info, pipe_name, dhcp_names, dhcp_json, machine_names[0])
