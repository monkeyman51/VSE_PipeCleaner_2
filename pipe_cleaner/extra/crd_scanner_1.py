import os
from colorama import Fore, Style
from pipe_cleaner.src.data_access import check_valid_request
import win32wnet

#  Erase Machine Name if Ticket is Invalid
erase_name: list = []
erase_target: list = []


def gen_from_name(machine_name: str):
    """
    Get Generation of Machine Name. Checks for correct form for Generation within Machine Name.
    :param machine_name:
    :return:
    """
    crd_gen_5 = 'Gen 5.x'
    crd_gen_6 = 'Gen 6.x'
    crd_gen_7 = 'Gen 7.x'
    crd_gen_8 = 'Gen 8.x'
    crd_gen_9 = 'Gen 9.x'
    crd_gen_10 = 'Gen 10.x'

    character_1 = machine_name[4]
    character_2 = machine_name[5]

    if character_1 == 'G' and character_2.isdigit():
        together = character_1 + character_2
        if together == 'G5':
            return crd_gen_5
        elif together == 'G6':
            return crd_gen_6
        elif together == 'G7':
            return crd_gen_7
        elif together == 'G8':
            return crd_gen_8
        elif together == 'G9':
            return crd_gen_9
        elif together == 'G10':
            return crd_gen_10
    else:
        return None


def break_target_configuration(target_configuration: str) -> list:
    """
    Break down target configuration into components for iterating.
    :param target_configuration:
    :return:
    """
    all_components: list = []

    initial = 0
    while initial < 50:
        try:
            component = target_configuration.split('[')[initial].replace(']', '')
            # Azure, System Types
            if ' ' not in component and '/' not in component:
                all_components.append(component)

            # System Types
            if ' ' in component:
                condition = 0
                while condition < 10:
                    try:
                        broken_component = component.split(' ')[condition]
                        all_components.append(broken_component)
                        condition += 1
                    except IndexError:
                        break

            # Gen Types
            if 'Gen' in component or 'GEN' in component or 'gen' in component:
                crd_gen = 'GEN'
                broken_gen = break_generation_component(component)
                for item in broken_gen:
                    all_components.append(f'{crd_gen}{item}')

            # Suppliers
            if '/' in component:
                condition = 0
                while condition < 10:
                    break_component = component.split('/')[condition]
                    all_components.append(break_component)
                    condition += 1
            initial += 1
        except IndexError:
            initial += 1
            break
        except AttributeError:
            initial += 1
            pass

    no_duplicates = list(set(all_components))
    no_empty = [string for string in no_duplicates if string != ""]

    return no_empty


def break_generation_component(raw_gen_component: str) -> list:
    """
    Break down Gen from Target Configuration for possible Gen numbers.
    :param raw_gen_component: raw
    :return:
    """
    broken_gens: list = []

    generation = 0
    specific_gen = 0
    while generation < 10:
        whole_gen = f'{generation}.{specific_gen}'
        if whole_gen in raw_gen_component:
            broken_gens.append(whole_gen)

        if specific_gen == 10:
            specific_gen = 0
            generation += 1

        if generation == 10:
            break

        specific_gen += 1

    return broken_gens


def access_z_drive():
    """
    In case needing to access Z:Drive, need option to put in sensitive information.
    :return:
    """
    path = r'\\\\172.30.1.100\\pxe\\Kirkland_Lab\\Microsoft_CSI\\Documentation\\CRD'

    print('\n  Please enter Username and Password to access VSE - Z:Drive:')
    username = input('  Username: ')
    password = input('  Password: ')

    try:
        # win32wnet.WNetAddConnection2(win32netcon.RESOURCETYPE_DISK, 'Z:', '\\\\192.168.1.18\\D$', None, username,
        #                              password, 0)
        # win32wnet.WNetAddConnection3(path, win32netcon.RESOURCETYPE_DISK, password, username, 0)
        win32wnet.WNetAddConnection2(0, None, path, None, username, password)
        # win32wnet.WNetAddConnection2(win32netcon.RESOURCETYPE_DISK, password, username, path)
        print('connection established successfully')
    except OSError:
        print('connection not established')

# def wnet_connect(host, username, password):
#     unc = ''.join(['\\\\', host])
#     try:
#         win32wnet.WNetAddConnection2(0, None, unc, None, username, password)
#     except Exception, err:
#         if isinstance(err, win32wnet.error):
#             # Disconnect previous connections if detected, and reconnect.
#             if err[0] == 1219:
#                 win32wnet.WNetCancelConnection2(unc, 0, 0)
#                 return wnet_connect(host, username, password)
#         raise err


def access_generation(crd_path: str, name_to_target: dict, machine_name: str) -> str:
    """
    Access Gen Folder based on Gen from Target Configuration
    :param machine_name: Name to call via Key-Value Pair
    :param crd_path: File Path to CSI CRD
    :param name_to_target: Machine Name to Target Configuration
    :return:
    """
    name_to_gen: dict = {machine_name: gen_from_name(machine_name)}

    try:
        for gen_folder in os.listdir(crd_path):
            if name_to_gen[machine_name] == gen_folder:
                return fr'{crd_path}\{gen_folder}'
    except OSError:
        access_z_drive()


def decode_processor(machine_name: str) -> str:
    """
    Get Processor from Machine Name
    :param machine_name:
    :return: Whole Name of Processor
    """
    processor = machine_name[6]
    if processor == 'I':
        return 'INTEL'
    elif processor == 'A':
        return 'AMD'
    elif processor == 'R':
        return 'ARM'
    else:
        return 'None'


def decode_system(machine_name: str) -> str:
    """
    Get System Type from Machine Name
    :param machine_name:
    :return: Whole Name of Processor
    """
    system_type = {'BAL': 'BALANCED',
                   'CPT': 'COMPUTE',
                   'OPT': 'OPTIMIZED',
                   'BAM': 'BALANCED',
                   'BAS': 'BALANCED SEARCH',
                   'UTL': 'UTILITY',
                   'WEB': 'WEB',
                   'XIO': 'XIO',
                   'XST': 'XSTORE',
                   'XDR': 'XDIRECT',
                   'VIZ': 'REMOTE',
                   'EWA': 'EOPWEB',
                   'ESD': 'EOPSTORAGE'}

    system = machine_name[8:11]

    for item in system_type:
        if system == item:
            return system_type.get(item)


def get_excel_name(gen_path: str, name_to_target: dict, machine_name: str):
    """
    Get the CRD file.
    :param machine_name:
    :param name_to_target:
    :param gen_path:
    :return:
    """
    target: str = name_to_target.get(machine_name)
    target_components: list = break_target_configuration(target)

    # Reduces possible_file
    reduce: list = []

    raw_file = []
    possible_file = []
    new_possible = []

    # Get Machine Name Processor
    reduce.append(decode_processor(machine_name))
    reduce.append(decode_system(machine_name))

    for crd_file in os.listdir(gen_path):
        raw_file.append(crd_file)
        for component in target_components:

            # Environment ie. Azure or Bing
            if 'Azure' in component or 'AZURE' in component or 'bing' in component:
                reduce.append('AZURE')
            elif 'Bing' in component or 'BING' in component or 'bing' in component:
                reduce.append('BING')

            # Processor ie. Intel, AMD, or ARM
            if 'Intel' in component or 'INTEL' in component or 'intel' in component:
                reduce.append('INTEL')
            elif 'Amd' in component or 'AMD' in component or 'amd' in component:
                reduce.append('AMD')
            elif 'Arm' in component or 'ARM' in component or 'arm' in component:
                reduce.append('ARM')

            if component in crd_file:
                possible_file.append(crd_file)

    reduce_components = list(set(reduce))

    # Removes items possible list for CRD
    initial = 0
    for file_tile in possible_file:
        for reduce_component in reduce_components:
            if reduce_component not in file_tile:
                new_possible.append(possible_file[initial])

    final_possible = list(set(new_possible))

    return final_possible


def warning_excel(possible_excel: list, machine_name: str, name_to_ticket: dict, crd_path: str, gen_folder_path: str):
    """
    Warns if anamalies in possible excel files for CRD
    :param gen_folder_path:
    :param crd_path:
    :param machine_name:
    :param name_to_ticket:
    :param possible_excel:
    :return:
    """
    location = 'Z:Drive'

    if len(possible_excel) > 1:
        print(f'   - {machine_name} -> TRR {Fore.RED}{name_to_ticket[machine_name]}{Style.RESET_ALL}  | '
              f'Multiple possible CRDs for {gen_from_name(machine_name)} in {location}')
        for possible_file in possible_excel:
            print(f'   - {possible_file}')
        return 'None'

    elif len(possible_excel) == 0:
        print(f'   - {machine_name} -> TRR {Fore.RED}{name_to_ticket[machine_name]}{Style.RESET_ALL} | '
              f'{Fore.RED}No Matching CRDs{Style.RESET_ALL} for {gen_from_name(machine_name)} in {location}')
        return 'None'

    else:
        return fr'{crd_path}\{gen_folder_path}{if_gen_6_path(machine_name)}\{possible_excel[0]}'


def if_gen_6_path(machine_name: str) -> str:
    """
    If it's a Gen 6 CRD, must decide whether Wiwynn or ZT folders based on current Z:Drive Structure - 9/25/2020
    :param machine_name:
    :return:
    """
    generation = machine_name[5]
    supplier = machine_name[7]

    if generation == '6' and supplier == 'Z':
        return r'\ZT'
    elif generation == '6' and supplier == 'W':
        return r'\Wiwynn'
    else:
        return ''


def process_name_to_target(name_to_target: dict) -> dict:
    """
    Check for anomalies in target.
    :param name_to_target:
    :return:
    """
    correct_number_of_characters: int = 15

    new_name_to_target: dict = {}

    # for machine_name in erase_name:
    #     print(f'name_to_target: {name_to_target[machine_name]}')
    #     name_to_target.pop(machine_name)

    for name in name_to_target:
        target = name_to_target[name]
        if len(target) == 0 or target is None:
            erase_target.append(name)
            pass

        elif '-VM-' in target:
            erase_target.append(name)
            pass

        elif 'CMA-' in target:
            erase_target.append(name)
            pass

        elif target.isspace():
            erase_target.append(name)
            pass

        elif len(target) != correct_number_of_characters:
            erase_target.append(name)
            pass

        else:
            new_name_to_target[name] = target

    return new_name_to_target


def process_name_to_ticket(name_to_ticket: dict) -> dict:
    """
    Check for anomalies in tickets.
    :param name_to_ticket:
    :return:
    """
    correct_number_of_characters_1 = 5
    correct_number_of_characters_2 = 6

    new_name_to_ticket: dict = {}

    for machine_name in name_to_ticket:
        ticket = name_to_ticket[machine_name]
        if len(ticket) == 0 or ticket is None:
            erase_name.append(machine_name)
            pass

        elif ticket.isspace():
            erase_name.append(machine_name)
            pass

        elif len(ticket) != correct_number_of_characters_1 and len(ticket) != correct_number_of_characters_2:
            erase_name.append(machine_name)
            pass

        elif not ticket.isdigit():
            erase_name.append(machine_name)
            pass

        elif check_valid_request(ticket) != 200:
            erase_name.append(machine_name)
            pass

        else:
            new_name_to_ticket[machine_name] = ticket

    return new_name_to_ticket


def main_method(name_to_target: dict, name_to_ticket: dict, full_name: str, description: str, break_length: str):
    """
    :param description:
    :param full_name:
    :param name_to_ticket: Machine Nmae to Ticket (TRR)
    :param name_to_target: Machine Name to Target Configuration
    :return:
    """
    name_to_crd: dict = {}

    crd_path = r'\\172.30.1.100\pxe\Kirkland_Lab\Microsoft_CSI\Documentation\CRD'

    print(f'\n  {break_length}  ')

    # Gets rid of empty values in dictionaries.
    new_name_to_target = {k: v for k, v in name_to_target.items() if v is not None}
    new_name_to_ticket = {k: v for k, v in name_to_ticket.items() if v}

    valid_name_to_ticket = process_name_to_ticket(new_name_to_ticket)
    valid_name_to_target = process_name_to_target(new_name_to_target)

    # for machine_name in erase_target:
    #     valid_name_to_target.pop(machine_name)

    # Scans for CRD in Pipe
    print(f'\n  CRD Summary - {full_name} | {description}:')
    for machine_name in valid_name_to_ticket:
        gen_folder_path = access_generation(crd_path, valid_name_to_target, machine_name)
        possible_excel = get_excel_name(gen_folder_path, valid_name_to_target, machine_name)
        single_crd_path = warning_excel(possible_excel, machine_name, valid_name_to_ticket, crd_path, gen_folder_path)
        name_to_crd[machine_name] = single_crd_path
