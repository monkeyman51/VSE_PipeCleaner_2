"""
Using Asyncio and AioHTTP to improve IO HTTP requests to Console Server and Azure Devops

Need to figure out the syntax of the Asyncio and AIOHTTP library, could improve by 5 to 6 times for performance.

There are two types of IO bound operations that we need to think about when dealing with Asynchronous programming.

One being the file system and other being the network, particularly using the requests module.

Typically speaking, we have 2 types of operations. Normally we write code that is synchronous. In this fashion, we
write code that waits until a sequence of executions is done before moving onto the next operations.  That's fine and
is very typical in most programming applications.

However, there will be a time when data needs to fetched with millions of requests that could take hours for a computer
to gather.  For this type of operation, we need to improve the performance and we could accomplish this with
non-blocking expressions that would help improve performance.
"""

import asyncio
import json

import aiohttp
from aiohttp import client_exceptions
from colorama import Fore, Style

from pipe_cleaner.src.data_access import write_host_groups_json, get_all_host_ids, get_dhcp

# from data_access import write_host_groups_json, get_all_host_ids, get_dhcp

# Store all tickets in a particular pipe. Clears after each new pipe
current_pipe_tickets: list = []
all_tickets: list = []
all_vm_data: dict = {}
total_systems_in_pipe: list = []
total_tickets_in_pipe: list = []
total_vms: list = []
total_systems: list = []

virtual_machine_data: list = []

# Stores the console server here, returns after collecting all data
console_server_data: dict = {}

# Stores Overall Workload of each VSE employee based on 'checked out to' in Console Server
overall_workload: dict = {}

# Commodity Serial Numbers
all_serial: list = []

# Commodity Part Numbers
all_part_numbers: list = []


async def generate_console_server_json(client: asyncio, host_id: str) -> str:
    """
    Generates the JSON data from the Host Details page using the Host ID.
    Returns product-serial if JSON data is generated properly.

    :param client:
    :param host_id: found on Host Details page in URL ?host_id=<some_id>get_system_json
    :return: returns product-serial string for getting JSON data
    """
    console_server_url = 'http://172.30.1.100/console/console_js.php'
    generate_data = {
        'action': 'get_json_data',
        'host_id': f'{host_id}'
    }
    host_name_data = {
        'action': 'get_host_name_data',
        'host_id': f'{host_id}'
    }
    async with client.post(url=console_server_url, json=generate_data) as response_generate:
        if response_generate.status == 200:
            async with client.post(url=console_server_url, json=host_name_data) as response_product_serial:
                assert response_product_serial.status == 200
                response_text = await response_product_serial.text()

                try:
                    product = json.loads(str(response_text))['host_name_data']['product']
                    serial = json.loads(str(response_text))['host_name_data']['serial']
                    return f'{product}-{serial}'
                except json.decoder.JSONDecodeError:
                    pass


async def get_system_json(client, product_serial: str, host_id: str) -> dict:
    """
    Gets the Generated data using the product_serial string and creates JSON file.
    Returns JSON of Host within Console Server.

    :param client:
    :param product_serial: string from generate_json_data method
    :param host_id: 24-character string host id in URL of Host Details in Console Server
    :return: JSON data of Host
    """
    data = {
        'action': 'get_json_data',
        'host_id': f'{host_id}'
    }
    async with client.post(url=f'http://172.30.1.100/results/{product_serial}.json', json=data) as response_data:
        # assert response_data.status == 200

        data = await response_data.text()
        try:
            return json.loads(data)
        except json.decoder.JSONDecodeError:
            pass


def calculate_last_time_alive(raw_last_found_alive: str) -> float:
    """
    Calculate last time alive in non-seconds. Raw last time alive is in seconds and ain't no body got time for that.
    :param raw_last_found_alive:
    :return:
    """
    import time
    return float(time.time()) - float(raw_last_found_alive)


def check_none_type(component):
    """
    Checks if NoneType exception occurred
    :param component:
    :return:
    """
    if component is None:
        return 'None'
    else:
        return component


def get_virtual_machine_data(system_json: dict, pipe_name: str) -> None:
    """
    Get virtual machine information for engineers to lower confusion on VM usage
    :param pipe_name: host group name
    :param system_json: system data based from Console Server
    :return: appends to list for later gathering in the Console Server data
    """
    # check_none_type function accounts for rare case of NoneType errors
    vm_data: dict = {'pipe_name': check_none_type(pipe_name),
                     'machine_name': check_none_type(system_json.get('machine_name')),
                     'host_id': check_none_type(system_json.get('id')),
                     'connection_status': check_none_type(system_json.get('connection_status')),
                     'last_found_alive': check_none_type(system_json.get('last_found_alive')),
                     'location': check_none_type(system_json.get('location')),
                     'comment': check_none_type(system_json.get('comment')),
                     'ssh_connection_string': check_none_type(system_json.get('ssh_connection_string')),
                     'rdp_connection_string': check_none_type(system_json.get('rdp_connection_string')),
                     'vnc_connection_string': check_none_type(system_json.get('vnc_connection_string')),
                     'host_ip': check_none_type(system_json.get('net', {}).get('interfaces', {})[0].get('ip')),
                     'checked_out_to': check_none_type(system_json.get('checked_out_to'))}

    virtual_machine_data.append(vm_data)


def check_username_is_empty(username: str) -> bool:
    """

    :param username: Username with underscore separator
    :return: True or False
    """
    if not username or username.upper() == 'NONE':
        return False
    else:
        return True


def check_username_in_userbase(user_name: str) -> bool:
    """
    Ensures the username is already in the dictionary

    :param user_name: Individual
    :return: True or False
    """
    user_base: dict = console_server_data['user_base']

    if check_username_is_empty(user_name) is True and user_name in user_base:
        return True
    else:
        return False


def count_username(clean_username: str) -> None:
    """

    :param clean_username:
    """
    check_user_base: bool = check_username_in_userbase(clean_username)
    is_username_empty: bool = check_username_is_empty(clean_username)

    if is_username_empty is False:
        pass

    elif check_user_base is False:
        console_server_data['user_base'][clean_username]: dict = {}
        console_server_data['user_base'][clean_username]['count']: int = 1
        console_server_data['user_base'][clean_username]['default_name']: str = clean_username.replace('_', '.')
        console_server_data['user_base'][clean_username]['alt_name']: str = clean_username.replace('_', ' '). \
            title().replace(' ', '')
        console_server_data['user_base'][clean_username]['systems']: dict = {}
        console_server_data['user_base'][clean_username]['virtual_machines']: dict = {}

    elif check_user_base is True:
        console_server_data['user_base'][clean_username]['count'] += 1


def store_machine_name(user_name: str, machine_name: str) -> None:
    """

    :param machine_name:
    :param user_name:
    """
    is_username_empty: bool = check_username_is_empty(user_name)

    if is_username_empty is False:
        pass

    elif machine_name not in console_server_data['user_base'][user_name]['systems']:
        if '-VM-' in machine_name:
            console_server_data['user_base'][user_name]['virtual_machines'][machine_name]: dict = {}

        else:
            console_server_data['user_base'][user_name]['systems'][machine_name]: dict = {}


def store_pipe_name(user_name: str, machine_name: str, pipe_name: str):
    """

    :param user_name:
    :param machine_name:
    :param pipe_name:
    :return:
    """
    is_username_empty: bool = check_username_is_empty(user_name)

    if is_username_empty is False:
        pass

    else:
        if '-VM-' in machine_name:
            console_server_data['user_base'][user_name]['virtual_machines'][machine_name]['pipe_name'] = pipe_name

        else:
            console_server_data['user_base'][user_name]['systems'][machine_name]['pipe_name'] = pipe_name


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

    return clean_data.replace('Pipe-', '').replace(last_part, '').strip()


def store_system_data(system_json: dict, pipe_name: str):
    """
    Grab relevant information based on needs. Stores into Console Server data structure
    :param pipe_name:
    :param system_json:
    """
    # Accounts for a system
    total_systems.append(1)
    # Temporary Data Structure to hold System Data
    try:
        current_system_data: dict = {'machine_name': str(system_json.
                                                         get('machine_name', 'None'))}
        current_machine_name: str = system_json.get('machine_name')

        if 'VSE' in current_machine_name and '-VM-' in current_machine_name:
            get_virtual_machine_data(system_json, pipe_name)

        if 'VSE' in current_machine_name and '-VM-' not in current_machine_name:
            setup_data: list = console_server_data[pipe_name]['setup_data']['total_systems']
            setup_data.append(1)

        current_system_data['ticket']: dict = str(system_json.get('ticket', 'None'))

        # Stores all tickets for unique tickets data and pipe-specific tickets
        ticket = current_system_data.get('ticket', 'None')
        if ticket != '' and ticket.isdigit() is True and '-VM-' not in current_system_data.get('machine_name'):
            all_tickets.append(ticket)
            current_pipe_tickets.append(ticket)
            setup_data: list = console_server_data[pipe_name]['setup_data']['systems_with_ticket']
            setup_data.append(1)

        # Username
        checked_out_to = str(system_json.get('checked_out_to'))
        current_system_data['checked_out_to']: dict = checked_out_to

        # Upper for consistent comparison
        upper_machine_name = str(current_system_data.get('machine_name', 'None'))

        connection_status = str(system_json.get('connection_status', 'None'))

        # Stores checked_out_to for pipe-level work
        clean_username: str = checked_out_to.lower().replace('.', '_')

        if clean_username not in overall_workload:
            overall_workload[clean_username] = {}

        if 'systems' not in overall_workload[clean_username]:
            overall_workload[clean_username]['systems'] = 0

        if 'systems' in overall_workload[clean_username]:
            overall_workload[clean_username]['systems'] += 1

        # Last Time Alive
        raw_last_found_alive = str(system_json.get('last_found_alive', 'None'))
        last_alive: float = calculate_last_time_alive(raw_last_found_alive)
        current_system_data['last_found_alive']: float = last_alive

        # Machine Serial Number
        machine_serial_number: str = system_json.get('bmc', {}).get('fru', {}).get('product', {}).get('serial')

        machine_product_name: str = system_json.get('bmc', {}).get('fru', {}).get('product', {}).get('name')

        # Comment / Status
        current_system_data['comment']: dict = str(system_json.
                                                   get('comment', 'None'))

        # Status ex. Idle, Offline
        current_system_data['system_status']: dict = str(system_json.
                                                         get('status', 'None'))

        # Connection Status
        current_system_data['connection_status']: dict = str(system_json.
                                                             get('connection_status', 'None'))

        # Location ex. R44U14N14
        current_system_data['location']: dict = str(system_json.
                                                    get('location', 'None'))

        host_ip: str = str(system_json['net']['interfaces'][0].get('ip', 'None'))
        current_system_data['host_ip']: dict = host_ip

        # BIOS ex. C2030.BS.2A42.AF1
        current_system_data['server_bios']: dict = str(system_json.
                                                       get('dmi', {}).
                                                       get('bios', {}).
                                                       get('version', 'None'))

        # Processor ex. Intel or AMD
        current_system_data['processors']: dict = str(system_json.
                                                      get('dmi', {}).
                                                      get('procs', 'None'))

        # BMC ex. 3.30
        current_system_data['server_bmc']: dict = str(system_json.
                                                      get('bmc', {}).
                                                      get('mc', {}).
                                                      get('firmware', 'None'))

        # CPLD Version ex. 000000014
        current_system_data['server_cpld']: dict = str(system_json.
                                                       get('cpld', {}).
                                                       get('secure_cpld_version', 'None'))

        # CPLD Sequence Version ex. 00000014
        current_system_data['sequence_cpld_version']: dict = str(system_json.
                                                                 get('cpld', {}).
                                                                 get('sequence_cpld_version', 'None'))

        # Platform Version
        current_system_data['server_os']: dict = str(system_json.
                                                     get('platform', {}).
                                                     get('version', 'None'))

        # TPM
        current_system_data['server_tpm']: dict = str(system_json.
                                                      get('tpm', {}).
                                                      get('version', 'None'))

        # Host ID ex. 5e5460a20b4b023d327dda51
        current_system_data['id']: dict = str(system_json.
                                              get('id', 'None'))

        # DIMMs ex.
        current_system_data['system_dimms']: dict = str(system_json.
                                                        get('dmi', {}).
                                                        get('dimms', 'None'))

        # Unique DIMMs
        unique_dimms: dict = system_json.get('dmi', {}).get('unique_dimms', 'None')
        for unique_dimm in unique_dimms:
            part_number: list = unique_dimm.get('part')
            count = int(unique_dimm.get('count'))

            clean_pipe_name: str = process_pipe_name(pipe_name)
            part_numbers: dict = console_server_data['part_numbers']

            if part_number in part_numbers:
                part_numbers[part_number]['quantity'] += count
                locations: dict = part_numbers[part_number]['locations']

                if clean_pipe_name in locations:
                    locations[clean_pipe_name]['count'] += count

                else:
                    locations[clean_pipe_name]: dict = {}
                    locations[clean_pipe_name]['count'] = count
                    locations[clean_pipe_name]['connection']: int = connection_status
                    locations[clean_pipe_name]['last_alive']: float = last_alive

            else:
                part_numbers[part_number]: dict = {}
                part_numbers[part_number]['locations']: dict = {}
                part_numbers[part_number]['locations'][clean_pipe_name]: dict = {}

                part_numbers[part_number]['quantity']: int = count
                part_numbers[part_number]['locations'][clean_pipe_name]['count']: int = count
                part_numbers[part_number]['locations'][clean_pipe_name]['connection']: int = connection_status
                part_numbers[part_number]['locations'][clean_pipe_name]['last_alive']: float = last_alive

            if not part_number:
                pass

            else:
                if part_number not in console_server_data['inventory']['commodities']['dimms']:
                    console_server_data['inventory']['commodities']['dimms'][part_number] = count
                else:
                    console_server_data['inventory']['commodities']['dimms'][part_number] += count

                if part_number not in console_server_data[pipe_name]['pipe_data']['pipe_inventory']['dimms']:
                    console_server_data[pipe_name]['pipe_data']['pipe_inventory']['dimms'][part_number] = count
                else:
                    console_server_data[pipe_name]['pipe_data']['pipe_inventory']['dimms'][part_number] += count

                try:
                    current_commodity: dict = {'pipe_name': pipe_name,
                                               'connection_status': connection_status,
                                               'last_found_alive': current_system_data['last_found_alive'],
                                               'machine_unique_id': f'{machine_product_name}-'
                                                                    f'{machine_serial_number}',
                                               'machine_sku': machine_product_name,
                                               'machine_serial': machine_serial_number,
                                               'supplier': unique_dimm['manufacturer'],
                                               'machine_name': upper_machine_name,
                                               'commodity_type': 'NVMe',
                                               'part_number': part_number,
                                               'count': count,
                                               'type': 'DIMM'}

                    all_part_numbers.append(current_commodity)
                except KeyError:
                    pass

        all_dimms: dict = system_json.get('dmi', {}).get('dimms')
        if not all_dimms:
            pass

        else:
            for current_dimm in all_dimms:
                if not current_dimm:
                    pass
                else:
                    try:
                        current_commodity: dict = {'pipe_name': pipe_name,
                                                   'connection_status': connection_status,
                                                   'machine_unique_id': f'{machine_product_name}-'
                                                                        f'{machine_serial_number}',
                                                   'machine_sku': machine_product_name,
                                                   'machine_serial': machine_serial_number,
                                                   'machine_name': upper_machine_name,
                                                   'commodity_type': 'DIMM',
                                                   'part_number': current_dimm['part'],
                                                   'serial_number': current_dimm['serial']}
                        all_serial.append(current_commodity)
                    except KeyError:
                        pass

        current_system_data['unique_dimms']: dict = unique_dimms

        # NVMes
        current_system_data['system_nvmes']: dict = str(system_json.
                                                        get('nvme', {}).
                                                        get('nvmes', 'None'))

        # Unique NVMes
        unique_nvmes: dict = system_json.get('nvme', {}).get('unique_nvmes')
        for unique_nvme in unique_nvmes:
            nvme_model = get_nvme_model(unique_nvme)
            count = int(unique_nvme.get('count'))

            clean_pipe_name: str = process_pipe_name(pipe_name)
            part_numbers: dict = console_server_data['part_numbers']

            if nvme_model in part_numbers:
                part_numbers[nvme_model]['quantity'] += count
                locations: dict = part_numbers[nvme_model]['locations']

                if clean_pipe_name in locations:
                    locations[clean_pipe_name]['count'] += count

                else:
                    locations[clean_pipe_name]: dict = {}
                    locations[clean_pipe_name]['count'] = count
                    locations[clean_pipe_name]['connection']: int = connection_status
                    locations[clean_pipe_name]['last_alive']: float = last_alive

            else:
                part_numbers[nvme_model]: dict = {}
                part_numbers[nvme_model]['locations']: dict = {}
                part_numbers[nvme_model]['locations'][clean_pipe_name]: dict = {}

                part_numbers[nvme_model]['quantity']: int = count
                part_numbers[nvme_model]['locations'][clean_pipe_name]['count']: int = count
                part_numbers[nvme_model]['locations'][clean_pipe_name]['connection']: int = connection_status
                part_numbers[nvme_model]['locations'][clean_pipe_name]['last_alive']: float = last_alive

            if not nvme_model:
                pass
            else:
                if nvme_model not in console_server_data['inventory']['commodities']['nvmes']:
                    console_server_data['inventory']['commodities']['nvmes'][nvme_model] = count
                else:
                    console_server_data['inventory']['commodities']['nvmes'][nvme_model] += count

                if nvme_model not in console_server_data[pipe_name]['pipe_data']['pipe_inventory']['nvmes']:
                    console_server_data[pipe_name]['pipe_data']['pipe_inventory']['nvmes'][nvme_model] = count
                else:
                    console_server_data[pipe_name]['pipe_data']['pipe_inventory']['nvmes'][nvme_model] += count

                try:
                    current_commodity: dict = {'pipe_name': pipe_name,
                                               'connection_status': connection_status,
                                               'last_found_alive': current_system_data['last_found_alive'],
                                               'machine_unique_id': f'{machine_product_name}-'
                                                                    f'{machine_serial_number}',
                                               'machine_sku': machine_product_name,
                                               'machine_serial': machine_serial_number,
                                               'machine_name': upper_machine_name,
                                               'commodity_type': 'NVMe',
                                               'part_number': nvme_model,
                                               'count': count,
                                               'type': 'NVMe'}

                    all_part_numbers.append(current_commodity)
                except KeyError:
                    pass

        current_system_data['unique_nvmes']: dict = unique_nvmes

        all_nvmes: dict = system_json.get('nvme', {}).get('nvmes')
        if not all_nvmes:
            pass

        else:
            for current_nvme in all_nvmes:

                if not current_nvme:
                    pass

                else:
                    try:
                        current_commodity: dict = {'pipe_name': pipe_name,
                                                   'connection_status': connection_status,
                                                   'machine_unique_id': f'{machine_product_name}-'
                                                                        f'{machine_serial_number}',
                                                   'machine_sku': machine_product_name,
                                                   'machine_serial': machine_serial_number,
                                                   'machine_name': upper_machine_name,
                                                   'commodity_type': 'NVMe',
                                                   'part_number': current_nvme['model'],
                                                   'serial_number': current_nvme['serial']}

                        all_serial.append(current_commodity)
                    except KeyError:
                        pass

        # Disks
        current_system_data['system_disks']: dict = str(system_json.
                                                        get('disk', {}).
                                                        get('disks', 'None'))

        # Unique Disks
        unique_disks: dict = system_json.get('disk', {}).get('unique_disks', 'None')

        for unique_disk in unique_disks:
            disk_part_number: str = get_disk_part_number(unique_disk)
            count = int(unique_disk.get('count'))

            clean_pipe_name: str = process_pipe_name(pipe_name)
            part_numbers: dict = console_server_data['part_numbers']

            if disk_part_number in part_numbers:
                part_numbers[disk_part_number]['quantity'] += count
                locations: dict = part_numbers[disk_part_number]['locations']

                if clean_pipe_name in locations:
                    locations[clean_pipe_name]['count'] += count

                else:
                    locations[clean_pipe_name]: dict = {}
                    locations[clean_pipe_name]['count'] = count
                    locations[clean_pipe_name]['connection']: int = connection_status
                    locations[clean_pipe_name]['last_alive']: float = last_alive

            else:
                part_numbers[disk_part_number]: dict = {}
                part_numbers[disk_part_number]['locations']: dict = {}
                part_numbers[disk_part_number]['locations'][clean_pipe_name]: dict = {}

                part_numbers[disk_part_number]['quantity']: int = count
                part_numbers[disk_part_number]['locations'][clean_pipe_name]['count']: int = count
                part_numbers[disk_part_number]['locations'][clean_pipe_name]['connection']: int = connection_status
                part_numbers[disk_part_number]['locations'][clean_pipe_name]['last_alive']: float = last_alive

            if not disk_part_number or disk_part_number == 'Virtual HD':
                pass

            else:
                if disk_part_number not in console_server_data['inventory']['commodities']['disks']:
                    console_server_data['inventory']['commodities']['disks'][disk_part_number] = count
                else:
                    console_server_data['inventory']['commodities']['disks'][disk_part_number] += count

                if disk_part_number not in console_server_data[pipe_name]['pipe_data']['pipe_inventory']['disks']:
                    console_server_data[pipe_name]['pipe_data']['pipe_inventory']['disks'][disk_part_number] = count
                else:
                    console_server_data[pipe_name]['pipe_data']['pipe_inventory']['disks'][disk_part_number] += count

                try:
                    current_commodity: dict = {'pipe_name': pipe_name,
                                               'connection_status': connection_status,
                                               'last_found_alive': current_system_data['last_found_alive'],
                                               'machine_unique_id': f'{machine_product_name}-'
                                                                    f'{machine_serial_number}',
                                               'machine_sku': machine_product_name,
                                               'machine_serial': machine_serial_number,
                                               'machine_name': upper_machine_name,
                                               'commodity_type': 'NVMe',
                                               'part_number': disk_part_number,
                                               'count': count,
                                               'type': 'Disk'}

                    all_part_numbers.append(current_commodity)
                except KeyError:
                    pass

        current_system_data['unique_disks']: dict = unique_disks

        all_disks: dict = system_json.get('disk', {}).get('disks')
        if not all_disks:
            pass

        else:
            for current_disk in all_disks:
                if not current_disk:
                    pass

                else:
                    try:
                        current_commodity: dict = {'pipe_name': pipe_name,
                                                   'connection_status': connection_status,
                                                   'machine_unique_id': f'{machine_product_name}-'
                                                                        f'{machine_serial_number}',
                                                   'machine_sku': machine_product_name,
                                                   'machine_serial': machine_serial_number,
                                                   'machine_name': upper_machine_name,
                                                   'commodity_type': 'Disk',
                                                   'part_number': current_disk['model'],
                                                   'serial_number': current_disk['serial']}
                        all_serial.append(current_commodity)

                    except KeyError:
                        pass

        count_username(clean_username)
        store_machine_name(clean_username, upper_machine_name)
        store_pipe_name(clean_username, upper_machine_name, pipe_name)

        # Avoid other
        if 'VSE' not in upper_machine_name and '-VM-' not in upper_machine_name:
            pass

        # Tallies total VMs and stores VM data
        elif '-VM-' in upper_machine_name:
            all_vm_data[upper_machine_name]: dict = {}
            all_vm_data[upper_machine_name]['name'] = upper_machine_name
            all_vm_data[upper_machine_name]['username']: dict = str(system_json.get('username', 'None'))
            all_vm_data[upper_machine_name]['host_ip']: dict = str(system_json.get('id', 'None'))
            total_vms.append(1)

        # Actually now stores data inside pipe information, CMA accounts for Chassis Manager Assembly, usually GEN 5
        elif 'VSE' in upper_machine_name and 'CMA-' not in upper_machine_name and '-VM-' not in upper_machine_name:

            # Accounts for setup information later in Setup Dashboard
            if not current_system_data.get('ticket') and str(current_system_data.get('ticket')).isdigit():
                total_tickets_in_pipe.append(1)
            total_systems_in_pipe.append(1)
            console_server_data[pipe_name]['pipe_data'][upper_machine_name]: dict = current_system_data

    except AttributeError:
        pass


def get_disk_part_number(unique_disk):
    split_parts: list = unique_disk.get('model').replace('  ', ' ').replace('_', ' ').split(' ')

    if len(split_parts) >= 3:
        for item in split_parts:
            if 'Micron' in item:
                return split_parts[-1].strip()
        else:
            return unique_disk.get("model").strip()

    count: int = 0
    for part_split in split_parts:
        if part_split.isalpha():
            count += 1

    if count == len(split_parts):
        return unique_disk.get('model').strip()

    else:
        for part_split in split_parts:
            if not part_split.isalpha():
                if not part_split:
                    return part_split
                else:
                    return part_split
        else:
            return unique_disk.get('model')


def get_nvme_model(unique_nvme):
    for part_split in unique_nvme.get('model').split(' '):
        if not part_split.isalpha():
            if not part_split:
                return part_split
            else:
                return part_split
    else:
        return ''


async def generate_json(client: asyncio, host_id: str, pipe_name: str):
    """
    Get JSON data from Console Server then stores data from
    :param client: async Session
    :param host_id:
    :param pipe_name:
    :return:
    """
    product_serial: str = await generate_console_server_json(client, host_id)
    try:
        upper_product_serial: str = product_serial.upper()

        if 'NONE' not in upper_product_serial:
            system_json: dict = await get_system_json(client, product_serial, host_id)

            store_system_data(system_json, pipe_name)

    except AttributeError:
        pass


async def start_loop(host_ids: list, pipe_name: str):
    """
    Generate Data, fetch data, and process data into a dictionary to analyze later
    :return:
    """
    event_loop = asyncio.get_running_loop()
    async with aiohttp.ClientSession(loop=event_loop) as client:
        tasks = [generate_json(client, host_id, pipe_name) for host_id in host_ids]
        await asyncio.gather(*tasks)


def get_group_unique_tickets(host_id: str, pipe_name: str) -> list:
    """
    Get the Host Group unique tickets for later extraction
    Note: THis is for all host
    :param pipe_name:
    :param host_id: Host Group ID
    :return:
    """
    # Store data
    host_group_unique_tickets: list = []

    # Single pipe data in JSON
    host_group_json: list = get_all_host_ids(host_id, pipe_name)

    for system in host_group_json:
        ticket: str = str(system['ticket'])

        # TRR IDs should only be numbers only
        if ticket.isdigit() is True:
            host_group_unique_tickets.append(ticket)

    # Insures only unique tickets
    return list(set(host_group_unique_tickets))


def get_overall_workload():
    return overall_workload


def get_total_systems():
    return sum(total_systems)


def main_method() -> dict:
    """
    Fetches all relevant information from Console Server using Async approach, accounts for Setup dashboard as well
    :return: returns dictionary showing for all pipe information first, note the first layer of the dictionary structure
    includes the unique tickets and will shoot an error message if not accounted for with 'Pipe-'
    """
    console_server_data['user_base']: dict = {}

    # Gets data and stores in JSON file for later use
    # Don't need async for this considering it's just one request
    json_file: dict = write_host_groups_json()
    host_groups_page: dict = json_file['host_groups']

    # Logs skipped or collected data based on naming of Host Groups
    log_total: list = []
    log_skipped: list = []
    log_collect: list = []
    log_non_pipes: list = []
    log_off_lines: list = []
    log_idle: list = []

    console_server_data['inventory']: dict = {}
    console_server_data['inventory']['commodities']: dict = {}
    console_server_data['inventory']['commodities']['dimms']: dict = {}
    console_server_data['inventory']['commodities']['nvmes']: dict = {}
    console_server_data['inventory']['commodities']['disks']: dict = {}

    console_server_data['part_numbers']: dict = {}

    # Tells user progress in Terminal environment
    print(f'\n\t=====================================================================')
    print(f'\t  Console Server - Collecting and Processing Data:')
    print(f'\t=====================================================================')
    print(f'\t\t  STATUS   |  REASON    |  HOST GROUP NAME')

    # Note that Host Groups is synonymous with Pipes
    for host_group in host_groups_page:

        # Gather basic Pipe information as Host Group, converts to upper and string to ensure proper comparison
        host_group_name: str = host_group.get('name')
        description: str = host_group.get('description')
        host_id: str = str(host_group.get('id')).upper()
        host_ids: list = host_group.get('host_ids')
        status: str = str(host_group.get('comment')).upper()
        checked_out_to: str = str(host_group.get('checked_out_to')).upper()

        # Stores checked_out_to for pipe-level work
        clean_username = checked_out_to.lower().replace('.', '_')
        if clean_username not in overall_workload:
            overall_workload[clean_username] = {}

        if 'pipes' not in overall_workload[clean_username]:
            overall_workload[clean_username]['pipes']: int = 0

        if 'pipes' in overall_workload[clean_username]:
            overall_workload[clean_username]['pipes'] += 1

        # Create pipe data structure within all the Console Server data
        console_server_data[host_group_name]: dict = {}

        # Host Groups that are actually getting data from
        # if 'Pipe-' in host_group_name and 'OFFLINE' not in host_group_name and 'OFFLINE' not in status:

        # Tells users progress
        print(f'\t\t- Collect  |  {Fore.GREEN}Success{Style.RESET_ALL}   |  {host_group_name}')
        log_total.append(1)
        log_collect.append(1)

        # Get all system data from Pipe
        current_pipe: dict = console_server_data[host_group_name]
        console_server_data[host_group_name]['pipe_data']: dict = {}

        # Stores
        current_pipe['pipe_data']: dict = {}

        current_pipe['pipe_data']['pipe_inventory']: dict = {}
        current_pipe['pipe_data']['pipe_inventory']['dimms']: dict = {}
        current_pipe['pipe_data']['pipe_inventory']['nvmes']: dict = {}
        current_pipe['pipe_data']['pipe_inventory']['disks']: dict = {}

        # Stores total systems in Pipe (non-vm or anything else)
        current_pipe['setup_data']: dict = {}
        current_pipe['setup_data']['total_systems']: dict = []
        current_pipe['setup_data']['systems_with_ticket']: dict = []

        # Runs async and stores data in module-level dictionary
        try:
            asyncio.run(start_loop(host_ids, host_group_name))
        except aiohttp.client_exceptions.ClientPayloadError:
            return main_method()

        # Store all system data from Pipe
        current_pipe['description']: str = description
        current_pipe['host_group_status']: str = status
        current_pipe['host_id']: str = host_id
        current_pipe['host_ids']: list = host_ids
        current_pipe['checked_out_to']: str = checked_out_to

        # Get and store ticket information
        group_unique_tickets: list = list(set(current_pipe_tickets))
        current_pipe['group_unique_tickets']: list = group_unique_tickets
        current_pipe_tickets.clear()

        # Store information for setup dashboard later
        current_pipe['ticket_tally']: dict = {}
        current_pipe['ticket_tally']['systems_with_tickets']: dict = sum(total_tickets_in_pipe)
        current_pipe['ticket_tally']['total_systems']: dict = sum(total_systems_in_pipe)
        current_pipe['ticket_tally']['total_vms']: dict = sum(total_vms)

        # Clear information on pipe tally for next iteration
        total_vms.clear()
        total_tickets_in_pipe.clear()
        total_systems_in_pipe.clear()

    # Stores host groups level information for Setup information later
    console_server_data['host_groups_data']: dict = {}
    console_server_data['host_groups_data']['all_unique_tickets']: dict = list(set(all_tickets))
    console_server_data['host_groups_data']['log_total']: str = str(sum(log_total))
    console_server_data['host_groups_data']['log_skipped']: str = str(sum(log_skipped))
    console_server_data['host_groups_data']['log_collect']: str = str(sum(log_collect))
    console_server_data['host_groups_data']['log_non_pipes']: str = str(sum(log_non_pipes))
    console_server_data['host_groups_data']['log_off_lines']: str = str(sum(log_off_lines))
    console_server_data['host_groups_data']['log_idle']: str = str(sum(log_idle))

    console_server_data['dhcp_data']: list = get_dhcp()

    # Sorted to keep consistent alphabetic order
    console_server_data['virtual_machine_data']: list = virtual_machine_data

    console_server_data['all_serial']: list = all_serial
    console_server_data['all_part_numbers']: list = all_part_numbers

    return console_server_data
