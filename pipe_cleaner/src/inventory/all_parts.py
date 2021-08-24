"""
First attempt to marry all data from Console Server and Inventory Account.

7/26/2021
"""
from csv import reader
import requests
from json import loads, dumps, decoder
import asyncio
from aiohttp import ClientSession, client_exceptions

hosts_data: list = []

# Store all tickets in a particular pipe. Clears after each new pipe
current_pipe_tickets: list = []
total_systems_in_pipe: list = []
total_tickets_in_pipe: list = []

# Stores the console server here, returns after collecting all data
console_server_data: dict = {"All": {}}


def store_inventory_count(csv_row: dict, total_count: dict) -> dict:
    """
    Store inventory count based off of the Inventory Team counting as of 7/22 count.
    """
    specific_location: str = csv_row[0]
    part_number: str = csv_row[1]
    count: str = csv_row[2]

    if not part_number or not count:
        return total_count

    elif part_number == "EMPTY" or part_number == "0":
        return total_count

    else:
        if part_number in total_count:
            total_count[part_number]["location"] = "Cage"
            total_count[part_number]["specific_location"] = specific_location
            total_count[part_number]["count"] += int(count)
            return total_count

        else:
            total_count[part_number]: dict = {}
            total_count[part_number]["location"] = "Cage"
            total_count[part_number]["specific_location"] = specific_location
            total_count[part_number]["count"] = int(count)
            return total_count


def get_inventory_count() -> dict:
    """
    Get inventory data from 7/22 count.
    """
    with open("settings/inventory_count.csv") as file:
        csv_data = reader(file, delimiter=",", quotechar='"')

        total_count: dict = {}
        for index, csv_row in enumerate(csv_data, start=0):
            if index == 0:
                pass

            else:
                total_count: dict = store_inventory_count(csv_row, total_count)

        return total_count


def get_all_host_ids() -> None:
    """
    Gather all host data from Console Server's All Host page. Purpose is to gather individual host ID and IP for later
    async fetching.
    """
    host_ids: list = get_host_ids_from_console_server()

    try:
        asyncio.run(start_loop(host_ids))

    except client_exceptions.ClientPayloadError:
        get_all_host_ids()


def get_host_ids_from_console_server():
    """
    Extract host IDs from Console Server
    """
    all_hosts: list = get_console_server_all_hosts()
    host_ids: list = []
    for host in all_hosts:
        machine_name: str = host.get('machine_name', 'None').upper()

        if '-VM-' not in machine_name \
                and machine_name.upper() != "NONE" \
                and machine_name.upper() != "":
            host_id: str = host.get('id', 'None')

            if host_id:
                host_ids.append(host_id)
    return host_ids


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
                    product = loads(str(response_text))['host_name_data']['product']
                    serial = loads(str(response_text))['host_name_data']['serial']
                    return f'{product}-{serial}'
                except decoder.JSONDecodeError:
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
        data = await response_data.text()
        try:
            return loads(data)
        except decoder.JSONDecodeError:
            pass


async def generate_json(client: asyncio, host_id: str):
    """
    Get JSON data from Console Server then stores data from
    """
    product_serial: str = await generate_console_server_json(client, host_id)
    try:
        upper_product_serial: str = product_serial.upper()

        if 'NONE' not in upper_product_serial:
            system_json: dict = await get_system_json(client, product_serial, host_id)
            hosts_data.append(system_json)

            # store_system_data(system_json)

    except AttributeError:
        pass


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


def calculate_last_time_alive(raw_last_found_alive: str) -> float:
    """
    Calculate last time alive in non-seconds. Raw last time alive is in seconds and ain't no body got time for that.
    :param raw_last_found_alive:
    :return:
    """
    import time
    return float(time.time()) - float(raw_last_found_alive)


def get_nvme_model(unique_nvme):
    for part_split in unique_nvme.get('model').split(' '):
        if not part_split.isalpha():
            if not part_split:
                return part_split
            else:
                return part_split
    else:
        return ''


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


def check_username_is_empty(username: str) -> bool:
    """

    :param username: Username with underscore separator
    :return: True or False
    """
    if not username or username.upper() == 'NONE':
        return False
    else:
        return True


def store_machine_name(user_name: str, machine_name: str) -> None:
    """
    Pass
    """
    is_username_empty: bool = check_username_is_empty(user_name)

    if is_username_empty is False:
        pass

    elif machine_name not in console_server_data['user_base'][user_name]['systems']:
        if '-VM-' in machine_name:
            console_server_data['user_base'][user_name]['virtual_machines'][machine_name]: dict = {}

        else:
            console_server_data['user_base'][user_name]['systems'][machine_name]: dict = {}


def store_system_data(system_json: dict):
    """
    Grab relevant information based on needs. Stores into Console Server data structure
    """
    try:
        current_system_data: dict = {'machine_name': str(system_json.get('machine_name', 'None'))}

        upper_machine_name = str(current_system_data.get('machine_name', 'None'))
        connection_status = str(system_json.get('connection_status', 'None'))

        # Last Time Alive
        raw_last_found_alive = str(system_json.get('last_found_alive', 'None'))
        last_alive: float = calculate_last_time_alive(raw_last_found_alive)
        current_system_data['last_found_alive']: float = last_alive

        # Status ex. Idle, Offline
        current_system_data['system_status']: dict = str(system_json.
                                                         get('status', 'None'))

        # Connection Status
        current_system_data['connection_status']: dict = str(system_json.
                                                             get('connection_status', 'None'))

        # Location ex. R44U14N14
        current_system_data['location']: dict = str(system_json.
                                                    get('location', 'None'))

        # Unique DIMMs
        unique_dimms: dict = system_json.get('dmi', {}).get('unique_dimms', 'None')
        for unique_dimm in unique_dimms:
            part_number: list = unique_dimm.get('part')
            count = int(unique_dimm.get('count'))

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
                                                   'machine_name': upper_machine_name,
                                                   'commodity_type': 'NVMe',
                                                   'part_number': current_nvme['model'],
                                                   'serial_number': current_nvme['serial']}

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

        current_system_data['unique_disks']: dict = unique_disks

    except AttributeError:
        pass


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


def get_console_server_all_hosts() -> list:
    """
    Get all hosts found in the All Hosts section of ZT Console Server.
    """
    generate_data: dict = {
        'action': 'get_host_status'
    }
    response = requests.post(url='http://172.30.1.100/console/console_js.php', data=dumps(generate_data))
    return loads(response.text)


async def start_loop(host_ids: list):
    """
    Generate Data, fetch data, and process data into a dictionary to analyze later
    :return:
    """
    event_loop = asyncio.get_running_loop()
    async with ClientSession(loop=event_loop) as client:
        tasks = [generate_json(client, host_id) for host_id in host_ids]
        await asyncio.gather(*tasks)


def get_inventory_from_machine(machine_json: dict, inventory_total: dict) -> dict:
    """
    Get inventory related data from machine JSON.
    """
    try:
        unique_disks: dict = machine_json.get('disk', {}).get('unique_disks', 'None')
        for unique_disk in unique_disks:
            disk_part_number: str = get_disk_part_number(unique_disk)
            count = int(unique_disk.get('count'))

            if disk_part_number not in inventory_total:
                inventory_total[disk_part_number]: int = count
            else:
                inventory_total[disk_part_number]: int = + count
    except AttributeError:
        pass

    try:
        unique_nvmes: dict = machine_json.get('nvme', {}).get('unique_nvmes')
        for unique_nvme in unique_nvmes:
            nvme_model: str = get_nvme_model(unique_nvme)
            count = int(unique_nvme.get('count'))

            if nvme_model not in inventory_total:
                inventory_total[nvme_model]: int = count
            else:
                inventory_total[nvme_model]: int = + count
    except AttributeError:
        pass

    try:
        unique_dimms: dict = machine_json.get('dmi', {}).get('unique_dimms', 'None')
        for unique_dimm in unique_dimms:
            part_number: list = unique_dimm.get('part')
            count = int(unique_dimm.get('count'))

            if part_number not in inventory_total:
                inventory_total[part_number]: int = count
            else:
                inventory_total[part_number]: int = + count
    except AttributeError:
        pass

    return inventory_total


def remove_key_from_dictionary(dictionary: dict, key: str):
    reaction = dict(dictionary)
    del reaction[key]
    return reaction


def clean_inventory_total(inventory_total: dict) -> dict:
    """
    Inventory data.
    """
    for part_name in inventory_total:
        if part_name == "Unknown":
            inventory_total: dict = remove_key_from_dictionary(inventory_total, part_name)

        elif part_name == "Virtual Disk":
            inventory_total: dict = remove_key_from_dictionary(inventory_total, part_name)

        elif part_name == "":
            inventory_total: dict = remove_key_from_dictionary(inventory_total, part_name)

        elif part_name == "Persistent memory disk":
            inventory_total: dict = remove_key_from_dictionary(inventory_total, part_name)

        elif part_name == "Array":
            inventory_total: dict = remove_key_from_dictionary(inventory_total, part_name)

    return inventory_total


def get_inventory_total():
    """

    """
    get_all_host_ids()
    inventory_total: dict = {}
    for machine_json in hosts_data:
        inventory_total: dict = get_inventory_from_machine(machine_json, inventory_total)
    return clean_inventory_total(inventory_total)


def main_method() -> None:
    """

    """
    inventory_count: dict = get_inventory_count()
    inventory_total: dict = get_inventory_total()

    combine_inventory: dict = {**inventory_count, **inventory_total}

    import json
    foo = json.dumps(combine_inventory, sort_keys=True, indent=4)
    print(foo)
    input()
