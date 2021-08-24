"""
Account for All Machines in Console Server for Kirkland Site.
"""


import requests
from json import loads, dumps
import time
import asyncio
from aiohttp import ClientSession, client_exceptions
from openpyxl import load_workbook


def get_all_hosts_id() -> list:
    """
    Gather all host data from Console Server's All Host page. Purpose is to gather individual host ID and IP for later
    async fetching.
    """
    generate_data = {
        'action': 'get_host_status'
    }
    response = requests.post(url='http://172.30.1.100/console/console_js.php', data=dumps(generate_data))
    all_host: list = loads(response.text)

    alive: int = 0
    dead: int = 0
    vms: int = 0

    host_ids: list = []
    for host in all_host:
        machine_name: str = host.get('machine_name', 'None').upper()

        if '-VM-' not in machine_name:
            host_id: str = host.get('id', 'None')
            connection_status: str = host.get('connection_status', 'None').upper()

            if host_id:
                host_ids.append(host_id)

            if connection_status != 'DEAD':
                alive += 1
            else:
                dead += 1
        elif '-VM-' in machine_name:
            vms += 1
        else:
            print(f'machine_name: {machine_name}')

    print(f'Total Host: {len(all_host)}')
    print(f'Total Machines: {len(host_ids)}')
    print(f'Alive Machines: {alive}')
    print(f'Dead Machines: {dead}')
    print(f'Total VMs: {vms}')
    # input()

    return host_ids


def generate_console_server_json(host_id: str) -> str:
    """
    Generates the JSON data from the Host Details page using the Host ID.
    Returns product-serial if JSON data is generated properly.

    :param host_id: found on Host Details page in URL ?host_id=<some_id>get_console_server_json
    :return: returns product-serial string for getting JSON data
    """
    generate_data = {
        'action': 'get_json_data',
        'host_id': f'{host_id}'
    }
    host_name_data = {
        'action': 'get_host_name_data',
        'host_id': f'{host_id}'
    }

    requests.post(url='http://172.30.1.100/console/console_js.php', json=generate_data)
    host_response = requests.post(url='http://172.30.1.100/console/console_js.php', json=host_name_data)

    product = loads(str(host_response.text))['host_name_data']['product']
    serial = loads(str(host_response.text))['host_name_data']['serial']

    return f'{product}-{serial}'


def generate_json_data(host_id: str) -> None:
    """
    Generate latest Console Server host data.
    """
    generate_data = {
        'action': 'get_json_data',
        'host_id': f'{host_id}'
    }
    requests.post(url='http://172.30.1.100/console/console_js.php', json=generate_data)


def get_console_server_json(product_serial: str) -> dict:
    """
    Gets the Generated data using the product_serial string and creates JSON file.
    Returns JSON of Host within Console Server.

    :param product_serial: string from generate_json_data method
    :return: JSON data of Host
    """
    data = {
        'action': 'get_json_data',
        'host_id': f'{product_serial}'
    }
    response = requests.post(url=f'http://172.30.1.100/results/{product_serial}.json', json=data)

    return loads(response.text)


async def generate_individual_json(host_id: str, index: int):
    """
    Grabs the information from Azure Devops per session depending on how many tickets in the form of URls
    """
    print(f'index: {index}')
    headers = {
        'action': 'get_json_data',
        'host_id': f'{host_id}'
    }
    # pass
    async with ClientSession() as session:
        # pass
        async with session.post(url='http://172.30.1.100/console/console_js.php', json=headers) as response:
            print(response)
        #     await asyncio.sleep(0.5)
        #     ticket_data = await response.text()
    #
    # return ticket_data


async def run_generate_json(host_ids):
    """
    Creates tasks for executing the event loop. Tasks are just requests sent quantified by number of unique tickets
    found in the Console Server
    """
    tasks: list = [asyncio.create_task(generate_individual_json(host_id, index))
                   for index, host_id in enumerate(host_ids, start=1)]

    return await asyncio.gather(*tasks)


def convert_machine_last_online_to_days(last_found_alive: float):
    """
    Convert seconds into days.  For getting machines that have been turned on for the last day.
    """
    # Account for non-existent last found alive.
    if last_found_alive == 0.00:
        return -1

    else:
        current_epoch: float = time.time()
        time_difference: float = current_epoch - last_found_alive
        seconds_in_day: float = 86400.00

        return int(time_difference / seconds_in_day)


# def convert_machine_last_online_to_days(last_found_alive: float):
#     """
#     Get days last found alive for reference later.
#     """
#     days = last_found_alive / 86400.00
#
#     if days < 1:
#         return 'Less 1 Day'
#
#     else:
#         first_part = str(days).split('.')[0]
#
#         if first_part == '1':
#             return f'{first_part} day last online'
#
#         else:
#             return f'{first_part} days last online'


def main_method() -> None:
    """

    :return:
    """
    # start: float = time.time()
    host_ids: list = get_all_hosts_id()
    # host_data: list = asyncio.run(get_ticket_data(ticket_urls))

    # asyncio.run(run_generate_json(host_ids))
    all_machines: list = []
    for index, host_id in enumerate(host_ids, start=1):
        generate_json_data(host_id)
        product_serial: str = generate_console_server_json(host_id)
        host_data: dict = get_console_server_json(product_serial)

        sku_name: str = get_sku_name(host_data)
        machine_name: str = get_machine_name(host_data)
        machine_serial: str = get_machine_serial(host_data)
        last_found_online: float = get_last_found_online(host_data)
        days_last_online: int = convert_machine_last_online_to_days(last_found_online)

        machine_data: dict = {'machine_name': machine_name,
                              'machine_serial': machine_serial,
                              'sku_name': sku_name,
                              'days_last_online': days_last_online}

        all_machines.append(machine_data)

    workbook = load_workbook('all_machines.xlsx')
    worksheet = workbook['all_machines']

    worksheet['A1'].value = 'Machine Name'
    worksheet['B1'].value = 'SKU Name'
    worksheet['C1'].value = 'Serial'
    worksheet['D1'].value = 'Days Last Online'

    for index, row in enumerate(all_machines, start=2):
        worksheet[f'A{index}'].value = row['machine_name']
        worksheet[f'B{index}'].value = row['sku_name']
        worksheet[f'C{index}'].value = row['machine_serial']
        worksheet[f'D{index}'].value = row['days_last_online']

    workbook.save('all_machines.xlsx')


def get_machine_location(host_data: dict) -> str:
    """
    Get machine location.
    """
    if host_data == {}:
        return "None"

    else:
        return host_data.get("location", "None").upper().strip()


def get_last_found_online(host_data: dict) -> float:
    """
    Get last found online.
    """
    if host_data == {}:
        return 0.00

    else:
        return host_data.get("last_found_alive", 0.00)


def get_machine_name(host_data: dict) -> str:
    """
    Get Machine Name from host.  ex. Machine
    """
    if host_data == {}:
        return ""

    else:
        machine_name: str = host_data.get("machine_name", "").strip().upper()

        if "VSE" not in machine_name:
            return "Invalid Name"

        elif "CMA" in machine_name:
            return "Invalid Name"

        elif "VSE" in machine_name:
            return machine_name


def get_sku_name(host_data: dict) -> str:
    """
    Get product name or SKU name.  ex. C2080
    """
    if host_data == {}:
        return "None"

    else:
        return host_data.get("baseboard", {}).\
            get("baseboard", {}).\
            get("product", "None")


def get_machine_serial(host_data: dict) -> str:
    """
    Get machine serial number. ex. 9J1000019220892J0G1
    """
    if host_data == {}:
        return "None"

    else:
        return host_data.get("dmi", {}).\
            get("baseboard", {}).\
            get("serial", "None").upper().strip()


def get_machine_host_ip(host_data: dict) -> str:
    """
    ex. 192.168.231.250
    """
    if host_data == {}:
        return "None"

    else:
        host_ip: str = host_data.get("net", {}).get("interfaces", {})[0].get("ip", "None")

        if "." not in host_ip:
            return "None"

