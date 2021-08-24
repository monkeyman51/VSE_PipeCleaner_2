import asyncio
from json import loads, dumps

import requests
from aiohttp import ClientSession
from openpyxl import load_workbook
from datetime import datetime


def get_days_since_last_active(last_found_alive: float) -> str:
    """

    """
    last_active_date = str(last_found_alive)[0:11]
    today_date: str = datetime.today().strftime('%Y-%m-%d')

    blade_year = int(last_active_date[0:4])
    blade_month = int(last_active_date[8:11])
    blade_day = int(last_active_date[5:7])

    today_year = int(today_date[0:4])
    today_month = int(today_date[8:11])
    today_day = int(today_date[5:7])

    diff_year: int = blade_year - today_year
    diff_month: int = blade_month - today_month
    diff_day: int = blade_day - today_day

    print(f'last_active_date: {last_active_date}')
    print(f'today_date: {today_date}')
    print(f'diff_year: {diff_year} / {diff_month} / {diff_day}\n')

    return str(last_found_alive)[0:10]


def get_all_hosts_data() -> dict:
    """
    Gather all host data from Console Server's All Host page. Purpose is to gather individual host ID and IP for later
    async fetching.
    """
    all_hosts: list = get_console_server_all_hosts()

    all_hosts_data: dict = {}

    active_machines: int = 0
    inactive_machines: int = 0
    virtual_machines: int = 0

    host_ids: list = []
    for host in all_hosts:

        # import json
        # foo = json.dumps(host, sort_keys=True, indent=4)
        # print(foo)
        # input()

        machine_name: str = host.get('machine_name', 'None').upper()
        host_ip: str = host.get('host_ip', 'None').upper()

        if '-VM-' not in machine_name:
            host_id: str = host.get('id', 'None')
            connection_status: str = host.get('connection_status', 'None').upper()
            last_found_alive = host.get('last_found_alive', 'None')
            print(f'last_found_alive: {host_ip} - {last_found_alive}')
            # date_last_alive: str = get_days_since_last_active(last_found_alive)
            # print(f'date_last_alive: {date_last_alive}')

            if host_id:
                host_ids.append(host_id)

            if connection_status != 'DEAD':
                active_machines += 1
            else:
                inactive_machines += 1

        elif '-VM-' in machine_name:
            virtual_machines += 1

    all_hosts_data["total_host"] = len(all_hosts)
    all_hosts_data["total_machines"] = len(host_ids)
    all_hosts_data["active_machines"] = active_machines
    all_hosts_data["inactive_machines"] = inactive_machines
    all_hosts_data["total_virtual_machines"] = virtual_machines
    all_hosts_data["host_ids"] = host_ids

    return all_hosts_data


def get_console_server_all_hosts() -> list:
    """
    Get all hosts found in the All Hosts section of ZT Console Server.
    """
    generate_data: dict = {
        'action': 'get_host_status'
    }
    response = requests.post(url='http://172.30.1.100/console/console_js.php', data=dumps(generate_data))
    return loads(response.text)


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
    async with ClientSession() as session:
        async with session.post(url='http://172.30.1.100/console/console_js.php', json=headers) as response:
            print(response)


async def run_generate_json(host_ids):
    """
    Creates tasks for executing the event loop. Tasks are just requests sent quantified by number of unique tickets
    found in the Console Server
    """
    tasks: list = [asyncio.create_task(generate_individual_json(host_id, index))
                   for index, host_id in enumerate(host_ids, start=1)]

    return await asyncio.gather(*tasks)


def get_last_active(last_found_alive):
    days = last_found_alive / 86400.00

    if days < 1:
        return 'Less than 1 Day'
    else:
        first_part = str(days).split('.')[0]
        if first_part == '1':
            return f'{first_part} day last online'
        else:
            return f'{first_part} days last online'


def main_method() -> None:
    """

    """
    all_hosts_data: dict = get_all_hosts_data()

    import json
    foo = json.dumps(all_hosts_data, sort_keys=True, indent=4)
    print(foo)
    input()

    host_ids: list = all_hosts_data["host_ids"]

    all_machines: list = get_all_machines_data(host_ids)

    workbook = load_workbook('all_machines.xlsx')
    worksheet = workbook['all_machines']

    worksheet['A1'].value = 'Machine'
    worksheet['B1'].value = 'Last Alive'

    for index, row in enumerate(all_machines, start=2):
        worksheet[f'A{index}'].value = row['machine_name']
        worksheet[f'B{index}'].value = row['last_found_alive']

    workbook.save('all_machines.xlsx')


def get_all_machines_data(host_ids: list) -> list:
    """
    Get all machine data from Console Server
    """
    all_machines: list = []

    for index, host_id in enumerate(host_ids, start=1):
        print(index)
        generate_json_data(host_id)
        product_serial: str = generate_console_server_json(host_id)
        host_data: dict = get_console_server_json(product_serial)

        try:
            last_found_alive = float(host_data.get('last_found_alive'))
            machine_data: dict = {'machine_name': host_data.get("machine_name"),
                                  'last_found_alive': get_last_active(last_found_alive)}
            all_machines.append(machine_data)

        except KeyError:
            print(f'machine_name: {host_data.get("machine_name")}')

        except TypeError:
            pass

    return all_machines


main_method()
