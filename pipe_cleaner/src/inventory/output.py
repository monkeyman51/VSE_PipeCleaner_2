"""
Alternative model to grab inventory data from Console Server
"""
import asyncio
import os
import time
from datetime import datetime
from json import loads, dumps

import requests
from aiohttp import ClientSession
from openpyxl import load_workbook
from openpyxl.styles import Alignment


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


def get_hosts_data() -> list:
    """
    Gather all host data from Console Server's All Host page. Purpose is to gather individual host ID and IP for later
    async fetching.  Duplicates... 1153(Unique) -> 1171(Count) = 18 Delta
    """
    all_hosts: list = get_console_server_all_hosts()

    serial_numbers: list = []

    hosts_data: list = []
    for host in all_hosts:

        host_data: dict = {}
        host_ip: str = host.get("host_ip", "None").upper()
        host_id: str = host.get("id", "None").upper()
        serial: str = host.get("serial", "None").upper()

        if serial not in serial_numbers:
            serial_numbers.append(serial)

            host_data["host_ip"]: str = host_ip
            host_data["host_id"]: str = host_id
            host_data["serial"]: str = serial

            hosts_data.append(host_data)

    return hosts_data


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
    days = last_found_alive / 86_400.00
    # input()
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


def calculate_last_time_alive(raw_last_found_alive: str):
    """
    Calculate last time alive in non-seconds. Raw last time alive is in seconds and ain't no body got time for that.
    :param raw_last_found_alive:
    :return:
    """
    diff = float(time.time()) - float(raw_last_found_alive)
    return diff


def get_all_machines_data(hosts_data: list) -> list:
    """
    Get all machine data from Console Server
    """
    all_machines: list = []

    workbook = load_workbook('settings/inventory_output.xlsx')
    worksheet = workbook['Serial Numbers']

    # await generate_console_server(hosts_data)

    count: int = 2
    for host_data in hosts_data:

        host_id: str = host_data["host_id"]

        generate_json_data(host_id)
        product_serial: str = generate_console_server_json(host_id)
        machine_data: dict = get_console_server_json(product_serial)

        # import json
        # foo = json.dumps(machine_data, sort_keys=True, indent=4)
        # print(foo)
        # input()

        sku_name: str = machine_data["dmi"]["system"]["product"]

        machine_name: str = machine_data.get("machine_name", "None")

        if "-VM-" in machine_name:
            pass

        else:
            host_serial: str = host_data["serial"]
            raw_last_found_alive: str = machine_data.get("last_found_alive", "None")
            connection_status: str = machine_data["connection_status"]
            location: str = machine_data.get("location", "None")
            last_alive = calculate_last_time_alive(raw_last_found_alive)
            last_found_online_in_days: str = get_last_active(last_alive)

            state: str = get_state(connection_status, last_found_online_in_days)

            unique_nvmes: list = machine_data["nvme"]["nvmes"]
            for unique_nvme in unique_nvmes:
                serial: str = unique_nvme.get("serial")
                part_number: str = unique_nvme.get("model")
                commodity_type: str = "NVMe"

                result = write_cell(commodity_type, count, host_serial, last_found_online_in_days,
                                    machine_name, part_number, serial, state, worksheet, location, sku_name)

                count += result

            unique_disks: list = machine_data["disk"]["disks"]
            for unique_disk in unique_disks:
                serial: str = unique_disk.get("serial", "None")
                part_number: str = unique_disk.get("model", "None")
                commodity_type: str = "Disk"

                result = write_cell(commodity_type, count, host_serial, last_found_online_in_days,
                                    machine_name, part_number, serial, state, worksheet, location, sku_name)
                count += result

            unique_dimms: list = machine_data["dmi"]["dimms"]
            for unique_dimm in unique_dimms:
                serial: str = unique_dimm.get("serial", "None")
                part_number: str = unique_dimm.get("part", "None")
                commodity_type: str = "DIMM"

                result = write_cell(commodity_type, count, host_serial, last_found_online_in_days,
                                    machine_name, part_number, serial, state, worksheet, location, sku_name)
                count += result

    workbook.save('inventory_output.xlsx')

    return all_machines


def write_cell(commodity_type, count, host_serial, last_found_online_in_days, machine_name, part_number, serial, state,
               worksheet, location, sku_name):
    if part_number == "Virtual Disk":
        return 0
    elif part_number == "Persistent Memory Disk":
        return 0
    elif part_number == "Array":
        return 0
    elif part_number == "Cruzer Fit":
        return 0
    elif part_number == "LSI Cobra":
        return 0

    else:
        print(count)
        worksheet[f'A{str(count)}'].value = machine_name
        worksheet[f'B{str(count)}'].value = location
        worksheet[f'C{str(count)}'].value = sku_name
        worksheet[f'D{str(count)}'].value = host_serial
        worksheet[f'E{str(count)}'].value = part_number
        worksheet[f'F{str(count)}'].value = serial
        worksheet[f'G{str(count)}'].value = commodity_type
        worksheet[f'H{str(count)}'].value = state
        worksheet[f'I{str(count)}'].value = last_found_online_in_days

        worksheet[f'A{str(count)}'].alignment = Alignment(horizontal='center')
        worksheet[f'B{str(count)}'].alignment = Alignment(horizontal='center')
        worksheet[f'C{str(count)}'].alignment = Alignment(horizontal='center')
        worksheet[f'D{str(count)}'].alignment = Alignment(horizontal='center')
        worksheet[f'E{str(count)}'].alignment = Alignment(horizontal='center')
        worksheet[f'F{str(count)}'].alignment = Alignment(horizontal='center')
        worksheet[f'G{str(count)}'].alignment = Alignment(horizontal='center')
        worksheet[f'H{str(count)}'].alignment = Alignment(horizontal='center')
        worksheet[f'I{str(count)}'].alignment = Alignment(horizontal='center')

        return 1


def get_state(connection_status, last_found_online_in_days):
    if 'ALIVE' in connection_status.upper():
        return "ONLINE <= 24 Hours"

    elif int(last_found_online_in_days) <= 30:
        return "OFFLINE <= 30 Days"

    elif int(last_found_online_in_days) > 30:
        return "OFFLINE > 30 Days"


async def generate_site(host_id: str):
    """

    """
    url: str = 'http://172.30.1.100/console/console_js.php'
    generate_data: dict = {
        'action': 'get_json_data',
        'host_id': f'{host_id}'
    }
    async with ClientSession() as session:
        async with session.post(url, json=generate_data) as response:
            return await response.text()


async def generate_console_server_data(hosts_id: list):
    """

    """
    tasks = [asyncio.create_task(generate_site(host_id)) for host_id in hosts_id]

    return await asyncio.gather(*tasks)


def generate_console_server(hosts_data: list):
    """

    """
    hosts_id: list = get_hosts_id(hosts_data)
    # async def get(url):
    #     async with aiohttp.ClientSession() as session:
    #         async with session.get(url) as response:
    #             return await response.content.read()
    #
    # loop = asyncio.get_event_loop()
    # tasks = [asyncio.ensure_future(get("http://example.com"))]
    # loop.run_until_complete(asyncio.wait(tasks))
    # print("Results: %s" % [task.result() for task in tasks])

    return asyncio.run(generate_console_server_data(hosts_id))


def get_hosts_id(hosts_data: list) -> list:
    """
    Extract host_id
    """
    hosts_id: list = []
    for host_data in hosts_data:
        host_id: str = host_data["host_id"]
        hosts_id.append(host_id)
    return hosts_id


def main_method() -> None:
    """

    """
    hosts_data: list = get_hosts_data()  # 1171
    get_all_machines_data(hosts_data)
    os.system(fr'start EXCEL.EXE inventory_output.xlsx')
