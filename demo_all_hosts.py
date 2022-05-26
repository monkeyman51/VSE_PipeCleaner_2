import asyncio
import json
from json import loads, dumps
from pymongo import MongoClient

import aiohttp.client_exceptions
import requests
from aiohttp import ClientSession
from openpyxl import load_workbook
from datetime import datetime
from datetime import date
from time import time, strftime, localtime


async def generate_jsons(machines: list) -> list:
    async with ClientSession() as session:
        all_machines: list = []

        for index, machine_data in enumerate(machines, start=1):
            print(index)
            host_id: str = machine_data["host_id"]
            product_serial: str = machine_data["product_serial"]

            generate_data = {
                'action': 'get_json_data',
                'host_id': f'{host_id}'
            }
            data = {
                'action': 'get_json_data',
                'host_id': f'{product_serial}'
            }
            async with session.post('http://172.30.1.100/console/console_js.php', json=generate_data):
                pass

            async with session.post(url=f'http://172.30.1.100/results/{product_serial}.json', json=data) as response:
                try:
                    json_response = await response.json()
                    all_machines.append(json_response)

                except aiohttp.client_exceptions.ContentTypeError:
                    all_machines.append(machine_data)

        return all_machines


def store_raw_host_data(raw_host_data: dict) -> dict:
    """
    For each real machine / blade in Console Server in the All Hosts tab.

    :param raw_host_data: Not processed data

    Host Data examples:
    - host_ip: 192.168.238.201
    - host_id: 5e7e640a0b4b02141535ead1
    - machine_name: VSE0G8IZUTL-973
    - last_alive: 2021-12-08 07:49:17
    - connection: ex. ALIVE, DEAD, ALMOST DEAD (recent dead)
    """
    host_ip: str = raw_host_data.get("host_ip", "None").upper()
    host_id: str = raw_host_data.get("id", "None")
    machine_name: str = raw_host_data.get("machine_name", "None").upper()
    connection_status: str = raw_host_data.get("connection_status", "None").upper()
    last_found_alive = raw_host_data.get("last_found_alive", "None")
    sku_name: str = raw_host_data["sku_name"]
    serial: str = raw_host_data["serial"]

    return {"host_ip": host_ip,
            "host_id": host_id,
            "product_serial": f"{sku_name}-{serial}",
            "machine_name": machine_name,
            "connection": connection_status,
            "last_alive": last_found_alive,
            "days_last_active": count_days_last_active(last_found_alive)}


def add_host_data() -> dict:
    """
    Account for host filled data.  Ex. Machine Name, Status, Ticket, Location
    """
    return {"machine_name": 0,
            "status": 0,
            "ticket": 0,
            "location": 0}


def create_hosts_raw_data() -> dict:
    """
    Dictionary data structure to store each host data from Console Server All Hosts tab.
    """
    return {"stats": {"hosts": 0,
                      "blades": 0,
                      "active": 0,
                      "inactive": 0,
                      "other": 0,
                      "virtual": 0,
                      "inactive_greater_than": {"1_day": 0,
                                                "1_week": 0,
                                                "1_month": 0,
                                                "3_month": 0,
                                                "6_month": 0,
                                                "1_year": 0}, },
            "machines": []}


def get_machines_raw_data() -> dict:
    """
    Gather all host data from Console Server's All Host page. Purpose is to gather individual host ID and IP for later
    async fetching.
    """
    all_hosts: list = get_console_server_all_hosts()
    hosts_raw_data: dict = create_hosts_raw_data()

    for host_source_data in all_hosts:
        machine_name: str = host_source_data["machine_name"]

        if "VSE0G5" in machine_name:  # Gen 5 blades are decommissioned in VSEI Kirkland site
            pass
        else:
            hosts_raw_data: dict = add_machine_data(host_source_data, hosts_raw_data)

    return hosts_raw_data


def count_days_last_active(last_alive: str) -> int:
    """
    Determine when blade was last active in days.
    """
    last_alive_date: str = last_alive[0:10]
    current_date: str = datetime.today().strftime('%Y-%m-%d')

    blade_year = int(last_alive_date[0:4])
    blade_month = int(last_alive_date[5:7])
    blade_day = int(last_alive_date[8:10])

    current_year = int(current_date[0:4])
    current_month = int(current_date[5:7])
    current_day = int(current_date[8:10])

    current_time = date(current_year, current_month, current_day)
    blade_time = date(blade_year, blade_month, blade_day)

    return (current_time - blade_time).days


def slot_machine_inactive(raw_host_data: dict, hosts_raw_data: dict, virtual_machine_hint: str) -> dict:
    """
    Account for slot machine.
    """
    days_last_active: int = raw_host_data["days_last_active"]
    inactive_greater_than: dict = hosts_raw_data["stats"]["inactive_greater_than"]
    machine_name: str = raw_host_data["machine_name"]

    if virtual_machine_hint not in machine_name:
        if days_last_active >= 1:
            inactive_greater_than["1_day"] += 1

        if days_last_active >= 7:
            inactive_greater_than["1_week"] += 1

        if days_last_active >= 28:
            inactive_greater_than["1_month"] += 1

        if days_last_active >= 84:
            inactive_greater_than["3_month"] += 1

        if days_last_active >= 168:
            inactive_greater_than["6_month"] += 1

        if days_last_active >= 365:
            # import json
            # foo = json.dumps(raw_host_data, sort_keys=True, indent=4)
            # print(foo)
            # input()
            inactive_greater_than["1_year"] += 1

    return hosts_raw_data


def add_machine_data(host_source_data: dict, hosts_raw_data: dict) -> dict:
    """
    Add real machine data that are not virtual machines.
    :param host_source_data:
    :param hosts_raw_data:
    """
    machines_stats: dict = hosts_raw_data["stats"]
    raw_host_data: dict = store_raw_host_data(host_source_data)
    machine_name: str = raw_host_data["machine_name"]

    virtual_machine_hint: str = "-VM"

    machines_stats["hosts"] += 1

    if virtual_machine_hint not in machine_name:
        connection: str = raw_host_data["connection"]

        hosts_raw_data["machines"].append(raw_host_data)
        machines_stats["blades"] += 1

        if connection == "ALIVE":
            machines_stats["active"] += 1

        elif connection == "DEAD":
            machines_stats["inactive"] += 1

        elif connection == "MOSTLY_DEAD":
            # Mostly Dead means blade was recently offline within 10 minutes
            machines_stats["inactive"] += 1

        else:
            machines_stats["other"] += 1

    elif virtual_machine_hint in machine_name:
        machines_stats["virtual"] += 1

    return slot_machine_inactive(raw_host_data, hosts_raw_data, virtual_machine_hint)


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

    try:
        return loads(response.text)

    except:
        return {}


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


def get_generation_from_bios(bios_version: str) -> str:
    """
    Return generation based on mapping to BIOS
    """
    if "C2010" in bios_version or \
            "C2050" in bios_version or \
            "C2020" in bios_version:
        return "GEN 6"

    elif "C2060" in bios_version or \
            "C2265" in bios_version or \
            "C2665" in bios_version or \
            "C2030" in bios_version or \
            "S2260" in bios_version or \
            "C2160" in bios_version or \
            "C2360" in bios_version or \
            "-[IVE" in bios_version or \
            "-[PSE" in bios_version:
        return "GEN 7"

    elif "C215" in bios_version or \
            "S2151" in bios_version or \
            "C2080" in bios_version or \
            "S2180" in bios_version or \
            "C2090" in bios_version:
        return "GEN 8"

    else:
        return "None"


def store_dimm_basic(dimm_serial: str, parts_serial: dict, machine_json: dict) -> dict:
    """

    """
    parts_serial[dimm_serial]["time"]["current_date"]: str = get_current_date()
    parts_serial[dimm_serial]["time"]["last_alive"]: str = machine_json.get("last_found_alive", "Unknown")

    parts_serial[dimm_serial]["location"]["rack"]: str = machine_json.get("location", "Unknown")
    parts_serial[dimm_serial]["location"]["machine"]: str = machine_json.get("machine_name", "Unknown")
    parts_serial[dimm_serial]["location"]["server_id"]: str = machine_json.get("server_id", "Unknown")
    parts_serial[dimm_serial]["location"]["ip_address"]: str = machine_json["net"]["interfaces"][0].get("ip", "Unknown")

    dmi_baseboard: dict = machine_json.get("dmi", {}).get("baseboard", {})
    parts_serial[dimm_serial]["baseboard"]["server_id"]: str = dmi_baseboard["manufacturer"]
    parts_serial[dimm_serial]["baseboard"]["product"]: str = dmi_baseboard["product"]
    parts_serial[dimm_serial]["baseboard"]["serial"]: str = dmi_baseboard["serial"]
    parts_serial[dimm_serial]["baseboard"]["version"]: str = dmi_baseboard["version"]
    parts_serial[dimm_serial]["baseboard"]["machine"]: str = machine_json.get("platform", {}).get("machine", "Unknown")

    parts_serial[dimm_serial]["platform"]["node"]: str = machine_json.get("platform", {}).get("node", "None")
    parts_serial[dimm_serial]["platform"]["release"]: str = machine_json.get("platform", {}).get("release", "Unknown")

    return parts_serial


def get_current_date() -> str:
    """

    :return:
    """
    current_date: str = datetime.today().strftime('%Y-%m-%d')

    current_year = int(current_date[0:4])
    current_month = int(current_date[5:7])
    current_day = int(current_date[8:10])

    return f"{current_month}/{current_day}/{current_year}"


def get_serials_data(machine_json: dict) -> dict:
    """
    Gather raw DIMMS data.
    """
    try:
        dimms_data: list = machine_json["dmi"]["dimms"]
    except KeyError:  # In case queries back incomplete JSON
        return {}

    parts_serial: dict = {}
    for dimm in dimms_data:
        dimm_serial: str = dimm["serial"].strip()

        if dimm_serial != "None" and dimm_serial != "":
            parts_serial: dict = store_dimm_data(dimm_serial, dimm, parts_serial)
            parts_serial: dict = store_dimm_basic(dimm_serial, parts_serial, machine_json)

    return parts_serial


def store_dimm_data(serial_number, dimm, parts_serial) -> dict:
    """

    :param serial_number: DIMM S/N from Console Server
    :param dimm:
    :param parts_serial:
    :return:
    """
    parts_serial[serial_number]: dict = {}

    parts_serial[serial_number]["time"]: dict = {}
    parts_serial[serial_number]["location"]: dict = {}

    parts_serial[serial_number]["location"]["node"]: str = dimm.get("locator").strip().title()
    parts_serial[serial_number]["location"]["slot"]: str = dimm.get("asset"). \
        replace("DIMM_", "").replace("_AssetTag", "")

    parts_serial[serial_number]["baseboard"]: dict = {}
    parts_serial[serial_number]["platform"]: dict = {}

    parts_serial[serial_number]["attributes"]: dict = {}
    parts_serial[serial_number]["attributes"]["part"]: str = dimm.get("part").strip()
    parts_serial[serial_number]["attributes"]["rank"]: str = dimm.get("rank").strip()
    parts_serial[serial_number]["attributes"]["manufacturer"]: str = dimm.get("manufacturer").strip()
    parts_serial[serial_number]["attributes"]["size"]: str = dimm.get("size").strip()
    parts_serial[serial_number]["attributes"]["speed"]: str = dimm.get("speed").strip()

    return parts_serial


def get_mongodb_document(client: MongoClient, collection_name: str, document_name: str) -> MongoClient:
    """
    Get specific record based on client, collection name, and document name.
    """
    return client[collection_name][document_name]


def get_mongodb_company(username: str, password: str, database: str) -> MongoClient:
    """
    Get client based on username, password, and database name.  SSL is true with no certification needed.
    """
    return MongoClient(f"mongodb+srv://{username}:{password}@inventory.daiqz.mongodb.net/{database}")


def access_database_document(database_name: str, document_name: str) -> MongoClient:
    """
    Get client from database
    """
    client = get_mongodb_company('vsei2881', 'FordFocus24', 'test')
    return get_mongodb_document(client, database_name, document_name)


def output_excel(current_serials: dict) -> None:
    """
    Output the
    :param current_serials:
    :return:
    """
    count_serials: int = len(current_serials)

    workbook = load_workbook("DIMMs_template.xlsx")
    worksheet = workbook["Sheet1"]

    print(f"count_serials: {count_serials}")
    for index, current_serial in enumerate(current_serials, start=2):
        serial_data: dict = current_serials[current_serial]
        manufacturer: str = serial_data["attributes"]["manufacturer"]
        part_number: str = serial_data["attributes"]["part"]
        rank: str = serial_data["attributes"]["rank"]
        size: str = serial_data["attributes"]["size"]
        speed: str = serial_data["attributes"]["speed"]
        baseboard_machine: str = serial_data["baseboard"]["machine"]
        baseboard_product: str = serial_data["baseboard"]["product"]
        baseboard_serial: str = serial_data["baseboard"]["serial"]
        baseboard_version: str = serial_data["baseboard"]["version"]
        ip_address: str = serial_data["location"]["ip_address"]
        machine_name: str = serial_data["location"]["machine"]
        node: str = serial_data["location"]["node"]
        rack: str = serial_data["location"]["rack"]
        server_id: str = serial_data["location"]["server_id"]
        slot: str = serial_data["location"]["slot"]
        current_date: str = serial_data["time"]["current_date"]
        last_alive: str = serial_data["time"]["last_alive"]

        date_time: str = strftime("%Y-%m-%d %H:%M:%S", localtime(float(last_alive)))

        # blade_year = last_alive[0:4]
        # blade_month = last_alive[5:7]
        # blade_day = last_alive[8:10]

        worksheet[f"A{index}"]: str = current_serial
        worksheet[f"B{index}"]: str = date_time
        worksheet[f"C{index}"]: str = slot
        worksheet[f"D{index}"]: str = server_id
        worksheet[f"E{index}"]: str = rack
        worksheet[f"F{index}"]: str = node
        worksheet[f"G{index}"]: str = machine_name
        worksheet[f"H{index}"]: str = ip_address
        worksheet[f"I{index}"]: str = baseboard_version
        worksheet[f"J{index}"]: str = baseboard_serial
        worksheet[f"K{index}"]: str = manufacturer
        worksheet[f"L{index}"]: str = part_number
        worksheet[f"M{index}"]: str = rank
        worksheet[f"N{index}"]: str = size
        worksheet[f"O{index}"]: str = speed

    workbook.save("DIMMs_history.xlsx")



def main_method() -> None:
    """

    """
    print(f"start...\n")
    # client_company = get_mongodb_company('joton51', 'FordFocus24', 'test')

    # client_personal = MongoClient("mongodb+srv://joton51:FordFocus24@cluster0.fueyc.mongodb.net/test")
    # print(f"client_company: {client_company}")
    # input()
    # collection = client_company["console_server"]["serials"]
    # collection = client_company.list_database_names()
    # print(f"collection: {collection}")
    # for item in collection:
    #     print(item)
    # collection = access_database_document("console_server", "serials").find()
    # mongodb+srv://<username>:<password>@inventory.daiqz.mongodb.net/test

    # print(f"collection: {collection}")
    # for item_1 in collection:
    #     print(f"item_1: {item_1}")
    # for item in collection:
    #     print(item["_id"])
    # print(collection)
    # collection.insert_one({"_id": "124",
    #                        "history": {},
    #                        "original": {}})

    current_serials: dict = get_current_serials()
    output_excel(current_serials)


def checks_serial_change(current_serial: str, current_serials: dict, last_entry: dict) -> bool:
    """
    Checks for if last database entry has any change with current serial.
    """
    serial_data: dict = current_serials[current_serial]
    entry_date: str = last_entry["time"]["date"]
    entry_part: str = last_entry["attributes"]["part"]
    entry_slot: str = last_entry["location"]["slot"]
    entry_baseboard_version: str = last_entry["baseboard"]["version"]

    serial_date: str = serial_data["time"]["date"]
    serial_part: str = serial_data["attributes"]["part"]
    serial_slot: str = serial_data["location"]["slot"]
    serial_baseboard_version: str = serial_data["baseboard"]["version"]

    if entry_date != serial_date \
            and entry_part != serial_part \
            and entry_slot != serial_slot \
            and entry_baseboard_version != serial_baseboard_version:
        return True
    else:
        return False


def get_current_serials() -> dict:
    """
    Console Server inforamation.
    """
    container_serials: dict = {}

    machines_data: list = get_all_machines_data()
    for machine_data in machines_data:
        machine_last_alive: float = machine_data.get("last_found_alive", 0.0)

        serials_data: dict = get_serials_data(machine_data)

        if serials_data != {}:
            for serial_number in serials_data:
                serial_data: dict = serials_data[serial_number]

                if serial_number in container_serials:
                    serial_last_alive: float = container_serials[serial_number]["time"]["last_alive"]

                    if machine_last_alive < serial_last_alive and machine_last_alive != 0.0:
                        container_serials[serial_number]: dict = serial_data

                else:
                    container_serials[serial_number]: dict = serial_data

    return container_serials


def start_timer():
    start_time: float = time()
    return start_time


def get_all_machines_data() -> list:
    """
    Get all machine data from Console Server
    """
    machines_raw_data: dict = get_machines_raw_data()
    machines: list = machines_raw_data["machines"]

    return asyncio.run(generate_jsons(machines))


main_method()
