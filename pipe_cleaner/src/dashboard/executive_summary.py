"""
Executive Summary meant for
"""
from csv import reader, writer
import asyncio
import base64
import subprocess
import json
import os
import sys
from openpyxl import load_workbook, Workbook

import requests
from aiohttp import ClientSession, client_exceptions
from bs4 import BeautifulSoup
from colorama import Fore, Style

from pipe_cleaner.src.credentials import AccessADO
from pipe_cleaner.src.data_access import request_ado_json
from pipe_cleaner.src.credentials import AccessADO as Ado


def get_current_tickets() -> list:
    """
    Get current tickets downloaded from ADO.
    """
    with open("settings/current_tickets.csv") as file:
        csv_data = reader(file, delimiter=",", quotechar='"')

        ticket_numbers: list = []
        for index, csv_row in enumerate(csv_data, start=0):
            if index == 0:
                pass

            else:
                ticket_number: str = csv_row[1]
                ticket_numbers.append(ticket_number)
        return ticket_numbers


def get_past_tickets() -> list:
    """
    Get past tickets downloaded from ADO.
    """
    with open("settings/past_tickets.csv") as file:
        csv_data = reader(file, delimiter=",", quotechar='"')

        ticket_numbers: list = []
        for index, csv_row in enumerate(csv_data, start=0):
            ticket_number: str = csv_row[0]
            if index == 0:
                pass

            elif ticket_number.isdigit():
                ticket_numbers.append(ticket_number)

        return ticket_numbers


def get_ticket_urls(unique_tickets: list) -> list:
    return [f'https://azurecsi.visualstudio.com/CSI%20Commodity%20Qualification/_apis/wit/workitems?'
            f'id={ticket}&$expand=all&api-version=5.1' for ticket in unique_tickets]


def get_ticket_numbers() -> list:
    """
    Get past and current ticket numbers.
    """
    current_tickets: list = get_current_tickets()
    past_tickets: list = get_past_tickets()
    total_tickets: list = current_tickets + past_tickets
    return list(set(total_tickets))


async def fetch_site(url, headers):
    """
    Grabs the information from Azure Devops per session depending on how many tickets in the form of URls
    :param url:
    :param headers:
    :return:
    """
    async with ClientSession(headers=headers) as session:
        try:
            async with session.get(url) as response:
                await asyncio.sleep(0.5)
                ticket_data = await response.text()

        # Rare occurrence dealing with Async,
        except client_exceptions.ClientOSError:
            print(f'\t{Fore.RED}[WinError 10054] An existing connection was forcibly closed '
                  f'by the remote host{Style.RESET_ALL}')
            print(f'\tPress {Fore.LIGHTBLUE_EX}ENTER{Style.RESET_ALL} to exit Pipe Cleaner...', end='')
            input()
            sys.exit()

    return ticket_data


def access_due_dates(ticket_json) -> dict:
    """
    Get actual due date from ADO
    :param ticket_json:
    :return:
    """
    due_dates: dict = {}

    try:
        expected_task_start = ticket_json['fields']['AzureCSI-V1.1.ExpectedTaskStart']
        due_dates['expected_task_start'] = expected_task_start
    except KeyError:
        pass

    try:
        expected_task_completion = ticket_json['fields']['AzureCSI-V1.1.ExpectedTaskCompletion']
        due_dates['expected_task_completion'] = expected_task_completion
    except KeyError:
        pass

    try:
        actual_qual_start_date = ticket_json['fields']['Custom.ActualQualStartDate']
        due_dates['actual_qual_start_date'] = actual_qual_start_date
    except KeyError:
        pass

    try:
        actual_qual_end_date = ticket_json['fields']['Custom.ActualQualEndDate']
        due_dates['actual_qual_end_date'] = actual_qual_end_date
    except KeyError:
        pass

    return due_dates


def clean_html_text(field_data):
    clean_text = str(field_data.text). \
        replace('\n', ''). \
        replace('  ', ' '). \
        replace('\xa0', ''). \
        replace('\u200b', ''). \
        replace('\u2013', '')
    return clean_text.strip()


def get_clean_table_data(table_rows: list) -> dict:
    """
    Gather data from description table given a TRR within ADO
    :param table_rows:
    :return:
    """
    all_table_data: dict = {}

    all_part_numbers: list = []
    for index, row in enumerate(table_rows):

        table_row_data: list = []
        for field_data in row.findAll(['td']):
            if field_data is None or field_data == '' or not field_data:
                continue
            else:
                clean_text: str = clean_html_text(field_data)

                if 'RQUEST TYPE' in clean_text.upper():
                    table_row_data.append('request type')
                else:
                    table_row_data.append(clean_text)

        try:
            # Replacing space for underscore for easier key calls for values later
            component_key = str(table_row_data[0]).lower().strip().replace(' ', '_')
            raw_component_key = table_row_data[0].upper()
            clean_value: str = clean_component_value(table_row_data[1])

            if 'PART' in raw_component_key and 'NUMBER' in raw_component_key:
                if not clean_value or clean_value == 'None':
                    pass
                else:
                    all_part_numbers.append(clean_value)

            elif 'DESCRIPTION' in raw_component_key:
                if not clean_value or clean_value == 'None':
                    pass
                else:
                    all_part_numbers.append(clean_value)

            elif 'MODEL' in raw_component_key and 'NUMBER' in raw_component_key:
                if not clean_value or clean_value == 'None':
                    pass
                else:
                    all_part_numbers.append(clean_value)

            elif 'FIRMWARE' in raw_component_key:
                if not clean_value or clean_value == 'None':
                    pass
                else:
                    all_part_numbers.append(clean_value)

            elif 'FIRMWARE' in raw_component_key and 'N-1' in raw_component_key:
                if not clean_value or clean_value == 'None':
                    pass
                else:
                    all_part_numbers.append(clean_value)

            potential_reference_test_plans: list = []
            # In case Reference Test Plans are empty and MFST has entered space
            # Doing so will disallow PipeCleaner to call proper key for the value
            if 'reference_test' in component_key:
                potential_reference_test_plans_index = index + 1
                field_reference_test = table_rows[potential_reference_test_plans_index].find(['td'])
                potential = str(field_reference_test.text).replace('\n', '').replace('  ', ' ').replace('\xa0', '')
                potential_exception: str = exception_check_reference_test_plans(potential)
                potential_reference_test_plans.append(potential_exception)

            if 'reference_test' in component_key:
                clean_key: str = 'reference_test_plans'
                clean_value: str = potential_reference_test_plans[0]
                all_table_data[clean_key] = clean_value
            else:
                clean_key: str = clean_component_key(component_key)
                clean_value: str = clean_component_value(table_row_data[1])
                all_table_data[clean_key] = clean_value

        except IndexError:
            pass

    all_table_data['all_part_numbers']: list = all_part_numbers

    return all_table_data


def clean_component_value(component_value: str) -> str:
    """
    Cleans component to make easier to call value through key later
    :param component_value:
    :return:
    """
    # print(f'[ Before ] Value Component: {component_value}')
    # Raise for easier comparison
    upper_component_key = component_value.upper()

    clean_key: str = upper_component_key. \
        replace(' - ', ' '). \
        replace('N/A', 'None')

    # print(f'[ After  ] Value Component: {component_value}')

    return clean_key


def clean_component_key(component_key: str) -> str:
    """
    Cleans component to make easier to call value through key later
    :param component_key:
    :return:
    """
    # Raise for easier comparison
    upper_component_key = component_key.upper()

    clean_key: str = upper_component_key. \
        replace(' - ', ' '). \
        replace('#', ''). \
        replace(' : ', ''). \
        replace('SATA', ''). \
        replace('\u00e2', ''). \
        replace('\u20ac', ''). \
        replace('\u201c', ''). \
        replace('(', ''). \
        replace(')', ''). \
        replace(r'/', ''). \
        replace('_-_', '_'). \
        replace('__', '_')

    clean_component_lower: str = clean_key.lower()

    return clean_component_lower


def exception_check_reference_test_plans(potential_test_plans: str):
    """
    Ensure that if an initial reference test plan from the TRR show empty, the next Table Row shows the in
    :param potential_test_plans:
    """
    upper_test_plans: str = potential_test_plans.upper()
    possible_key_words: list = [
        'Q1', 'Q2', 'Q3', 'Q4',
        '2020', '2021', '2022', '2023', '2024', '2025'
                                                '2026', '2027', '2028', '2029', '2030'
    ]
    for key_work in possible_key_words:
        if key_work in potential_test_plans:
            return possible_key_words
    else:
        return 'None'


def store_tickets_data(raw_tickets_data: list) -> dict:
    """

    """
    all_tickets_data: dict = {}

    # Used for checking in CRDs, see whether components are in there or missing
    all_bios: list = []
    all_bmc: list = []
    all_cpld: list = []
    all_os: list = []

    print(f'\n\t=====================================================================')
    print(f'\t  Tickets (TRRs) - Collecting and Processing Data')
    print(f'\t=====================================================================')
    print(f'\t\t  STATUS   |  REASON    |  TICKET')

    for raw_ticket_data in raw_tickets_data:
        ticket_data: dict = {}
        ticket_json: dict = json.loads(raw_ticket_data)
        # foo = json.dumps(ticket_json, sort_keys=True, indent=4)
        # print(foo)
        # input()

        try:
            # For some reason the ticket_json["id"] data type is int
            ticket_id: str = str(ticket_json["id"])
            print(f'\t\t- Collect  |  {Fore.GREEN}Success{Style.RESET_ALL}   |  {ticket_id}')

            # Table Data
            table_data = ticket_json['fields']['System.Description']
            table_data_soup = BeautifulSoup(table_data, 'html.parser')

            table_rows: list = table_data_soup.findAll('tr')
            table_data: dict = get_clean_table_data(table_rows)

            # Checking for CRDs later
            server_bios: str = table_data.get('server_bios', 'None')
            server_bmc: str = table_data.get('server_bmc', 'None')
            server_cpld: str = table_data.get('server_cpld', 'None')
            server_os: str = table_data.get('server_os', 'None')

            if server_bios != 'None' and server_bios != '':
                all_bios.append(server_bios)
            if server_bmc != 'None' and server_bmc != '':
                all_bmc.append(server_bmc)
            if server_cpld != 'None' and server_cpld != '':
                all_cpld.append(server_cpld)
            if server_os != 'None' and server_os != '':
                all_os.append(server_os)

            # Table Data Stored
            ticket_data['table_data']: dict = table_data

            ticket_data['title'] = ticket_json['fields']['System.Title']

            ticket_data['trr_type']: int = ticket_json.get('fields', {}).get('Custom.TRRType')

            # Due Dates
            ticket_data['due_dates']: dict = access_due_dates(ticket_json)

            # Assigned To
            ticket_data['assigned_to']: str = ticket_json['fields']['System.AssignedTo']['displayName']

            # State of Qual
            ticket_data['state']: str = ticket_json['fields']['System.State']

            # Test Plans Hyperlink, used later for Async

            all_tickets_data[ticket_id] = ticket_data

        except json.decoder.JSONDecodeError:
            print(f'\tMicrosoft Azure DevOps did not return all data.')
            print(f'\tError is at TRR {ticket_id}. Recollecting...\n')

        except KeyError:
            print(f'\tNo Description Table... Grabbing from Summary Table Instead in TRR')

            # For some reason the ticket_json["id"] data type is int
            try:
                ticket_id = str(ticket_json["id"])
                # print(f'\t\t- Collect  |  {Fore.GREEN}Success{Style.RESET_ALL}   |  {ticket_id}')

                # Table Data
                table_data = ticket_json['fields']['Custom.Summary']
                table_data_soup = BeautifulSoup(table_data, 'html.parser')

                table_rows: list = table_data_soup.findAll('tr')
                table_data: dict = get_clean_table_data(table_rows)

                # Checking for CRDs later
                server_bios: str = table_data.get('server_bios', 'None')
                server_bmc: str = table_data.get('server_bmc', 'None')
                server_cpld: str = table_data.get('server_cpld', 'None')
                server_os: str = table_data.get('server_os', 'None')

                if server_bios != 'None' and server_bios != '':
                    all_bios.append(server_bios)
                if server_bmc != 'None' and server_bmc != '':
                    all_bmc.append(server_bmc)
                if server_cpld != 'None' and server_cpld != '':
                    all_cpld.append(server_cpld)
                if server_os != 'None' and server_os != '':
                    all_os.append(server_os)

                # Table Data Stored
                ticket_data['table_data']: dict = table_data

                ticket_data['title'] = ticket_json['fields']['System.Title']

                ticket_data['trr_type']: int = ticket_json.get('fields', {}).get('Custom.TRRType')

                # Due Dates
                ticket_data['due_dates']: dict = access_due_dates(ticket_json)

                # Assigned To
                ticket_data['assigned_to']: str = ticket_json['fields']['System.AssignedTo']['displayName']

                # State of Qual
                ticket_data['state']: str = ticket_json['fields']['System.State']

                all_tickets_data[ticket_id] = ticket_data
            except KeyError:
                pass

    return all_tickets_data


async def get_ticket_data(ticket_urls):
    """
    Creates tasks for executing the event loop. Tasks are just requests sent quantified by number of unique tickets
    found in the Console Server
    """
    user_password: str = f'{AccessADO.token_name}:{AccessADO.personal_access_token}'
    base64_user_password = base64.b64encode(user_password.encode()).decode()
    headers = {'Authorization': f'Basic {base64_user_password}'}

    tasks = [asyncio.create_task(fetch_site(request, headers)) for request in ticket_urls]

    return await asyncio.gather(*tasks)


def get_ticket_states(ticket_data: dict) -> dict:
    """
    Get states of each ticket.
    """
    all_states: dict = {}

    for ticket in ticket_data:
        current_data: dict = ticket_data[ticket]
        state: str = current_data.get("state", "None")

        if state == "None":
            pass

        elif state in all_states:
            all_states[state] += 1

        else:
            all_states[state] = 1
    return all_states


def get_ticket_types(ticket_data: dict) -> dict:
    """
    Get primary or secondary of each ticket.
    """
    ticket_types: dict = {"primary": 0, "secondary": 0}

    for ticket in ticket_data:
        current_data: dict = ticket_data[ticket]
        ticket_type: str = current_data.get("trr_type", "None")

        if ticket_type == "None":
            pass

        elif ticket_type == 1:
            ticket_types["primary"] += 1

        elif ticket_type == 2:
            ticket_types["secondary"] += 1
    return ticket_types


def get_ticket_commodity(ticket_data: dict) -> dict:
    """
    Get primary or secondary of each ticket.
    """
    ticket_commodities: dict = {}

    for ticket in ticket_data:
        current_data: dict = ticket_data[ticket]

        request_type: str = current_data.get("table_data", {}).get("request_type", "None"). \
            upper().replace(" TEST", "").replace("TEST", "")
        print(f'request_type: {request_type}')

        if request_type in ticket_commodities:
            ticket_commodities[request_type] += 1

        else:
            ticket_commodities[request_type] = 1

    return ticket_commodities


def write_tickets_data(tickets_data: dict) -> None:
    """

    """
    workbook = load_workbook("settings/weekly_report.xlsx")
    worksheet = workbook.get_sheet_by_name("weekly")

    for index, ticket in enumerate(tickets_data, start=2):
        current_data: dict = tickets_data[ticket]

        state: str = current_data.get("state", "None")
        ticket_type: str = current_data.get("ticket_type", "None")
        request_type: str = current_data.get("table_data", {}).get("request_type", "None"). \
            upper().replace(" TEST", "").replace("TEST", "")
        assigned_to: str = current_data.get("assigned_to", "None")
        expected_start: str = current_data.get("due_dates", {}).get("expected_task_start", "None")
        expected_end: str = current_data.get("due_dates", {}).get("expected_task_completion", "None")
        actual_start: str = current_data.get("due_dates", {}).get("actual_qual_start_date", "None")
        actual_end: str = current_data.get("due_dates", {}).get("actual_qual_end_date", "None")

        worksheet[f"A{index}"]: str = ticket
        worksheet[f"B{index}"]: str = state
        worksheet[f"C{index}"]: str = ticket_type
        worksheet[f"D{index}"]: str = request_type
        worksheet[f"E{index}"]: str = assigned_to
        worksheet[f"F{index}"]: str = expected_start
        worksheet[f"G{index}"]: str = expected_end
        worksheet[f"H{index}"]: str = actual_start
        worksheet[f"I{index}"]: str = actual_end

    workbook.save("ado_data.xlsx")


def main_method() -> None:
    """
    Executive Summary
    """
    unique_tickets: list = get_ticket_numbers()
    ticket_urls: list = get_ticket_urls(unique_tickets)
    raw_tickets_data: list = asyncio.run(get_ticket_data(ticket_urls))
    tickets_data: dict = store_tickets_data(raw_tickets_data)

    write_tickets_data(tickets_data)

    # ticket_states: dict = get_ticket_states(tickets_data)
    # ticket_types: dict = get_ticket_types(tickets_data)
    # ticket_commodities: dict = get_ticket_commodity(tickets_data)
    #
    # import json
    # foo = json.dumps(ticket_commodities, sort_keys=True, indent=4)
    # print(foo)
    # input()
