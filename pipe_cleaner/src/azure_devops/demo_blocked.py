"""
Gather information.
"""
"""
Get count for all commodity types.
"""
import csv
import asyncio
import base64
import json
import sys

from aiohttp import ClientSession, client_exceptions
from bs4 import BeautifulSoup
from colorama import Fore, Style

from pipe_cleaner.src.credentials import AccessADO


def get_ticket_urls(ticket_numbers: list) -> list:
    """
    Store in ticket urls for later iteration in Async format.
    """
    return [f'https://azurecsi.visualstudio.com/CSI%20Commodity%20Qualification/_apis/wit/workitems?'
            f'id={ticket}&$expand=all&api-version=5.1' for ticket in ticket_numbers]


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

        try:
            ticket_json: dict = json.loads(raw_ticket_data)

            # For some reason the ticket_json["id"] data type is int
            ticket_id: str = str(ticket_json["id"])
            print(f'\t\t- Collect  |  {Fore.GREEN}Success{Style.RESET_ALL}   |  {ticket_id}')

            # Table Data
            table_data: dict = ticket_json['fields']['System.Description']
            table_data_soup = BeautifulSoup(table_data, 'html.parser')

            table_rows: list = table_data_soup.findAll('tr')
            table_data: dict = get_clean_table_data(table_rows)

            # Table Data Stored
            ticket_data['table_data']: dict = table_data

            # Due Dates
            ticket_data['due_dates']: dict = access_due_dates(ticket_json)

            ticket_data['title'] = ticket_json['fields']['System.Title']

            ticket_data['trr_type']: int = ticket_json.get('fields', {}).get('Custom.TRRType')

            # Assigned To
            ticket_data['assigned_to']: str = ticket_json['fields']['System.AssignedTo']['displayName']

            # State of Qual
            ticket_data['state']: str = ticket_json['fields']['System.State']
            ticket_data["fault_code"]: str = ticket_json.get("fields", {}).get("Custom.QCLFaultcode", "None")

            # if "Blocked" in ticket_data["state"]:
            #     foo = json.dumps(ticket_json, sort_keys=True, indent=4)
            #     print(foo)
            #     input()

            all_tickets_data[ticket_id] = ticket_data

        except json.decoder.JSONDecodeError:
            print(f'\tMicrosoft Azure DevOps did not return all data.')
            print(f'\tError is at TRR {ticket_id}. Recollecting...\n')

        except KeyError:
            print(f'\tNo Description Table... Grabbing from Summary Table Instead in TRR')

            ticket_json: dict = json.loads(raw_ticket_data)

            # For some reason the ticket_json["id"] data type is int
            ticket_id: str = str(ticket_json["id"])
            # print(f'\t\t- Collect  |  {Fore.GREEN}Success{Style.RESET_ALL}   |  {ticket_id}')

            # Table Data
            try:
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


                # Assigned To
                ticket_data['assigned_to']: str = ticket_json['fields']['System.AssignedTo']['displayName']

                # State of Qual
                ticket_data['state']: str = ticket_json['fields']['System.State']


                all_tickets_data[ticket_id] = ticket_data
            except KeyError:
                pass

    all_tickets_data['unique_bios'] = list(set(all_bios))
    all_tickets_data['unique_bmc'] = list(set(all_bmc))
    all_tickets_data['unique_cpld'] = list(set(all_cpld))
    all_tickets_data['unique_os'] = list(set(all_os))

    return all_tickets_data


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


def get_ticket_numbers(csv_file_path: str) -> list:
    """
    Get ADO TRR numbers based off of manual CSV download from ADO.
    """
    with open(csv_file_path) as file:
        file_data = csv.reader(file, delimiter=",", quotechar='"')

        tickets: list = []
        for row in file_data:
            ticket_number: str = row[1]

            if ticket_number.isdigit():
                tickets.append(ticket_number)

        return tickets


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


def get_tickets_data(csv_file_path: str) -> dict:
    """
    Get TRR data based off of ticket numbers manually downloaded from ADO. Below is for reference.

    https://azurecsi.visualstudio.com/CSI%20Commodity%20Qualification/_queries/query/
    5233817c-b790-4482-8cb7-200aae92f508/
    """
    ticket_numbers: list = get_ticket_numbers(csv_file_path)
    ticket_urls: list = get_ticket_urls(ticket_numbers)
    raw_tickets_data: list = asyncio.run(get_ticket_data(ticket_urls))

    return store_tickets_data(raw_tickets_data)


def get_monthly_data(commodity_types: dict):
    month_data: dict = {}
    for pair in commodity_types:
        month = pair[0]
        commodity_type: str = pair[1]

        if month not in month_data:
            month_data[month]: dict = {}
            month_data[month][commodity_type] = 1

        else:
            if commodity_type not in month_data[month]:
                month_data[month][commodity_type] = 1
            else:
                month_data[month][commodity_type] += 1

    return month_data


def get_commodity_types(tickets_data: dict) -> list:
    """

    """
    commodity_types: list = []
    for ticket_number in tickets_data:

        if str(ticket_number).isdigit():
            ticket_data: dict = tickets_data[ticket_number]
            state: str = ticket_data.get("state")
            expected_task_start: str = ticket_data.get("due_dates", {}).get("expected_task_start")

            if "BLOCKED" in state.upper():
                year: str = expected_task_start[0:4]
                month_name: str = get_month_name(expected_task_start, year)
                fault_code: str = ticket_data.get("fault_code")
                assigned_to: str = ticket_data.get("assigned_to")
                print(f'{month_name} | {assigned_to} | {fault_code}')

                # if request_type and expected_task_start:
                #     year: str = expected_task_start[0:4]
                #     month_name: str = get_month_name(expected_task_start, year)
                #     commodity_type: str = request_type.upper().replace(" TEST", "").replace("TEST", "")
                #     month_to_commodity: tuple = (month_name, commodity_type)
                #
                #     commodity_types.append(month_to_commodity)

    return commodity_types


def get_month_name(expected_task_start: str, year: str) -> str:
    """

    """
    month: str = expected_task_start[5:7]
    if month == "01":
        return f"January-{year}"

    elif month == "02":
        return f"February-{year}"

    elif month == "03":
        return f"March-{year}"

    elif month == "04":
        return f"April-{year}"

    elif month == "05":
        return f"May-{year}"

    elif month == "06":
        return f"June-{year}"

    elif month == "07":
        return f"July-{year}"

    elif month == "08":
        return f"August-{year}"

    elif month == "09":
        return f"September-{year}"

    elif month == "10":
        return f"October-{year}"

    elif month == "11":
        return f"November-{year}"

    elif month == "12":
        return f"December-{year}"


def main() -> None:
    """
    Main function to gather information on ADO tickets for commodity types.
    """
    csv_file_path: str = "../../../settings/active_tickets.csv"
    tickets_data: dict = get_tickets_data(csv_file_path)
    commodity_types: list = get_commodity_types(tickets_data)
    monthly_data: dict = get_monthly_data(commodity_types)

    import json
    foo = json.dumps(monthly_data, sort_keys=True, indent=4)
    print(foo)
    input()

main()