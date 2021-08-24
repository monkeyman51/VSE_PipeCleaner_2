"""
Access ADO's Task
"""
import asyncio
import sys
from base64 import b64encode
from json import loads

import requests
from aiohttp import ClientSession, client_exceptions
from bs4 import BeautifulSoup

from pipe_cleaner.src.credentials import AccessADO as Ado


def get_response_text_from_ado() -> str:
    """
    Get content from all recent ADO work items for grabbing TRR IDs. This is an effort towards automating part number
    library.
    """
    # site_url: str = 'https://azurecsi.visualstudio.com/CSI%20Commodity%20Qualification/_workitems/recentlycreated/'
    # site_url: str = 'https://azurecsi.visualstudio.com/CSI%20Commodity%20Qualification/_workitems/recentlyupdated/'
    # site_url: str = 'https://azurecsi.visualstudio.com/CSI%20Commodity%20Qualification/_queries/query/' \
    #                 '5233817c-b790-4482-8cb7-200aae92f508/'
    # site_url: str = 'https://azurecsi.visualstudio.com/CSI%20Commodity%20Qualification/_queries/query/' \
    #                 '3eddd432-529e-499d-a7d1-b4e639e95754/'
    site_url: str = "https://azurecsi.visualstudio.com/CSI%20Commodity%20Qualification/_queries/query/" \
                    "5233817c-b790-4482-8cb7-200aae92f508/"
    user_password: str = f'{Ado.token_name}:{Ado.personal_access_token}'

    headers = {'Authorization': f'Basic {b64encode(user_password.encode()).decode()}'}

    # return requests.get(site_url, headers=headers).text
    return requests.get(site_url, headers=headers).text


def find_work_items_from_response(response_text: str) -> dict:
    """
    Find the work items from HTTP text response.
    """
    soup = BeautifulSoup(response_text, 'html.parser')
    data = str(soup.findAll('script', type='application/json')).replace('</script>', ''). \
        replace('<script id="dataProviders" type="application/json">', '')

    print(soup)
    input()

    json_data = loads(data)

    import json
    foo = json.dumps(json_data, sort_keys=True, indent=4)
    print(foo)
    input()

    return loads(data)[0]['data']['ms.vss-work-web.new-work-items-hub-recentlycreated-tab-data-provider']['fieldValues']


def collect_test_requests_from_work_items(work_items: dict) -> list:
    """
    Collect only TRR IDs from work items that are for Test Run Requests
    """
    actual_test_requests: list = []

    count: int = 0
    for work_item in work_items:
        count += 1
        trr_id: str = work_item['data'][0]
        print(trr_id)
        work_item_name: str = str(work_item['data'][1]).upper()
        # print(f'work_item_name: {work_item}')

        if 'TEST RUN REQUEST' in work_item_name:
            content: dict = work_item['data'][0]

            # import json
            # foo = json.dumps(content, sort_keys=True, indent=4)
            # print(foo)
            # input()

            actual_test_requests.append(content)

    print(f'count: {count}')

    return actual_test_requests


def get_trr_urls(trr_ids: list) -> list:
    """
    Store the trr id in a url format for easier fetching of data.
    """
    return [f'https://azurecsi.visualstudio.com/CSI%20Commodity%20Qualification/_apis/wit/workitems?'
            f'id={ticket}&$expand=all&api-version=5.1' for ticket in trr_ids]


async def fetch_site(url: str, headers: dict) -> str:
    """
    Grabs the information from Azure Devops per session depending on how many tickets in the form of URls
    """
    async with ClientSession(headers=headers) as session:
        try:
            async with session.get(url) as response:
                await asyncio.sleep(0.5)
                ticket_data = await response.text()

        # Rare occurrence dealing with Async,
        except client_exceptions.ClientOSError:
            print(f'\t[WinError 10054] An existing connection was forcibly closed '
                  f'by the remote host')
            print(f'\tPress ENTER to exit Pipe Cleaner...', end='')
            input()
            sys.exit()

    return ticket_data


async def get_ticket_data(ticket_urls):
    """
    Creates tasks for executing the event loop. Tasks are just requests sent quantified by number of unique tickets
    found in the Console Server
    """
    user_password: str = f'{Ado.token_name}:{Ado.personal_access_token}'
    base64_user_password: str = b64encode(user_password.encode()).decode()
    headers: dict = {'Authorization': f'Basic {base64_user_password}'}

    tasks: list = [asyncio.create_task(fetch_site(request, headers)) for request in ticket_urls]

    return await asyncio.gather(*tasks)


def clean_html_text(field_data) -> str:
    """
    Clean up HTML oriented extra stuff to just get the actual text within the field.
    """
    return str(field_data.text). \
        replace('\n', ''). \
        replace('  ', ' '). \
        replace('\xa0', ''). \
        replace('\u200b', ''). \
        replace('\u2013', ''). \
        strip()


def clean_component_value(component_value: str) -> str:
    """
    Cleans component to make easier to call value through key later.
    """
    return component_value. \
        upper(). \
        replace(' - ', ' '). \
        replace('N/A', 'None')


def get_clean_table_data(table_rows: list) -> dict:
    """
    Gather data from description table given a TRR within ADO
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
            raw_component_key: str = table_row_data[0].upper()
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

        except IndexError:
            pass

    all_table_data['all_part_numbers']: list = all_part_numbers

    return all_table_data


def store_part_numbers_data(raw_tickets_data: list) -> dict:
    """

    """
    all_part_numbers: dict = {}

    print(f'\n\t=====================================================================')
    print(f'\t  Tickets (TRRs) - Collecting and Processing Data')
    print(f'\t=====================================================================')
    print(f'\t\t  STATUS   |  REASON    |  TICKET')

    for raw_ticket_data in raw_tickets_data:
        ticket_json: dict = loads(raw_ticket_data)
        ticket_id = str(ticket_json["id"])
        print(f'\t\t- Collect  |  Success   |  {ticket_id}')

        ticket_data = {'table_data': get_table_data(ticket_json),
                       'title': ticket_json['fields']['System.Title'],
                       'state': ticket_json['fields']['System.State']}

        # import json
        # foo = json.dumps(ticket_data, sort_keys=True, indent=4)
        # print(foo)
        # input()

    return all_part_numbers


def get_table_data(ticket_json: dict) -> dict:
    """
    Get table in key-value pair
    """
    table_data_soup = BeautifulSoup(ticket_json['fields']['System.Description'], 'html.parser')
    table_data = table_data_soup.findAll('tr')

    return get_clean_table_data(table_data)


def request_ado() -> None:
    """
    Requests data from ADO
    """
    response_text: str = get_response_text_from_ado()
    print(response_text)
    input()
    work_items: dict = find_work_items_from_response(response_text)
    trr_ids: list = collect_test_requests_from_work_items(work_items)
    # print(f'{len(trr_ids)}')

    trr_urls: list = get_trr_urls(trr_ids)
    raw_tickets_data: list = asyncio.run(get_ticket_data(trr_urls))
    store_part_numbers_data(raw_tickets_data)


request_ado()
