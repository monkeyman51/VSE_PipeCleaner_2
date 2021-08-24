import base64
import csv
import os
import re
from json import dumps
from json import loads
from json.decoder import JSONDecodeError

import requests
from bs4 import BeautifulSoup

from pipe_cleaner.src.credentials import AccessADO as Ado
from pipe_cleaner.src.credentials import Path

json_data: dict = {}


def request_ado(trr_id: str):
    """
    Requests data from ADO
    :param trr_id:
    :return:
    """
    # If change, also change check_valid_request function
    base_url = 'https://azurecsi.visualstudio.com/'
    path_url = 'CSI%20Commodity%20Qualification/_apis/wit/workitems?'
    query_parameter = f'id={trr_id}&$expand=all&api-version=5.1'

    user_password = Ado.token_name + ':' + Ado.personal_access_token
    web_address = base_url + path_url + query_parameter
    base64_user_password = base64.b64encode(user_password.encode()).decode()
    headers = {'Authorization': 'Basic %s' % base64_user_password}

    try:
        ado_response = requests.get(
            web_address, headers=headers, timeout=1)
        ado_response.raise_for_status()
        get_csv(trr_id, ado_response)
        get_json(trr_id)
    except requests.exceptions.Timeout:
        print(f'\n  * ADO Response: Timeout Occurred... attempting again\n')
        ado_response = requests.get(
            web_address, headers=headers, timeout=5)
        ado_response.raise_for_status()
        get_csv(trr_id, ado_response)
        get_json(trr_id)
    except requests.exceptions.ConnectionError:
        print(f'\n  * ADO Response: Timeout Occurred... attempting again\n')
        ado_response = requests.get(
            web_address, headers=headers, timeout=5)
        ado_response.raise_for_status()
        get_csv(trr_id, ado_response)
        get_json(trr_id)
    except requests.exceptions.HTTPError:
        print(f'\n  * ADO Response: Timeout Occurred... attempting again\n')
        ado_response = requests.get(
            web_address, headers=headers, timeout=5)
        ado_response.raise_for_status()
        get_csv(trr_id, ado_response)
        get_json(trr_id)


def request_ado_json(ticket_number: str) -> dict:
    """
    Get ADO JSON
    :param ticket_number: TRR
    :return:
    """
    base_url = 'https://azurecsi.visualstudio.com/'
    path_url = 'CSI%20Commodity%20Qualification/_apis/wit/workitems?'
    query_parameter = f'id={ticket_number}&$expand=all&api-version=5.1'

    user_password = f'{Ado.token_name}:{Ado.personal_access_token}'
    web_address = base_url + path_url + query_parameter
    base64_user_password = base64.b64encode(user_password.encode()).decode()
    headers = {'Authorization': 'Basic %s' % base64_user_password}

    file_path = f'pipe_cleaner/data/ticket_{ticket_number}.json'
    time_out = '\tADO Response: Timeout Occur... attempting again\n'

    try:
        ado_response = requests.get(
            web_address, headers=headers, timeout=1)

        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(ado_response.text)

        with open(file_path, 'r') as f:
            json_file = loads(f.read())

        return json_file

    except requests.exceptions.Timeout:
        print(time_out)
        ado_response = requests.get(
            web_address, headers=headers, timeout=1)

        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(ado_response.text)

        with open(file_path, 'r') as f:
            json_file = loads(f.read())

        return json_file

    except requests.exceptions.ConnectionError:
        print(time_out)
        ado_response = requests.get(
            web_address, headers=headers, timeout=1)

        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(ado_response.text)

        with open(file_path, 'r') as f:
            json_file = loads(f.read())

        return json_file

    except requests.exceptions.HTTPError:
        print(time_out)
        ado_response = requests.get(
            web_address, headers=headers, timeout=1)

        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(ado_response.text)

        with open(file_path, 'r') as f:
            json_file = loads(f.read())

        return json_file


def get_dhcp():
    """
    Get Dynamic Host Configuration Protocol
    :return:
    """
    data = {
        'action': 'get_reservations',
    }
    response = requests.post(url=f'http://172.30.1.100/console/console_js.php', json=data)
    return loads(response.text)


def check_valid_request(ticket_number: str):
    """
    Checks whether valid TRR in ADO.
    :param ticket_number: TRR ID
    :return:
    """
    base_url = 'https://azurecsi.visualstudio.com/'
    path_url = 'CSI%20Commodity%20Qualification/_apis/wit/workitems?'
    query_parameter = f'id={ticket_number}&$expand=all&api-version=5.1'

    user_password = f'{Ado.token_name}:{Ado.personal_access_token}'
    web_address = base_url + path_url + query_parameter
    base64_user_password = base64.b64encode(user_password.encode()).decode()
    headers = {'Authorization': 'Basic %s' % base64_user_password}

    try:
        ado_response = requests.get(
            web_address, headers=headers, timeout=1)
        return ado_response.status_code

    except requests.exceptions.Timeout:
        print('  ADO Response: Timeout Occur... attempting again')
        ado_response = requests.get(
            web_address, headers=headers, timeout=5)
        return ado_response.status_code

    except requests.exceptions.ConnectionError:
        print(f'\n  * ADO Response: Timeout Occurred... attempting again\n')
        ado_response = requests.get(
            web_address, headers=headers, timeout=5)
        return ado_response.status_code

    except requests.exceptions.HTTPError:
        print(f'\n  * ADO Response: Timeout Occurred... attempting again\n')
        ado_response = requests.get(
            web_address, headers=headers, timeout=5)
        return ado_response.status_code


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

    # Needs a functional error handle for wrong Host ID input, previous attempts haven't worked
    try:
        requests.post(url='http://172.30.1.100/console/console_js.php', json=generate_data)
    except JSONDecodeError:
        print('Need to be figured out.')

    host_response = requests.post(url='http://172.30.1.100/console/console_js.php', json=host_name_data)

    product = loads(str(host_response.text))['host_name_data']['product']
    serial = loads(str(host_response.text))['host_name_data']['serial']

    return f'{product}-{serial}'


def get_console_server_json(product_serial: str, host_id: str) -> object:
    """
    Gets the Generated data using the product_serial string and creates JSON file.
    Returns JSON of Host within Console Server.

    :param product_serial: string from generate_json_data method
    :param host_id: 24-character string host id in URL of Host Details in Console Server
    :return: JSON data of Host
    """
    clean_host_id = host_id.replace('/', '')

    data = {
        'action': 'get_json_data',
        'host_id': f'{product_serial}'
    }
    response = requests.post(url=f'http://172.30.1.100/results/{product_serial}.json', json=data)

    with open(f'{Path.info}/{clean_host_id}.json', 'w') as f:
        f.write(response.text)

    return response


def get_all_hosts_console_server():
    """
    All Hosts in Console Server

    :return: JSON data of Host
    """
    data = {
        'action': 'get_host_status',
    }
    response = requests.post(url=f'http://172.30.1.100/console/console_js.php', json=data)

    return loads(response.text)


def write_csv_table(ticket_id: str, ado_response: requests.models.Response, table_name: str, table_number: int):
    """
    Write CSV table for JSON files for later extraction
    :param ticket_id: TRR number
    :param ado_response: response from Requests Module
    :param table_name:
    :param table_number:
    :return:
    """

    with open(f'{Path.info}{ticket_id}/{table_name}.csv', 'w', newline='', encoding='utf-8') as f:
        soup = BeautifulSoup(ado_response.text, 'html.parser')
        table = soup.findAll('table')[table_number]
        rows = table.findAll('tr')
        writer = csv.writer(f)
        for row in rows:
            csv_row = []
            for cell in row.findAll(['td', 'th']):
                csv_row.append(cell.get_text())
                cell = []
                for i in csv_row:
                    cell_i = str(re.sub(r"\\n {2}", '', i))
                    cell.append(cell_i)
                row = []
                for i in cell:
                    row.append(re.sub(r'\xa0', '', i))
            writer.writerow(row)


def get_csv(trr_id: str, ado_response: object.__text_signature__):
    try:
        os.mkdir(f'{Path.info}/{trr_id}')
    except FileExistsError:
        pass

    try:
        write_csv_table(trr_id, ado_response, 'table_1', 0)
        write_csv_table(trr_id, ado_response, 'table_2', 1)
        write_csv_table(trr_id, ado_response, 'table_3', 2)
        write_csv_table(trr_id, ado_response, 'table_4', 3)

    except IndexError:
        pass


def replace_csv_characters(csv_reader_obj):
    try:
        for key, value in csv_reader_obj:
            parsed_key = replace_characters(str(key))
            parsed_value = replace_characters(str(value))

            json_data.update({parsed_key: parsed_value})

    except ValueError:
        pass


# noinspection PyGlobalUndefined
def create_path_csv(trr_id: str, table_name: str):
    global csv_reader
    try:
        csv_reader = csv.reader(open(f'{Path.info}{trr_id}/{table_name}.csv', 'r'))
    except FileNotFoundError:
        pass
    return csv_reader


def get_json(trr_id: str):
    json_data.clear()
    csv_reader_1 = create_path_csv(trr_id, 'table_1')
    csv_reader_2 = create_path_csv(trr_id, 'table_2')
    csv_reader_3 = create_path_csv(trr_id, 'table_3')
    csv_reader_4 = create_path_csv(trr_id, 'table_4')

    replace_csv_characters(csv_reader_1)
    replace_csv_characters(csv_reader_2)
    replace_csv_characters(csv_reader_3)
    replace_csv_characters(csv_reader_4)

    with open(f'{Path.info}{trr_id}/final.json', 'w') as f:
        f.write(dumps(json_data, indent=4))


def write_host_groups_json() -> dict:
    """
    Writes all Host Groups information from Console Server into a JSON file for later extraction.
    :return: None
    """
    generate_data = {
        'action': 'get_host_groups'
    }
    response = requests.post(url='http://172.30.1.100/console/console_js.php', data=dumps(generate_data))

    return loads(response.text)


def get_host_groups_json():
    """
    Direct gets the Host Groups data from Console Server in JSON.
    :return:
    """
    generate_data = {
        'action': 'get_host_groups'
    }
    response = requests.post(url='http://172.30.1.100/console/console_js.php', data=dumps(generate_data))

    return loads(response.text)


def get_reservations(json_file_name) -> list:
    """
    Writes all Host Groups information from Console Server into a JSON file for later extraction.
    :param json_file_name: Name of JSON file name
    :return: List of Host Groups JSON
    """
    generate_data = {
        'action': 'get_reservations'
    }
    response = requests.post(url='http://172.30.1.100/console/console_js.php', data=dumps(generate_data))

    with open(json_file_name, 'w') as f:
        f.write(response.text)

    with open(json_file_name, 'r') as f:
        host_groups_json = loads(f.read())

    return host_groups_json


def get_all_pipe_names(json_file: dict) -> list:
    pipe_all_host_ids = []

    initial = 0
    while initial < 10_000:
        try:
            pipe_all_host_ids.append(json_file['host_groups'][initial]['name'])
            initial += 1
        except IndexError:
            break
        except TypeError:
            break
    return pipe_all_host_ids


def host_group_id_with_name(json_file: dict, host_group_name: str) -> str:
    """
    Get Host Group ID from Console Server using the JSON File containing all Host Group Names (aka Pipes),
    To get the Host Group ID for extracting that Pipe information.
    :param json_file:
    :param host_group_name:
    :return:
    """
    host_group_id = []

    initial = 0
    while initial < 10_000:
        try:
            if host_group_name in json_file['host_groups'][initial]['name']:
                host_group_id.append(json_file['host_groups'][initial]['id'])
                break
            initial += 1
        except IndexError:
            break

    return host_group_id[0]


def get_all_host_groups(json_file: dict) -> list:
    all_host_groups = []

    initial = 0
    while initial < 10_000:
        try:
            all_host_groups.append(json_file['host_groups'][initial]['host_ids'])
            initial += 1
        except IndexError:
            break
    return all_host_groups


def get_all_descriptions(json_file: dict) -> list:
    """
    Get Descriptions from Host Groups page. Used for additional information.
    :param json_file:
    :return:
    """
    all_host_groups = []

    initial = 0
    while initial < 10_000:
        try:
            all_host_groups.append(json_file['host_groups'][initial]['description'])
            initial += 1
        except IndexError:
            break
    return all_host_groups


def machine_name_to_ticket() -> dict:
    """
    Create dictionary from Machine Name to Ticket TRR
    :return:
    """
    name_to_ticket: dict = {}

    return name_to_ticket


def gen_num_crd_repository(target: str) -> int:
    """
    Based on the Target Configuration pulled from the TRR number, gets the Gen number for CRD folder
    :param target: Target from TRR
    :return: Gen Number (5, 6, 7, 8, 9, 10)
    """
    upper_target_configuration = str(target).upper()
    upper_target_configuration.replace(' ', '')
    if 'GEN5.' in upper_target_configuration:
        return 5
    elif 'GEN6.' in upper_target_configuration:
        return 6
    elif 'GEN7.' in upper_target_configuration:
        return 7
    elif 'GEN8.' in upper_target_configuration:
        return 8
    elif 'GEN9.' in upper_target_configuration:
        return 9
    elif 'Gen10.' in upper_target_configuration:
        return 10


def target_configuration_gen_num(target_configuration: str) -> str:
    """
    Based on the Target Configuration from TRR, get precise Gen Number for CRD file
    :param target_configuration: Target Configuration from TRR
    :return: Gen Number (5.x, 6.x, 7.x, 8.x, 9.x, 10.x)
    """
    upper_target_configuration = str(target_configuration).upper()
    target = upper_target_configuration.replace(' ', '')

    gen_num = 5
    while gen_num < 11:
        gen_num_str = str(gen_num)
        if gen_num_str in target:
            gen_specific = 0
            while gen_specific < 11:
                gen_specific_str = str(gen_specific)
                gen_specific_dot = f'.{gen_specific_str}'
                if gen_specific_dot in target:
                    whole_gen = f'GEN{gen_num_str}{gen_specific_dot}'
                    return whole_gen
                gen_specific += 1
        gen_num += 1


def get_all_host_ids(host_group_id: str, host_group_name: str) -> list:
    """
    Get all host ids based on host group ID
    :param host_group_name:
    :param host_group_id:
    :return:
    """
    url = 'http://172.30.1.100/console/console_js.php'
    data = {
        'action': 'get_host_group_host_list',
        'host_group_id': f'{host_group_id}'
    }
    response = requests.post(url=url, data=dumps(data))

    with open(f'pipe_cleaner/data/{host_group_name}.json', 'w') as f:
        f.write(response.text)

    with open(f'pipe_cleaner/data/{host_group_name}.json', 'r') as f:
        json_file = loads(f.read())

    return json_file


def request_json_from_ado(ticket_number: str):
    """
    Requests data from ADO. Creates JSON to extract data.
    :param ticket_number: TRR Number
    :return: JSON File
    """
    base_url = 'https://azurecsi.visualstudio.com/'
    path_url = 'CSI%20Commodity%20Qualification/_apis/wit/workitems?'
    query_parameter = f'id={ticket_number}&$expand=all&api-version=5.1'

    user_password = f'{Ado.token_name}:{Ado.personal_access_token}'
    web_address = f'{base_url}{path_url}{query_parameter}'
    base64_user_password = base64.b64encode(user_password.encode()).decode()
    headers = {'Authorization': 'Basic %s' % base64_user_password}

    try:
        ado_response = requests.get(
            web_address, headers=headers, timeout=1)
        return ado_response
    except requests.exceptions.Timeout:
        print(f'\t* ADO Response: Timeout Occurred... attempting again')
        requests.get(web_address, headers=headers, timeout=5)

    except requests.exceptions.ConnectionError:
        print(f'\t* ADO Response: Timeout Occurred... attempting again')
        requests.get(web_address, headers=headers, timeout=5)

    except requests.exceptions.HTTPError:
        print(f'\t* ADO Response: Timeout Occurred... attempting again')
        requests.get(web_address, headers=headers, timeout=5)


def replace_characters(csv_item: str) -> str:
    """
    Replace unnecessary characters from keys and values from TRR Tables
    :param csv_item: key (left) or value (right)  side of TRR (Test Run Request) table column in ADO (Azure DevOps)
    """
    csv_item.replace('\n  ', '')
    csv_item.replace('\\n  ', '')
    csv_item.replace('aaa10302', '')
    csv_item.replace('\xa0', '')
    csv_item.replace('\u00a0', '')
    csv_item.replace('\u00c2', '')
    csv_item.replace('\u00c2', '')

    return csv_item
