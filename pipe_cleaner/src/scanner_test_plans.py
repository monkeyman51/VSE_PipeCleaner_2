import base64
import os
from json import loads

import requests
from bs4 import BeautifulSoup
from colorama import Fore, Style

from pipe_cleaner.src.credentials import AccessADO as Ado
from pipe_cleaner.src.data_access import request_json_from_ado
from pipe_cleaner.src.terminal_properties import terminal_header_section


def check_repositories(username_info: dict):
    """
    Checks to see in Z:Drive whether user's repository for Qual preparation is set
    :param username_info:
    :return:
    """
    z_drive_users: str = r'Z:\Kirkland_Lab\PipeCleaner_Users'

    try:
        os.mkdir(fr'{z_drive_users}\{username_info["default_name"]}')
    except FileExistsError:
        pass

    return z_drive_users


def check_user_paths(user_path: str):
    """
    Checks that the Pipe Cleaner in the Z: Drive is working or existent
    :param user_path:
    :return:
    """
    try:
        os.mkdir(user_path)
    except FileExistsError:
        pass


def json_from_ado(file_name: str, ado_response: requests.request) -> dict:
    """
    Create JSON file from ADO Response. Creates reading file to read JSON file.
    :param file_name:
    :param ado_response:
    :return:
    """
    with open(f'pipe_cleaner/src/{file_name}.json', 'w') as f:
        f.write(ado_response.text)

    with open(f'pipe_cleaner/src/{file_name}.json', 'r') as f:
        json_file = loads(f.read())

    return json_file


def html_from_ado(file_name, ado_response):
    """
    Create JSON file from ADO Response.
    :param file_name:
    :param ado_response:
    :return:
    """
    with open(f'pipe_cleaner/src/{file_name}.html', 'w', encoding='utf-8') as f:
        f.write(ado_response.text)

    with open(f'pipe_cleaner/src/{file_name}.html', 'r') as f:
        html_file = f.read()

    return html_file


def check_broken_hyper_link(ticket_number: str, hyper_link: str):
    """
    Sometimes Test plan could not be found. Warns User that test plan does not exist.
    :param ticket_number:
    :param hyper_link:
    :return:
    """
    user_password = Ado.token_name + ':' + Ado.personal_access_token
    web_address = hyper_link
    base64_user_password = base64.b64encode(user_password.encode()).decode()
    headers = {'Authorization': 'Basic %s' % base64_user_password}

    try:
        ado_response = requests.get(
            web_address, headers=headers, timeout=1)
        ado_json = html_from_ado(f'{ticket_number}-test_plan', ado_response)
        return ado_json

    except requests.exceptions.Timeout:
        print(f'\n  * ADO Response: Timeout Occurred... attempting again\n')
        ado_response = requests.get(
            web_address, headers=headers, timeout=5)
        ado_json = html_from_ado(f'{ticket_number}-test_plan', ado_response)
        return ado_json

    except requests.exceptions.ConnectionError:
        print(f'\n  * ADO Response: Timeout Occurred... attempting again\n')
        ado_response = requests.get(
            web_address, headers=headers, timeout=5)
        ado_json = html_from_ado(f'{ticket_number}-test_plan', ado_response)
        return ado_json

    except requests.exceptions.HTTPError:
        print(f'\n  * ADO Response: Timeout Occurred... attempting again\n')
        ado_response = requests.get(
            web_address, headers=headers, timeout=5)
        ado_json = html_from_ado(f'{ticket_number}-test_plan', ado_response)
        return ado_json


def continue_ado_json(test_case):
    """
    Depending on the TRR, the structure of the JSON file might be different. Unfortunately, need to account for that.
    :return:
    """
    try:
        return test_case['testCases']
    except KeyError:
        return test_case['testPoints']


def get_test_cases(ticket_number: str, json_file: dict) -> list:
    """
    Check if there are more than 3 files inside the Attachment part of the TRR. It's for PM Review.
    Also, validates for those 3 files. ie. CRD, PDF, SKUDOC
    :param json_file: JSON file from TRR
    :param ticket_number: TRR Number
    :return: name_to_url, name_to_gen
    """
    # test_cases: list = []

    for item in json_file['relations']:
        if item['rel'] == 'Hyperlink':
            file_url = item['url']
            link_file = check_broken_hyper_link(ticket_number, file_url)

            soup = BeautifulSoup(link_file, 'html.parser')
            test_case_json = loads(soup.find('script', type='application/json').contents[0])

            test_case = test_case_json['data']['ms.vss-test-web.test-plans-hub-refresh-data-provider']
            info = continue_ado_json(test_case)
            return info

    # return test_cases[0]


def get_test_suites(request_type: str, ticket: str, toolkit_version: str):
    """
    Get correct test suite from Z: Drive based on request type
    :return:
    """

    test_suites = {'DIMM': r'Z:\Kirkland_Lab\Microsoft_CSI\Tools\Toolkit_Releases'
                           fr'\Veritas\{toolkit_version}\TestController\TestSuites\Dimm.ps1',
                   'HDD': r'Z:\Kirkland_Lab\Microsoft_CSI\Tools\Toolkit_Releases'
                          fr'\Veritas\{toolkit_version}\TestController\TestSuites\Dimm.ps1',
                   'NVME': r'Z:\Kirkland_Lab\Microsoft_CSI\Tools\Toolkit_Releases'
                           fr'\Veritas\{toolkit_version}\TestController\TestSuites\Dimm.ps1',
                   'SATA': r'Z:\Kirkland_Lab\Microsoft_CSI\Tools\Toolkit_Releases'
                           fr'\Veritas\{toolkit_version}\TestController\TestSuites\Dimm.ps1'}

    parsed_type = request_type.upper()

    test_suite_path = test_suites.get(parsed_type)

    if test_suite_path is None:

        # Warns if no test suite found
        print(f'\t{Fore.RED}Did not find test suite in Z: Drive{Style.RESET_ALL} based on {ticket} '
              f'Request Type({request_type})')
        print(f'\t\t- Current pull from {toolkit_version} in Z:Drive')

        for test in test_suites:
            print(f'\t\t\t* {test}')
    else:
        return test_suite_path


def get_test_numbers(test_cases: list):
    """
    Get Test Cases based on 5 digit numbers in the front to parse for PowerShell Test Suites later
    :param test_cases:
    :return:
    """
    test_numbers: list = []

    for case in test_cases:
        raw_case = str(case['workItem']['name'])
        test_numbers.append(raw_case.replace('[', '')[0:5])

    return test_numbers


def create_power_shell(user_ticket_path: str, test_suite_path: str, test_numbers: list, request_type: str):
    """
    Create PowerShell Script
    :return:
    """
    print(f'Test Numbers: {test_numbers}')
    # Had to use 2 Context Managers because information will only write once with context manager for some reason
    # with open(test_suite_path, 'r') as file:
    #     data_file = file.read()
    #     for line in data_file:
    #         print(line)
    #
    # with open(fr'{user_ticket_path}\{request_type}-TestSuite.ps1', 'w') as file:
    #     # server_list = file_data_1.replace('# {{insert_server_list}}', insert_server_list)
    #     file.write(data_file)


def start_test_plans(username_info: dict, unique_tickets: list, name_to_path: dict):
    """

    :return:
    """
    current_version_toolkit: str = name_to_path.get('toolkit_version')
    users_path = name_to_path.get('users')
    default_name = username_info.get('default_name')

    terminal_header_section('Extract TRR Test Plans *** BETA ***', 'Manipulate PowerShell Scripts for Qual')
    print(f"\t{Fore.GREEN}Extracting Test Plans{Style.RESET_ALL} from TRRs...")
    print(f'\t\t- Using Toolkit {Fore.GREEN}{current_version_toolkit}{Style.RESET_ALL} in Z: Drive\n')

    print(f'Length of Unique Tickets: {len(unique_tickets)}')
    print(f'Content of Unique Tickets: {unique_tickets}')

    for ticket_number in unique_tickets:
        # Gather Data
        ado_json_response: requests.request = request_json_from_ado(ticket_number)
        azure_ticket_json: dict = json_from_ado('link_files', ado_json_response)
        request_type = azure_ticket_json['fields']['Custom.QTComponentCategory']
        test_cases: list = get_test_cases(ticket_number, azure_ticket_json)
        test_numbers: list = get_test_numbers(test_cases)
        print(f'Test Numbers: {test_numbers}')

        # Setup Paths
        test_suite_path = get_test_suites(request_type, ticket_number, current_version_toolkit)
        user_ticket_path = fr'{users_path}\{default_name}\{ticket_number}'
        check_user_paths(user_ticket_path)

        # Create PowerShell
        create_power_shell(user_ticket_path, test_suite_path, test_numbers, request_type)

    check_repositories(username_info)
