"""
Using asynchronous programming to fetch data from Azure DevOps including Ticket table data, due dates, and state of the
ticket information.
"""

import asyncio
import base64
import subprocess
import json
import os
import sys

import requests
from aiohttp import ClientSession, client_exceptions
from bs4 import BeautifulSoup
from colorama import Fore, Style

from pipe_cleaner.src.credentials import AccessADO
from pipe_cleaner.src.data_access import request_ado_json
from pipe_cleaner.src.credentials import AccessADO as Ado


def exception_check_reference_test_plans(potential_test_plans: str):
    """
    Ensure that if an initial reference test plan from the TRR show empty, the next Table Row shows the in
    :param potential_test_plans:
    :return:
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


def find_other_potential_components(component_key, index, table_rows):
    """
    Sometimes in the ADO website, table rows and their data do not align. Therefore, when calling the key,
    the value sometimes does not show up due to weird configuration of the ticket table. To ensure that
    :param component_key:
    :param index:
    :param table_rows:
    :return:
    """
    potential_reference_test_plans = []

    if 'reference_test' in component_key:
        potential_reference_test_plans_index = index + 1
        field_reference_test = table_rows[potential_reference_test_plans_index].find(['td'])
        potential = str(field_reference_test.text).replace('\n', '').replace('  ', ' ').replace('\xa0', '')
        potential_exception: str = exception_check_reference_test_plans(potential)
        potential_reference_test_plans.append(potential_exception)


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


def clean_html_text(field_data):
    clean_text = str(field_data.text). \
        replace('\n', ''). \
        replace('  ', ' '). \
        replace('\xa0', ''). \
        replace('\u200b', '').\
        replace('\u2013', '')
    return clean_text.strip()


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


def check_attachment_folders(document_paths: dict, processor_generations: tuple) -> bool:
    """
    Creates directories for CRD and SKUDOC in Z:Drive. Pass if already created.
    Ensures repositories for CRD/SKUDOC.
    """
    crd_path: str = document_paths.get('crd_path')
    skudoc_path: str = document_paths.get('skudoc_path')

    try:
        os.mkdir(crd_path)
        os.mkdir(skudoc_path)

    except FileExistsError:
        pass

    except OSError:
        print(f'\n\t{Fore.RED}Shared Drive:{Style.RESET_ALL} Might not have access to the Z: Drive.')
        print(f'\t\t- Please enter personal credentials for Z: Drive access first before using Pipe Cleaner.')
        print(f'\n\t{Fore.LIGHTBLUE_EX}PRESS ENTER{Style.RESET_ALL} to exit program...')
        input()
        sys.exit()
        # return False

    try:
        for generation in processor_generations:
            os.mkdir(fr'{crd_path}\{generation}')
            os.mkdir(fr'{skudoc_path}\{generation}')

    except FileExistsError:
        pass

    return True


def get_generation_from_name(file_name: str) -> str:
    """
    Checks Generation Number based on file name from Attachment Section in TRR
    For key-value pair in dictionary ie. File Name to Processor Generation
    :param file_name:
    :return: Gen number
    """
    initial = 4  # Generation to start
    while initial < 10:  # Gen 10 not coming out soon - 9/26/2020
        if f'GEN{str(initial)}' in file_name:
            return f'Gen_{str(initial)}.x'
        initial += 1


def request_attachment_file(file_url, headers, file_path, generation, file_name, ticket_id) -> str:
    """
    Requests attachment file from ticket (TRR), the response is downloaded
    :param file_url:
    :param headers:
    :param file_path:
    :param generation:
    :param file_name:
    :param ticket_id:
    :return:
    """
    ado_response = requests.get(file_url, headers=headers, timeout=1)

    try:
        with open(fr'{file_path}\{generation}\{file_name}', 'wb') as attachment_file:
            if ado_response.content == 'null':
                print(f'{Fore.RED}NULL{Style.RESET_ALL}')
            else:
                attachment_file.write(ado_response.content)

            # Terminal progress report
            print(f'\t\t- Stored   |  {Fore.GREEN}New File{Style.RESET_ALL}  |  {ticket_id}')
            # Returns the drive file path
            return fr'{file_path}\{generation}\{file_name}'
    except OSError:
        pass


def download_attachment_file_from_ticket(ticket_json: dict, file_name: str, file_path: str, ticket_id: str):
    """
    
    :param ticket_json: TRR data for accessing and downloading attachment file
    :param file_name:
    :param file_path:
    :param ticket_id:
    :return:
    """
    for item in ticket_json['relations']:
        if item['rel'] in 'AttachedFile':

            # Get information from JSON file
            current_attached_file_name: str = item['attributes']['name']

            if file_name.upper() in current_attached_file_name.upper():
                file_url: str = item['url']
                generation: str = get_generation_from_name(file_name)

                user_password = f'{Ado.token_name}:{Ado.personal_access_token}'
                base64_user_password = base64.b64encode(user_password.encode()).decode()
                headers = {'Authorization': 'Basic %s' % base64_user_password}
                # headers = {'Authorization': f'Basic {base64_user_password}'}

                try:
                    return request_attachment_file(file_url, headers, file_path, generation, file_name, ticket_id)

                # Below are known exceptions, should not happen often
                except requests.exceptions.Timeout:
                    print(f'\n  * ADO Response: {Fore.RED}Timeout Occurred{Style.RESET_ALL}... attempting again\n')
                    return request_attachment_file(file_url, headers, file_path, generation, file_name, ticket_id)
                except requests.exceptions.ConnectionError:
                    print(f'\n  * ADO Response: {Fore.RED}Timeout Occurred{Style.RESET_ALL}... attempting again\n')
                    return request_attachment_file(file_url, headers, file_path, generation, file_name, ticket_id)
                except requests.exceptions.HTTPError:
                    print(f'\n  * ADO Response: {Fore.RED}Timeout Occurred{Style.RESET_ALL}... attempting again\n')
                    return request_attachment_file(file_url, headers, file_path, generation, file_name, ticket_id)


def get_drive_path(document_type: str, file_name: str, attached_file_name_paths: dict) -> str:
    """

    :param document_type: CRD, SKUDOC
    :param attached_file_name_paths:
    :param file_name:
    :return:
    """
    # import json
    # print(json.dumps(attached_file_name_paths, sort_keys=True, indent=4))
    # input()
    if document_type.upper() == 'CRD':
        crd_base_path: str = attached_file_name_paths.get('base_paths', {}).get('crd_path', 'None')
        generation: str = get_generation_from_name(file_name)
        return fr'{crd_base_path}\{generation}\{file_name}'

    elif document_type.upper() == 'SKUDOC':
        skudoc_base_path: str = attached_file_name_paths.get('base_paths', {}).get('skudoc_path', 'None')
        generation: str = get_generation_from_name(file_name)
        return fr'{skudoc_base_path}\{generation}\{file_name}'


def check_skudoc_file_name(attachment_file_name: str) -> bool:
    """
    Ensure that there is a SKUDOC from the TRR
    :param attachment_file_name:
    :return:
    """
    attachment_file_name = attachment_file_name.upper()

    if attachment_file_name[0] == 'M' and attachment_file_name[1].isdigit() and '.XLSX' in attachment_file_name:
        return True
    else:
        return False


def check_crd_file_name(attachment_file_name: str) -> bool:
    """
    Make sure attachment file name from TRR is a CRD
    :param attachment_file_name:
    :return:
    """
    attachment_file_name = attachment_file_name.upper()

    if 'CRD' in attachment_file_name and attachment_file_name[0] == 'M' and '.XLSX' in attachment_file_name:
        return True
    else:
        return False


def get_crd_drive_file_path(file_names_from_drive_paths: list, attachment_file_name: str, ticket_json: dict,
                            ticket_id: str, attached_file_name_paths: dict) -> str:
    """
    Return the drive file path for later hyperlink in the excel output

    :param file_names_from_drive_paths: list of current unique document file names ie. CRD, SKUDOC, Other, Data
    :param attachment_file_name: current file name from ticket ie. CRD, SKUDOC, Other, Data
    :param ticket_json: contains data for current ticket ie. TRR number
    :param ticket_id: TRR number, used for later accessing and downloading from ADO if necessary
    :param attached_file_name_paths: file names already stored in the shared drive
    :return: shared drive file path for hyperlink later in the excel output
    """
    for file_path in file_names_from_drive_paths:
        # 80 characters is the current number of characters that include the base file path plus generation folder
        # not ideal for scalability, getting length of file base path is not do-able since Python has difficulty reading
        # back slashes ie. \
        # Might be better solution out there as file path length can change
        file_base_path = str(file_path)[0:80]
        # print(f'file_base_path: {file_base_path}')
        # Replaces with file base path plus generation path ex. \Gen_5.x
        # Function creating drive paths or downloading attachment files returns the drive path and requires just the
        # file name without the file path prepended to the file name
        file_name: str = file_path.replace(file_base_path, '')

        if attachment_file_name in file_name:
            return get_drive_path('CRD', attachment_file_name, attached_file_name_paths)

    # If none of the file paths matches with what is already in the shared drive,
    # then downloads the new attachment file from the ticket, should not be prompted to get a new file when rerun
    # since the file attachment is now in the shared drive
    crd_path: str = attached_file_name_paths.get('base_paths', {}).get('crd_path')
    return download_attachment_file_from_ticket(ticket_json, attachment_file_name, crd_path, ticket_id)


def get_skudoc_drive_file_path(file_names_from_drive_paths: list, attachment_file_name: str, ticket_json: dict,
                               ticket_id: str, attached_file_name_paths: dict) -> str:
    """
    Return the drive file path for later hyperlink in the excel output

    :param file_names_from_drive_paths: list of current unique document file names ie. CRD, SKUDOC, Other, Data
    :param attachment_file_name: current file name from ticket ie. CRD, SKUDOC, Other, Data
    :param ticket_json: contains data for current ticket ie. TRR number
    :param ticket_id: TRR number, used for later accessing and downloading from ADO if necessary
    :param attached_file_name_paths: file names already stored in the shared drive
    :return: shared drive file path for hyperlink later in the excel output
    """
    for file_path in file_names_from_drive_paths:
        # 80 characters is the current number of characters that include the base file path plus generation folder
        # not ideal for scalability, getting length of file base path is not do-able since Python has difficulty reading
        # back slashes ie. \
        # Might be better solution out there as file path length can change
        file_base_path = str(file_path)[0:84]
        # print(f'file_base_path: {file_base_path}')
        # Replaces with file base path plus generation path ex. \Gen_5.x
        # Function creating drive paths or downloading attachment files returns the drive path and requires just the
        # file name without the file path prepended to the file name
        file_name: str = file_path.replace(file_base_path, '')

        if attachment_file_name in file_name:
            return get_drive_path('SKUDOC', attachment_file_name, attached_file_name_paths)

    # If none of the file paths matches with what is already in the shared drive,
    # then downloads the new attachment file from the ticket, should not be prompted to get a new file when rerun
    # since the file attachment is now in the shared drive
    skudoc_path: str = attached_file_name_paths.get('base_paths', {}).get('skudoc_path')
    return download_attachment_file_from_ticket(ticket_json, attachment_file_name, skudoc_path, ticket_id)


def check_attachment_files(ticket_json: dict, attached_file_name_paths: dict, ticket_id: str) -> dict:
    """
    Get actual due date from ADO
    :param ticket_json:
    :param attached_file_name_paths:
    :param ticket_id:
    :return:
    """
    drive_file_paths: dict = {}

    # Get all attachment file names from TRR for easier iteration later
    attachment_file_names_from_ticket: list = []
    attachment_files = ticket_json.get('relations', {})
    for file in attachment_files:
        if file['rel'] in 'AttachedFile':
            attachment_file_name = file['attributes']['name']
            attachment_file_names_from_ticket.append(attachment_file_name)

    # Iterates through the file names found in the attachment folder within the TRR
    for attachment_file_name in attachment_file_names_from_ticket:

        crd_file_names: list = attached_file_name_paths.get('crd')
        skudoc_file_names: list = attached_file_name_paths.get('skudoc')
        attachment_file_name: str = attachment_file_name.upper().replace('.XLSX', '.xlsx')

        # CRD - Checks whether attachment file is a CRD
        if check_crd_file_name(attachment_file_name) is True:
            crd_drive_path: str = get_crd_drive_file_path(crd_file_names, attachment_file_name,
                                                          ticket_json, ticket_id, attached_file_name_paths)
            drive_file_paths['crd_drive_path'] = crd_drive_path

        if check_skudoc_file_name(attachment_file_name) is True:
            skudoc_drive_path: str = get_skudoc_drive_file_path(skudoc_file_names, attachment_file_name,
                                                                ticket_json, ticket_id, attached_file_name_paths)
            drive_file_paths['skudoc_drive_path'] = skudoc_drive_path

        # # SKUDOC
        # if any(attachment_file_name in file_path_name for file_path_name in skudoc_file_names):
        #     skudoc_drive_path: str = get_drive_path('SKUDOC', attachment_file_name, attached_file_name_paths)
        #     drive_file_paths['skudoc_drive_path'] = skudoc_drive_path
        #     print(f'IN: SKUDOC - {attachment_file_name}')
        #     pass
        #
        # else:
        #     crd_path: str = attached_file_name_paths.get('base_paths', {}).get('crd_path')
        #     skudoc_drive_path: str = get_attachment_file(ticket_json, crd_path, ticket_id, attachment_file_name)
        #     drive_file_paths['skudoc_drive_path'] = skudoc_drive_path
        #     print(f'OUT: SKUDOC - {attachment_file_name}')
        #     pass

        # for file_skudoc_path in skudoc_file_names:
        #     if attachment_file_name in file_skudoc_path and '.xlsx' in attachment_file_name \
        #             and 'CRD' not in attachment_file_name and attachment_file_name[0] == 'M' \
        #             and 'GEN' in attachment_file_name:
        #
        #         drive_file_paths['skudoc_file_path'] = attachment_file_name
        #
        #     elif attachment_file_name not in file_skudoc_path and '.xlsx' in attachment_file_name \
        #             and 'CRD' not in attachment_file_name and attachment_file_name[0] == 'M' \
        #             and 'GEN' in attachment_file_name:
        #
        #         skudoc_path: str = attached_file_name_paths.get('base_paths', {}).get('skudoc_path')
        #         get_attachment_file(ticket_json, skudoc_path, ticket_id, attachment_file_name)
        #         drive_file_paths['skudoc_file_path'] = file_skudoc_path
        #         break

    return drive_file_paths


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


def store_tickets_data(raw_tickets_data: list, attached_file_names: dict, console_server_data: dict) -> dict:
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
            table_data = ticket_json['fields']['System.Description']
            table_data_soup = BeautifulSoup(table_data, 'html.parser')

            table_rows: list = table_data_soup.findAll('tr')
            table_data: dict = get_clean_table_data(table_rows)

            ticket_data['qcl_parts']: list = get_qcl_parts(table_data)

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

            ticket_data['attachment_file_paths']: dict = check_attachment_files(ticket_json,
                                                                                attached_file_names, ticket_id)

            # State of Qual
            ticket_data['state']: str = ticket_json['fields']['System.State']

            # Test Plans Hyperlink, used later for Async
            ticket_data['test_plan_hyperlink']: str = get_test_plan_hyperlink(ticket_json)

            all_tickets_data[ticket_id] = ticket_data

        except json.decoder.JSONDecodeError:
            print(f'\tMicrosoft Azure DevOps did not return all data.')
            print(f'\tError is at TRR {ticket_id}. Recollecting...\n')
            main_method(console_server_data)

        except KeyError:
            print(f'\tNo Description Table... Grabbing from Summary Table Instead in TRR')

            ticket_json: dict = json.loads(raw_ticket_data)

            # For some reason the ticket_json["id"] data type is int
            ticket_id: str = str(ticket_json["id"])
            # print(f'\t\t- Collect  |  {Fore.GREEN}Success{Style.RESET_ALL}   |  {ticket_id}')

            # Table Data
            table_data = ticket_json['fields']['Custom.Summary']
            table_data_soup = BeautifulSoup(table_data, 'html.parser')

            table_rows: list = table_data_soup.findAll('tr')
            table_data: dict = get_clean_table_data(table_rows)

            ticket_data['qcl_parts']: list = get_qcl_parts(table_data)

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

            ticket_data['attachment_file_paths']: dict = check_attachment_files(ticket_json,
                                                                                attached_file_names, ticket_id)

            # State of Qual
            ticket_data['state']: str = ticket_json['fields']['System.State']

            # Test Plans Hyperlink, used later for Async
            ticket_data['test_plan_hyperlink']: str = get_test_plan_hyperlink(ticket_json)

            all_tickets_data[ticket_id] = ticket_data

    all_tickets_data['unique_bios'] = list(set(all_bios))
    all_tickets_data['unique_bmc'] = list(set(all_bmc))
    all_tickets_data['unique_cpld'] = list(set(all_cpld))
    all_tickets_data['unique_os'] = list(set(all_os))

    return all_tickets_data


def get_qcl_parts(table_data):
    qcl_parts: list = []

    for item in table_data['all_part_numbers']:
        qcl_parts.append(item)

    for component in table_data:
        component = component.strip()
        component_value: str = table_data[component]

        # if ticket_id == '447910':
        #     print(f'{component} -- {component_value}')

        if not component_value:
            pass
        # elif '447910' in ticket_id:
        #     if 'QCL' in component.upper() or component == 'part_number' or 'DIMM' in component.upper() or \
        #             'NVME' in component.upper() or 'HDD' in component.upper() or 'SDD' in component.upper():
        #         print(f'{component} -- {component_value}')
        #         qcl_parts.append(component_value)
        elif 'part_number' in component:
            qcl_parts.append(component_value)

        elif 'qcl' in component or 'dimm' in component or 'nvme' in component or \
                'hdd' in component or 'sdd' in component or 'part_number' in component:
            qcl_parts.append(component_value)

    return qcl_parts


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


def access_due_dates_well(unique_tickets) -> dict:
    """
    Get actual due date from ADO
    :param unique_tickets:
    :return:
    """
    due_dates: dict = {}

    for ticket in unique_tickets:
        json_fields: dict = request_ado_json(ticket).get('fields', {})

        due_dates['expected_task_start']: dict = str(json_fields['AzureCSI-V1.1.ExpectedTaskStart'])
        due_dates['expected_task_completion']: dict = str(json_fields['AzureCSI-V1.1.ExpectedTaskCompletion'])
        due_dates['actual_qual_start_date']: dict = str(json_fields['Custom.ActualQualStartDate'])
        due_dates['actual_qual_end_date']: dict = str(json_fields['Custom.ActualQualEndDate'])

    return due_dates


def calculate_test_progress(results_outcome: int, results_state: int) -> str:
    """
    HTML file for Test Plans do not show the actual status for some reason.
    They are encoded in these outcome and state numbers which refer to the status.
    :param results_outcome:
    :param results_state:
    :return:
    """
    if results_outcome == 0 and results_state == 1:
        return 'Active'
    elif results_outcome == 2 and results_state == 2:
        return 'Passed'
    elif results_outcome == 3 and results_state == 3:
        return 'Failed'
    elif results_outcome == 11 and results_state == 2:
        return 'Not Applicable'
    elif results_outcome == 7 and results_state == 3:
        return 'Blocked'
    else:
        'None'


def store_test_cases_execute(test_case_json: dict) -> dict:
    """
    After accessing test cases, store relevant information into dictionary for key-value calls later.
    :param test_case_json:
    :return:
    """

    test_points = test_case_json.get('data', {}). \
        get('ms.vss-test-web.test-plans-hub-refresh-data-provider', {}). \
        get('testPoints', {})

    all_test_case_data = {}
    for index, test_case in enumerate(test_points, start=1):
        test_case_data: dict = {'secondary_id': test_case.get('id', 'None'),
                                'is_active': test_case.get('isActive', 'None'),
                                'is_automated': test_case.get('isAutomated', 'None'),
                                'results_outcome': test_case.get('results', {}).get('outcome', 'None'),
                                'results_state': test_case.get('results', {}).get('state', 'None'),
                                'test_case_id': test_case.get('testCaseReference', {}).get('id', 'None'),
                                'test_name': test_case.get('testCaseReference', {}).get('name', 'None'),
                                'test_state': test_case.get('testCaseReference', {}).get('state', 'None'),
                                'test_summary': calculate_test_progress(
                                    test_case.get('results', {}).get('outcome', 'None'),
                                    test_case.get('results', {}).get('state', 'None'))}

        all_test_case_data[f'test_case_{str(index)}'] = test_case_data

    return all_test_case_data


def access_test_cases(hyperlink: str):
    """
    Access Test Cases based from ticket number's hyperlink to test plans
    :param hyperlink:
    :return:
    """
    # Define URL version is usually blank for Test Cases progress
    # Execute URL usually contains the progress of a ticket
    execute_link: str = hyperlink.replace('define?planId=', 'execute?planId=')

    user_password: str = Ado.token_name + ':' + Ado.personal_access_token
    base64_user_password: str = base64.b64encode(user_password.encode()).decode()
    headers = {'Authorization': 'Basic %s' % base64_user_password}

    response_test_cases = requests.get(execute_link, headers=headers)

    soup = BeautifulSoup(response_test_cases.text, 'html.parser')
    test_cases_json = json.loads(soup.find('script', type='application/json').contents[0])
    test_cases_data = store_test_cases_execute(test_cases_json)

    return test_cases_data


def get_test_plan_hyperlink(ticket_json: dict) -> str:
    """
    Get the hyperlink for the Test Plans for later Async
    :param ticket_json:
    :return:
    """
    potential_hyperlinks: list = []

    # If the status of the ticket is already signed off
    # Then there would be no access to the test cases
    ticket_state = str(ticket_json.get(f'fields', {}).get('System.State', 'None')).lower()
    if ticket_state == 'signed off':
        return 'None'

    # Grabbing any potential hyperlinks that lead to test cases
    attachment_files = ticket_json.get('relations', {})
    for file in attachment_files:
        if file['rel'] == 'Hyperlink':
            hyperlink = file['url']
            potential_hyperlinks.append(hyperlink)

    # There should be only one test plan for each ticket
    if len(potential_hyperlinks) == 0:
        return 'None'
    elif len(potential_hyperlinks) == 1:
        # Hyperlinks gives us the define not the execute
        # The execute is what we want which contains the progress of the test plans
        return potential_hyperlinks[0].replace('define?planId=', 'execute?planId=')


def get_hyperlink_test_plan(ticket_json: dict) -> dict:
    """
    Navigating through the JSON to extract Hyperlink which then leads to the test plans
    :param ticket_json:
    :return:
    """
    all_files: list = []

    # If the status of the ticket is already signed off
    # Then there would be no access to the test cases
    ticket_state = str(ticket_json.get(f'fields', {}).get('System.State', 'None')).lower()
    if ticket_state == 'signed off':
        return {}

    # Grabbing any potential hyperlinks that lead to test cases
    attachment_files = ticket_json.get('relations', {})
    for file in attachment_files:
        if file['rel'] == 'Hyperlink':
            hyperlink = file['url']
            all_files.append(hyperlink)

    # There should be only one test plan for each ticket
    if len(all_files) == 0:
        return {}
    elif len(all_files) == 1:
        return access_test_cases(all_files[0])


def store_hyperlinks(processed_tickets_data: dict) -> list:
    """
    Get the hyperlinks, state of the tickets for later Async
    :param processed_tickets_data:
    :return:
    """
    all_ticket_data: list = []
    for ticket_number in processed_tickets_data:
        if 'broken_targets' not in ticket_number:
            ticket_data: dict = {'ticket': ticket_number,
                                 'hyperlink': processed_tickets_data.get(ticket_number).get('test_plan_hyperlink'),
                                 'state': processed_tickets_data.get(ticket_number).get('state')}

            all_ticket_data.append(ticket_data)

    return all_ticket_data


async def get_test_cases_data(hyperlinks, headers):
    """

    :param hyperlinks:
    :param headers:
    :return:
    """
    tasks = [asyncio.create_task(fetch_test_case_data(test_case, headers)) for test_case in hyperlinks]

    return await asyncio.gather(*tasks)


async def fetch_test_case_data(test_case_data: dict, headers):
    """
    Grabs the information from Azure Devops per session depending on how many tickets in the form of URls
    :param test_case_data:
    :param headers:
    :return:
    """
    import codecs
    test_case_url = test_case_data.get('hyperlink')
    ticket = test_case_data.get('ticket')

    if test_case_url == 'None':
        return 'None'
    else:
        async with ClientSession(headers=headers) as session:
            async with session.get(test_case_url) as response:
                await asyncio.sleep(0.5)
                # ticket_data: str = await response.text()
                ticket_data: str = await response.text()

    return ticket_data


def drive_access_prompt() -> None:
    """
    If OSError via os.mk_dir fails, force user to enter credentials to access Z:Drive
    :return:
    """
    print('\tNeed to access Z:Drive. Please enter credentials... (Note: Password not visible for confidentiality)')
    print('\tNOTE: Accessing Z:Drive is problematic right now. Get out of Pipe Cleaner and manually access Z:Drive')
    print('\tWith proper credentials. Then restart Pipe Cleaner....')

    #  Colorama does not work with input, must do print first
    print(f'\t{Fore.LIGHTBLUE_EX}Username: {Style.RESET_ALL}', end='')
    username = input(f'')

    print(f'\t{Fore.LIGHTBLUE_EX}Password: {Style.RESET_ALL}', end='')
    # password = getpass(f'')
    password = input(f'')

    print(f'username: {username} | {password}')
    try:
        subprocess.call(fr'net use z: \\172.30.1.100\pxe /u:172.30.0.100\{username} {password}', shell=True)
        # os.system(r"NET USE z: \\172.30.1.100\pxe %s /USER:%s\%s" % (password, username, username))
        # subprocess.call(fr'net use z: \\172.30.1.100\pxe /user:172.30.0.100\{username} {password}', shell=True)
    except OSError:
        print(f'\t{Fore.RED}Wrong Credentials{Style.RESET_ALL} - Please try again.')
        drive_access_prompt()

    input()


def access_prompt():
    drive_access_prompt()


def get_existing_attachment_files(current_path: str, document_paths: dict, processor_generations: tuple) -> list:
    """
    Grab files downloaded from TRRs' CRD and SKUDOCs already
    """
    file_path: str = document_paths.get(current_path)

    existing_attachment_files: list = []
    if 'OTHER' in file_path.upper():
        try:
            other_attached_files: list = os.listdir(file_path)

            for file_name in other_attached_files:
                file_name: str = file_name.upper().replace('.XLSX', '.xlsx')
                existing_attachment_files.append(file_name)

        except OSError:
            pass

    else:
        for generation in processor_generations:
            generation_path: str = fr'{file_path}\{generation}'
            try:
                generation_attached_files: list = os.listdir(generation_path)

                for file_name in generation_attached_files:
                    file_name: str = file_name.upper().replace('.XLSX', '.xlsx')
                    entire_path: str = fr'{file_path}\{generation}\{file_name}'

                    existing_attachment_files.append(entire_path)
            except OSError:
                pass

    return existing_attachment_files


def get_documents_from_shared_drive() -> dict:
    """
    Get CRD, SKUDOCs, or Datasheets from VSE shared drive.
    """
    document_paths: dict = get_shared_drive_paths()
    processor_generations: tuple = ('Gen_5.x', 'Gen_6.x', 'Gen_7.x', 'Gen_8.x')

    check_attachment_folders(document_paths, processor_generations)

    crd_file_names: list = get_existing_attachment_files('crd_path', document_paths, processor_generations)
    skudoc_file_names: list = get_existing_attachment_files('skudoc_path', document_paths, processor_generations)
    other_file_names: list = get_existing_attachment_files('other_path', document_paths, processor_generations)

    return {'crd': crd_file_names, 'skudoc': skudoc_file_names, 'other': other_file_names, 'base_paths': document_paths}


def get_shared_drive_paths() -> dict:
    shared_drive_path: str = r'\\172.30.1.100\pxe\Kirkland_Lab\Microsoft_CSI\Documentation\PipeCleaner_Attachments'

    return {'skudoc_path': fr'{shared_drive_path}\SKUDOC',
            'crd_path': fr'{shared_drive_path}\CRD',
            'other_path': fr'{shared_drive_path}\Other'}


def get_ticket_urls(console_server_data) -> list:
    unique_tickets: list = console_server_data.get('host_groups_data', {}).get('all_unique_tickets')

    return [f'https://azurecsi.visualstudio.com/CSI%20Commodity%20Qualification/_apis/wit/workitems?'
            f'id={ticket}&$expand=all&api-version=5.1' for ticket in unique_tickets]


def main_method(console_server_data: dict):
    """
    Using asynchronous for getting all relevant data from Azure DevOps
    """
    ticket_urls: list = get_ticket_urls(console_server_data)

    attached_file_names: dict = get_documents_from_shared_drive()

    raw_tickets_data: list = asyncio.run(get_ticket_data(ticket_urls))

    return store_tickets_data(raw_tickets_data, attached_file_names, console_server_data)
