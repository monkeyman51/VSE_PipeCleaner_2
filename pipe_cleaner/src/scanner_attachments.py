import base64
import os
from datetime import datetime
from json import loads

import requests
from colorama import Fore, Style

from pipe_cleaner.src.credentials import AccessADO as Ado
from pipe_cleaner.src.terminal import number_of_things
from pipe_cleaner.src.terminal_properties import terminal_header_section
from pipe_cleaner.src.data_access import request_json_from_ado

"""
Responsible for extracting data from Attachment Files within TRR.  Prints out Terminal information.
"""

# For Summary Tallies
total_tally: list = []
crd_tally: list = []
skudoc_tally: list = []
other_tally: list = []


def check_attachment_folders(attachment_path: str, crd_path: str, skudoc_path: str) -> list:
    """
    Creates directories for CRD and SKUDOC in Z:Drive. Pass if already created.
    Ensures repositories for CRD/SKUDOC
    :param attachment_path: Base Directory for TRR attachment files
    :param crd_path:
    :param skudoc_path:
    :return:
    """
    # Gen 4 and 5 becoming obsolete. Gen 9 coming out soon.
    generations: list = ['Gen_5.x',
                         'Gen_6.x',
                         'Gen_7.x',
                         'Gen_8.x']

    # Create Main Directory
    try:
        os.mkdir(attachment_path)
        os.mkdir(crd_path)
        os.mkdir(skudoc_path)

    except FileExistsError:
        pass

    # Create Gen Directories, none for Datasheet since no Gen
    try:
        for processor in generations:
            os.mkdir(fr'{crd_path}\{processor}')
            os.mkdir(fr'{skudoc_path}\{processor}')

    except FileExistsError:
        pass

    return generations


def check_existing_files(file_path: str, gen_path: str, file_name: str) -> bool:
    """
    Check if CRD/SKUDOC/Datasheet is in the directories before attempting to download information from TRR
    :return:
    """
    # Stores all files
    folder = []

    try:
        # Usually PDF
        if gen_path == 'None':
            for file in os.listdir(f'{file_path}'):
                folder.append(file)
            if file_name in folder:
                return True
            else:
                return False

        # CRD/SKUDOC
        elif isinstance(gen_path, str):
            for file in os.listdir(f'{file_path}/Gen_{gen_path}.x'):
                folder.append(file)
            if file_name in folder:
                return True
            else:
                return False
    except FileExistsError:
        print(f'\tWARNING: {Fore.RED}File Path{Style.RESET_ALL} not found -> {file_path}')


def request_attachments(ticket_number: str, attachment_files: dict, crd_path: str, skudoc_path: str):
    """
    Requests data from ADO
    :param ticket_number:
    :param skudoc_path:
    :param crd_path:
    :param attachment_files: Contains name_to_url, name_to_gen, name_to_type
    :return:
    """
    crd_info = {}

    # Unpack dictionary
    name_to_url = attachment_files['name_to_url']
    name_to_gen = attachment_files['name_to_gen']
    name_to_type = attachment_files['name_to_type']

    user_password = f'{Ado.token_name}:{Ado.personal_access_token}'
    base64_user_password = base64.b64encode(user_password.encode()).decode()
    headers = {'Authorization': 'Basic %s' % base64_user_password}

    new = f'{Fore.GREEN}NEW{Style.RESET_ALL}'

    print(f'\n\tTRR {ticket_number} - {len(name_to_url)} Attachment Files Found:')

    # Adds attachment files from the TRR to Z:Drive
    initial = 0
    gen_number = initial + 4
    while initial < len(name_to_url):
        # Index dictionary to make while loop work
        index_url = list(name_to_url.keys())[initial]
        index_gen = list(name_to_gen.keys())[initial]
        index_type = list(name_to_type.keys())[initial]
        gen_file_path = str(name_to_gen[index_gen])

        try:
            ado_response = requests.get(name_to_url[index_url], headers=headers, timeout=1)
        except requests.exceptions.Timeout:
            print('\n  * ADO Response: {Fore.RED}Timeout Occurred{Style.RESET_ALL}... attempting again\n')
            requests.get(name_to_url[index_url], headers=headers, timeout=5)
        except requests.exceptions.ConnectionError:
            print('\n  * ADO Response: {Fore.RED}Timeout Occurred{Style.RESET_ALL}... attempting again\n')
            requests.get(name_to_url[index_url], headers=headers, timeout=5)
        except requests.exceptions.HTTPError:
            print('\n  * ADO Response: {Fore.RED}Timeout Occurred{Style.RESET_ALL}... attempting again\n')
            requests.get(name_to_url[index_url], headers=headers, timeout=5)

        crd_file_path = f'{crd_path}/Gen_{gen_file_path}.x/{index_url}'

        # CRD attachment file
        if name_to_type[index_type] == 'CRD':
            file_exist = check_existing_files(crd_path, gen_file_path, index_url)
            if file_exist is True:
                print(f'\t\t- OLD {name_to_type[index_type]}: File Exists')
                crd_tally.append(1)

                crd_info['crd_file'] = index_url
                crd_info['file_path'] = crd_file_path

            elif file_exist is False:
                with open(f'{crd_file_path}', 'wb') as f:
                    f.write(ado_response.content)

                print(f'\t\t- {new} {Fore.GREEN}{name_to_type[index_type]}{Style.RESET_ALL}: Downloaded '
                      f'\n\t\t>> {index_url}')

                crd_info['crd_file'] = index_url
                crd_info['file_path'] = crd_file_path

                skudoc_tally.append(1)

        # SKUDOC attachment file
        elif name_to_type[index_type] == 'SKUDOC':
            file_exist = check_existing_files(skudoc_path, gen_file_path, index_url)
            if file_exist is True:
                print(f'\t\t- OLD {name_to_type[index_type]}: File Exists')
                skudoc_tally.append(1)

            elif file_exist is False:
                with open(f'{skudoc_path}/Gen_{gen_file_path}.x/{index_url}', 'wb') as f:
                    f.write(ado_response.content)
                print(f'\t\t- {new} {Fore.GREEN}{name_to_type[index_type]}{Style.RESET_ALL}: Downloaded '
                      f'\n\t\t>> {index_url}')
                skudoc_tally.append(1)

        # PDF attachment file
        # elif name_to_type[index_type] == 'PDF':
        #     file_exist = check_existing_files(datasheet_path, gen_file_path, index_url)
        #     if file_exist is True:
        #         print(f'   - OLD {name_to_type[index_type]}: File Exists -> {index_url}')
        #
        #     elif file_exist is False:
        #         with open(f'{datasheet_path}/{index_url}', 'wb') as f:
        #             f.write(ado_response.content)
        #         print(f'   - NEW {name_to_type[index_type]}: Downloaded -> {index_url}')

        # Other attachment file
        else:
            print(f'\t\t- OTHER: Not CRD or SKUDOC')
            other_tally.append(1)

        initial += 1
        gen_number += 1

    return crd_info


def get_generation_from_name(file_name: str) -> int:
    """
    Checks Generation Number based on file name from Attachment Section in TRR
    For key-value pair in dictionary ie. File Name to Processor Generation
    :param file_name:
    :return: Gen number
    """
    initial = 4  # Generation to start
    while initial < 10:  # Gen 10 not coming out soon - 9/26/2020
        if f'GEN{str(initial)}' in file_name:
            return initial
        initial += 1


def get_type_from_name(file_name: str) -> str:
    """
    Checks File Type based on either CRD, PDF, or SKUDOC from Attachment Section in TRR
    For key-value pair in dictionary ie. File Name to Processor Generation
    :param file_name:
    :return: CRD, PDF, or SKUDOC
    """
    # Checks for these items
    crd = 'CRD'
    gen = 'GEN'
    pdf = '.pdf'
    excel_file = '.xlsx'
    old_excel = '.xls'
    first_character = 'M'
    word = '.docx'
    msg = '.msg'

    # Gets Previous, Current, Next Year for Fiscal Year Check in File Name to verify SKUDOC
    today_date = datetime.today()
    previous_year = str(today_date.year - 1)[2:]
    current_year = str(today_date.year)[2:]
    next_year = str(today_date.year + 1)[2:]

    # CRD
    if crd in file_name and \
            excel_file in file_name and \
            gen in file_name and \
            first_character in file_name[0]:
        return 'CRD'

    # SKUDOC
    elif file_name[0] in first_character and \
            gen in file_name and \
            excel_file in file_name:
        if f'FY{previous_year}' not in file_name \
                or f'FY{current_year}' not in file_name \
                or f'FY{next_year}' not in file_name:
            return 'SKUDOC'

    # PDF
    elif pdf in file_name:
        return 'PDF'

    # Word
    elif word in file_name:
        return 'Word'

    # Old Excel
    elif old_excel in file_name:
        return 'Old_Excel'

    # MSG
    elif msg in file_name:
        return 'MSG'

    else:
        return 'None'


def check_attachment_files(ticket_number: str, json_file: dict) -> dict:
    """
    Check if there are more than 3 files inside the Attachment part of the TRR. It's for PM Review.
    Also, validates for those 3 files. ie. CRD, PDF, SKUDOC
    :param json_file: JSON file from TRR
    :param ticket_number: TRR Number
    :return: name_to_url, name_to_gen
    """
    # Total needed for attachments
    total_needed = 3

    attachment_file_names: list = []
    attachment_file_types: list = []

    # File Name from Attachments section in TRR as key for File URL
    name_to_url: dict = {}
    name_to_gen: dict = {}
    name_to_type: dict = {}  # File Name to File Type ie. CRD, PDF, or SKUDOC

    # Stores name_to_url, name_to_gen into dict for return
    attachment_files_info: dict = {}

    # Current File Names start with M. - 9/26/2020
    # first_letter_file_name: str = 'M'

    # Looks for attachment files names in TRR and stores in a list.
    # Stores File Name to URL in Dict
    # Stores File Name to Generation Number
    # Stores File Name to File Type
    for item in json_file['relations']:
        if item['rel'] in 'AttachedFile':
            # Get information from JSON file
            file_url = item['url']
            file_name = item['attributes']['name']
            upper_name = str(file_name).upper()  # cap all to compare easier later
            gen_number = get_generation_from_name(file_name)
            file_type = get_type_from_name(file_name)

            # Stores in dict, list
            name_to_gen[upper_name] = gen_number
            name_to_url[upper_name] = file_url
            name_to_type[upper_name] = file_type
            attachment_file_names.append(upper_name)
            attachment_file_types.append(file_type)

    # Stores dictionaries in list for return
    attachment_files_info['name_to_gen'] = name_to_gen
    attachment_files_info['name_to_url'] = name_to_url
    attachment_files_info['name_to_type'] = name_to_type

    # Warns if files in attachment is less than 3
    if len(attachment_file_names) < total_needed:
        print(f'\n\tWARNING: TRR {Fore.RED}{ticket_number}{Style.RESET_ALL} - '
              f'Currently has {Fore.RED}{len(attachment_file_names)} '
              f'Attachment {number_of_things(attachment_file_names, "File")}'
              f'{Style.RESET_ALL}. Needs 3 or more files from TRR.')

        for file_name in attachment_file_types:

            # Warns no CRD file
            if 'CRD' not in attachment_file_types:
                print(f'\t\t- {Fore.RED}CRD{Style.RESET_ALL} file not found.')

            # Warns no SKUDOC file
            if 'SKUDOC' not in attachment_file_types:
                print(f'\t\t- {Fore.RED}SKUDOC{Style.RESET_ALL} file not found.')

            # Warns no PDF or no datasheets
            if 'PDF' not in attachment_file_types:
                print(f'\t\t- {Fore.RED}PDF/WORD/MSG{Style.RESET_ALL} file not found for '
                      f'{Fore.RED}Datasheet{Style.RESET_ALL}.')

    return attachment_files_info


def json_from_ado(ado_response):
    """
    Create JSON file from ADO Response.
    :param ado_response:
    :return:
    """
    with open(f'pipe_cleaner/src/attachment_files.json', 'w', encoding='utf-8') as f:
        f.write(ado_response.text)

    with open(f'pipe_cleaner/src/attachment_files.json', 'r') as f:
        json_file = loads(f.read())

    return json_file


def main_method(unique_tickets: list) -> dict:
    """
    Gets Unique Tickets from Pipe then grabs attachments associated to them
    :param unique_tickets: Unique TRRs
    :return: Ticket Number to JSON
    """
    ticket_to_crd: dict = {}
    # ticket_to_json: dict = {}

    attachment_path: str = r'\\172.30.1.100\pxe\Kirkland_Lab\Microsoft_CSI\Documentation\PipeCleaner_Attachments'

    crd_path: str = r'\\172.30.1.100\pxe\Kirkland_Lab\Microsoft_CSI\Documentation' \
                    r'\PipeCleaner_Attachments\CRD'

    skudoc_path: str = r'\\172.30.1.100\pxe\Kirkland_Lab\Microsoft_CSI\Documentation' \
                       r'\PipeCleaner_Attachments\SKUDOC'

    datasheet_path: str = r'\\172.30.1.100\pxe\Kirkland_Lab\Microsoft_CSI\Documentation' \
                          r'\PipeCleaner_Attachments\Other'

    # Checks if CRD, SKUDOC, Datasheet are established
    check_attachment_folders(attachment_path, crd_path, skudoc_path)

    terminal_header_section("Attachment Files", "Pulling Information from Test Run Requests")

    print(f'\t{Fore.GREEN}Extracting Attachment Files{Style.RESET_ALL} from TRRs...\n')

    print(f'\tFile Paths:')
    print(f'\tCRD Path -> {crd_path}')
    print(f'\tSKUDOC Path -> {skudoc_path}')

    for ticket_number in unique_tickets:
        response = request_json_from_ado(ticket_number)
        json_file = json_from_ado(response)
        attachment_files = check_attachment_files(ticket_number, json_file)
        crd_info = request_attachments(ticket_number, attachment_files, crd_path, skudoc_path)
        # ticket_to_json[ticket_number] = json_file
        ticket_to_crd[ticket_number] = crd_info

    total_tally.append(sum(crd_tally))
    total_tally.append(sum(skudoc_tally))
    total_tally.append(sum(other_tally))

    print(f'\n\t{Fore.YELLOW}File Attachments Summary{Style.RESET_ALL}:')
    print(f'\t\t- Total TRRs: {len(unique_tickets)}')
    print(f'\t\t- Total Files: {sum(total_tally)}')
    print(f'\t\t- CRD Files: {sum(crd_tally)} Excel')
    print(f'\t\t- SKUDOC Files: {sum(skudoc_tally)} Excel')
    print(f'\t\t- Other Files: {sum(other_tally)}')

    return ticket_to_crd
    # return ticket_to_json


# tickets = ['307674']
# #
# # HMA84GR7CJR4N-VK
# # HMA84GR7JJR4N-VK
# # Boot: 0004.0100
# # ImageA: V010D.E84
# # ImageB: V010D.E84
#
# main_method(tickets)
