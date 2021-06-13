"""
Terminal output giving user information on Pipe or TRR processing
"""

import socket
import subprocess
import sys
from getpass import getpass

from colorama import Fore, Style
from colorama import init as enable_text_color
from datetime import datetime

from pipe_cleaner.src.terminal_properties import intro_section
from pipe_cleaner.src.log_database import access_database_document


def is_current_version(version_number: str, current_version: str) -> bool:
    """
    Assures Pipe Cleaner version is up to date.  Fetches string from MongoDB.
    """
    print(f'\t\t- Confirming version...')

    if version_number not in current_version:
        return False
    else:
        return True


def get_version_from_database() -> str:
    """
    Assure Pipe Cleaner version is up to date fetched from database
    """
    document = access_database_document('in_house', 'pipe_cleaner')

    version_db: dict = document.find_one({'_id': 'current_version'})
    if not version_db:
        document.insert_one({'_id': 'current_version',
                             'version': '2.7.8',
                             'date_time': datetime.today().strftime('%Y-%m-%d-%H:%M:%S')})
        return version_db['version']
    else:
        return version_db['version']


def show_pipe_cleaner_banner() -> None:
    """
    Create Intro Banner for Terminal. Says Pipe Cleaner
    """
    print(f'\n {Fore.LIGHTBLUE_EX}')
    print('  #######    ##   #######    #######           ######   ##        #######       ##       ##    ##   #######'
          '   #######')
    print('  ##    ##   ##   ##    ##   ##              ##         ##        ##           ####      ###   ##   ##     '
          '   ##    ##')
    print('  ##    ##   ##   ##    ##   ##             ##          ##        ##          ##  ##     ####  ##   ##     '
          '   ##    ##')
    print('  #######    ##   #######    ######         ##          ##        ######     ##    ##    ## ## ##   ###### '
          '   #######')
    print('  ##         ##   ##         ##             ##          ##        ##        ##########   ##  ####   ##     '
          '   ##  ##')
    print('  ##         ##   ##         ##              ##         ##        ##        ##      ##   ##   ###   ##      '
          '  ##   ##')
    print('  ##         ##   ##         #######           ######   #######   #######   ##      ##   ##    ##   ####### '
          '  ##    ##')
    print(f' {Style.RESET_ALL}')


def show_intro_sentence(version_number: str, user_name: str, current_location: str) -> None:
    """
    Statement after Banner
    :return:
    """
    intro_section(version_number, user_name, current_location)


def is_vpn_on(current_location: str) -> bool:
    """
    Checks if operating system is on VPN.
    VSE Kirkland Lab is on the 172.18.xxx network.
    VSE Thailand and Europe Network will be put up later.
    """
    print(f'\t\t- Confirming VPN...')

    if 'Kirkland' in current_location:
        vse_kirkland_network: str = '172.'
        ip_addresses: list = socket.gethostbyname_ex(socket.gethostname())[-1]

        for ip_address in ip_addresses:
            if vse_kirkland_network in ip_address[0:4]:
                return True

        else:
            return False


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
    password = getpass(f'')

    try:
        subprocess.call(fr'net use z: \\172.30.1.100\pxe /u:172.30.0.100\{username} {password}', shell=True)

    except OSError:
        print(f'\t{Fore.RED}Wrong Credentials{Style.RESET_ALL} - Please try again.')
        drive_access_prompt()


def number_of_things(attachments: list, thing: str) -> str:
    """
    Files or file.
    :param thing:
    :param attachments:
    :return:
    """
    if len(attachments) == 0 or len(attachments) == 1:
        return f'{thing}'
    else:
        return f'{thing}s'


def initialize_text_color() -> None:
    """
    init() from colorama needs to be called in order to print color text in Terminal Output
    """
    enable_text_color()


def get_available_chooses() -> str:
    """
    Give user chooses between default, send, and inventory.
    """
    print(f'\n\n\tChoose between these options...\n')
    print(f'\t\tR  ->  Request Form - New Inventory')
    print(f'\t\tU  ->  Update Form - Log S/Ns')

    print(f'\n\t\tN  ->  Normal Mode')
    print(f'\t\tT  ->  Total Inventory')
    return input(f'\n\tChoose letter and press enter: ')


def get_locations_for_material(location: str) -> str:
    """
    Get starting location of inventory movement.
    """
    print(f'\n\n\t{"-" * 60}')
    print(f'\n\n\t{location.title()} Location:')

    print(f'\n\t\tP  ->  Pipe - From Racks')
    print(f'\t\tC  ->  Cage - Inventory Storage')
    print(f'\t\tS  ->  Shipment - Inventory Storage')

    print(f'\n\t\tH  ->  Hard Drive Room - Isolated Inventory')
    print(f'\t\tQ  ->  Quarantine - Isolated Inventory')
    print(f'\t\tM  ->  Mini Labs - Isolated Inventory')

    chosen_letter: str = input(f'\n\tChoose letter and press enter: ')
    location: str = location_letter_to_word(chosen_letter, location)

    print(f'\tChosen Location: {location}')

    return location


def location_letter_to_word(letter: str, location: str) -> str:
    """
    Convert letter for start and end location to word.
    """
    letter: str = letter.upper()

    if letter == 'P':
        return 'Pipe'

    elif letter == 'C':
        return 'Cage'

    elif letter == 'S':
        return 'Shipment'

    elif letter == 'H':
        return 'Hard Drive Room'

    elif letter == 'Q':
        return 'Quarantine'

    elif letter == 'M':
        return 'Mini Labs'

    else:
        print(f'\tProblem: Unavailable Letter. Try again.')
        get_locations_for_material(location)


def get_user_response_on_chooses(user_response: str) -> str:
    """
    Get response based on user input from terminal.
    :param user_response:
    :return:
    """
    user_input: str = user_response.upper()

    if user_input == 'N':
        return 'N'

    elif 'NORMAL' in user_response:
        return 'N'

    elif user_input == 'R':
        return 'R'

    elif 'REQUEST' in user_input:
        return 'R'

    elif user_input == 'U':
        return 'U'

    elif 'UPDATE' in user_input:
        return 'U'

    elif 'TOTAL' in user_input:
        return 'T'

    elif user_input == 'T':
        return 'T'

    elif 'UPDATE' in user_input:
        return 'U'

    elif user_input == 'U':
        return 'U'

    else:
        print(f'\t{Fore.RED}Unavailable response{Style.RESET_ALL}...')
        user_chose: str = get_available_chooses()
        get_user_response_on_chooses(user_chose)


def print_terminal_intro(current_location: str, default_user_name: str, version_number: str) -> None:
    """
    Start the terminal intro for users to see.
    """
    initialize_text_color()
    show_pipe_cleaner_banner()
    show_intro_sentence(version_number, default_user_name, current_location)


def print_unmet_requirements(vpn_status: bool, version_status: bool, current_version: str,
                             updated_version: str) -> None:
    """
    If VPN is not on and/or Pipe Cleaner version is not on, will alert user.
    """
    print(f'\n\tProblem:')
    if not vpn_status and not version_status:
        print(f'\t- VPN: OFF')
        print(f'\t- Version: WRONG')
        print(f'\n\tSolution:')
        print(f'\t- Turn on GlobalProtect VPN.  Contact IT if issues persist.')
        print(f'\t- Update to latest version - {updated_version}.  Outdated version - {current_version}')

    elif not vpn_status and version_status:
        print(f'\t- VPN: OFF')
        print(f'\n\tSolution:')
        print(f'\t- Turn on GlobalProtect VPN.  Contact IT if issues persist.')

    elif vpn_status and not version_status:
        print(f'\t- Version: WRONG')
        print(f'\n\tSolution:')
        print(f'\t- Update to latest version - {updated_version}.  Current version - {current_version}')

    input(f'\n\tPress enter to exit:')
    sys.exit()


def get_inventory_position(location: str) -> str:
    """
    Location position as a word
    """
    return get_locations_for_material(location)


def asks_user_quantity() -> str:
    """
    Requests the user to enter the amount of inventory needed to moved.
    """
    print(f'\n\n{"-" * 60}\n\n')
    print(f'\tQuantity Moving:')
    print(f'\n\tWhat is the quantity of commodities to be moved?')
    quantity: str = input(f'\tRequest Quantity: ')

    if not quantity.isdigit():
        print(f'\tISSUE: Not a digit. Try again.')
        asks_user_quantity()

    elif '.' in quantity:
        print(f'\tISSUE: Not a whole number. Try again.')
        asks_user_quantity()

    elif int(quantity) <= 0:
        print(f'\tISSUE: Must be greater than 0.')
        asks_user_quantity()

    else:
        print(f'\tQuantity Entered: {quantity}')
        return quantity


def ask_inventory_questions(response: str) -> dict:
    """
    Ask user other questions pertaining to inventory.
    """
    start_location: str = get_inventory_position('start')
    end_location: str = get_inventory_position('end')

    if start_location != end_location:
        return {'start': start_location,
                'end': end_location,
                'quantity': asks_user_quantity(),
                'letter': response}
    else:
        ask_inventory_questions(response)


def run_terminal(version_number: str, default_user_name: str, current_location: str) -> dict:
    """
    Generic Intro for Pipe Cleaner including Banner, Intro Sentence, and Local Network checks.
    Then asks for user Role.
    """
    print_terminal_intro(current_location, default_user_name, version_number)

    vpn_status: bool = is_vpn_on(current_location)
    updated_version: str = get_version_from_database()
    version_status: bool = is_current_version(version_number, updated_version)

    if vpn_status and version_status:
        user_chose: str = get_available_chooses()
        response: str = get_user_response_on_chooses(user_chose)

        if 'R' == response:
            return ask_inventory_questions(response)

        else:
            return {'letter': response}

    else:
        print_unmet_requirements(vpn_status, version_status, version_number, updated_version)
