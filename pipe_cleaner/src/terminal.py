"""
Terminal output giving user information on Pipe or TRR processing
"""

import socket
import subprocess
from getpass import getpass

from colorama import Fore, Style
from colorama import init as enable_text_color

from pipe_cleaner.src.terminal_properties import intro_section


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


def check_local_network() -> None:
    """
    Checks if operating system is on VPN.
    VSE Kirkland Lab is on the 172.18.xxx network.
    VSE Thailand and Europe Network will be put up later.
    :return: True for validation
    """
    vse_kirkland_network: str = '172.'

    print(f'\t{Fore.GREEN}Attempting to connect{Style.RESET_ALL} to network...')

    condition = 0
    while condition == 0:
        ip_addresses: list = socket.gethostbyname_ex(socket.gethostname())[-1]
        initial = 0
        while initial < len(ip_addresses):
            if vse_kirkland_network in ip_addresses[initial]:
                print(f'\n\tConnected to GlobalProtect. Connected to VSE Kirkland network...\n')

                condition += 1
                break
            initial += 1
        if condition == 0:
            print(f"\n  {Fore.RED}Not connected{Style.RESET_ALL} to GlobalProtect!")
            print("\n  Connect to VPN first then wait a few seconds for connection.")
            input("  Press enter to try again...")


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
    :return:
    """
    print(f'\n\tChoose between these options...')
    print(f'\tn -> Normal Mode')
    print(f'\ti -> Inventory Mode')
    print(f'\ts -> Send Inventory')
    print(f'\tt -> Total Inventory')
    return input(f'\n\tChoose option: ')


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

    elif user_input == 'I':
        return 'I'

    elif 'INVENTORY' in user_input:
        return 'I'

    elif user_input == 'S':
        return 'S'

    elif 'SEND' in user_input:
        return 'S'

    elif 'TOTAL' in user_input:
        return 'T'

    elif user_input == 'T':
        return 'T'

    else:
        print(f'\t{Fore.RED}Unavailable response{Style.RESET_ALL}...')
        user_chose: str = get_available_chooses()
        get_user_response_on_chooses(user_chose)


def run_terminal(version_number: str, default_user_name: str, current_location: str) -> str:
    """
    Generic Intro for Pipe Cleaner including Banner, Intro Sentence, and Local Network checks.
    Then asks for user Role.
    """
    initialize_text_color()
    show_pipe_cleaner_banner()
    show_intro_sentence(version_number, default_user_name, current_location)
    check_local_network()

    user_chose: str = get_available_chooses()
    return get_user_response_on_chooses(user_chose)
