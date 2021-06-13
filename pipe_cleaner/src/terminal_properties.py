from colorama import Fore, Style

row_length = 112
space = '  '


class LogTerminalStatement:
    @staticmethod
    def log_finished() -> str:
        print(f'\n\tPipe Cleaner Finished...')

    @staticmethod
    def optional_joke() -> str:
        return f'\n\tRandom Joke (Submit SFW joke via DM/Email/Teams):\n\n  {get_joke()}'

    @staticmethod
    def folder_location(pipe_number) -> str:
        print(f'\n\tPress {Fore.LIGHTBLUE_EX}ENTER{Style.RESET_ALL} to create {pipe_number} folder...', end="")
        input()

    @staticmethod
    def connected_vpn() -> str:
        print('\n\tConnected to GlobalProtect. Correct Network...')

    @staticmethod
    def not_connected_vpn() -> str:
        print('\n\tHello')


class Headers(object):
    """
    Terminal Header w/ Header Title and Section Description
    # TODO - Provide Alpha, Beta States
    """
    blue_start = f'{Fore.BLUE}'
    green_start = f'{Fore.GREEN}'
    style_end = f'{Style.RESET_ALL}'

    def __init__(self, center_statement: str, repeat_character: str, separate: str, left_arrow: str, right_arrow: str):
        self.center_statement = center_statement
        self.repeat_character = repeat_character
        self.separate = separate
        self.left_arrow = left_arrow
        self.right_arrow = right_arrow

    def header_line(self):

        break_length = self.repeat_character * 114

        word_len = len(self.center_statement)
        break_len = len(break_length)

        difference = int(((break_len - word_len) / 2) - 1)

        if 'Feedback' in self.center_statement:
            full = f'{Headers.green_start}{self.center_statement}{Headers.style_end}'
        else:
            full = self.center_statement

        word_break = f'  {difference * self.repeat_character}{self.left_arrow}{self.separate}{full}' \
                     f'{self.separate}{self.right_arrow}{difference * self.repeat_character}'

        return word_break


def blue_font(text: str):
    """
    Turns the default font color to blue
    :return:
    """
    blue_text = f'{Fore.BLUE}{text}{Style.RESET_ALL}'

    return blue_text


def terminal_header_section(header_title: str, header_explanation: str):
    """
    Reusable for for title and explanation of a terminal section
    :param header_title: Name of the terminal section
    :param header_explanation: Explanation of the terminal section
    :return:
    """
    characters = '='

    title_length = len(header_title)
    explanation_length = len(header_explanation)

    # Diff = Difference
    title_diff = int((row_length - title_length) / 2)
    explanation_diff = int((row_length - explanation_length) / 2)
    lower_diff = int(row_length + 4)

    header = f'{space}{title_diff * f"{characters}"}[ {header_title} ]{title_diff * f"{characters}"}'

    explanation = f'{space}{explanation_diff * " "}  {header_explanation}  ' \
                  f'{explanation_diff * " "}'

    lower_bar = f'{space}{lower_diff * f"{characters}"}'

    print(f'\n{header}')
    print(f'{explanation}')
    print(f'{lower_bar}')
    print('\n')


def intro_section(version: str, user_name: str, current_location: str):
    """
    Reusable for for title and explanation of a terminal section
    :param version: Pipe Version
    :param user_name: name via network
    :param current_location: site using Pipe Cleaner
    :return:
    """
    characters = "-"
    # user_name: str = user_name.replace('.', ' ').title()

    header_1_content: str = 'Veritas Services & Engineering'
    header_2_content: str = 'Developer: Joe Ton'
    header_3_content: str = 'Project Manager: Bruce Saari'
    header_4_content: str = f'User: {user_name}'
    header_5_content: str = f'Location: {current_location}'
    header_6_content: str = f'Version: {version}'

    header_1_length = len(header_1_content)
    header_2_length = len(header_2_content)
    header_3_length = len(header_3_content)
    header_4_length = len(header_4_content)
    header_5_length = len(header_5_content)
    header_6_length = len(header_6_content)

    # Diff = Difference
    header_1_diff = int((row_length - header_1_length) / 2)
    header_2_diff = int((row_length - header_2_length) / 2) + 2
    header_3_diff = int((row_length - header_3_length) / 2) + 2
    header_4_diff = int((row_length - header_4_length) / 2) + 2
    header_5_diff = int((row_length - header_5_length) / 2) + 2
    header_6_diff = int((row_length - header_6_length) / 2) + 2

    header_1 = f"{space}{header_1_diff * f'{characters}'}[ {header_1_content} ]{header_1_diff * f'{characters}'}"
    header_2 = f"{space}{header_2_diff * ' '}{header_2_content}{header_2_diff * ' '}"
    header_3 = f"{space}{header_3_diff * ' '}{header_3_content}{header_3_diff * ' '}"
    header_4 = f"{space}{header_4_diff * ' '}{header_4_content}{header_4_diff * ' '}"
    header_5 = f"{space}{header_5_diff * ' '}{header_5_content}{header_5_diff * ' '}"
    header_6 = f"{space}{header_6_diff * ' '}{header_6_content}{header_6_diff * ' '}"

    print(f"{header_1}")
    print(f"{header_2}")
    print(f"{header_3}\n")
    print(f"{header_4}")
    print(f"{header_5}")
    print(f"{header_6}\n")

    print(f"\tChecking system requirements:")
