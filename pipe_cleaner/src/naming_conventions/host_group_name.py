"""
Check naming convention standard as agreed upon by everyone in VSE. Checks specifically in the ZT Console Server
Host Groups page in the Host Group name (Pipe Names) fields.

Important for reports and gathering Pipe only data from Console Server.
"""

# Standard feedback for report later. Index feedback is for IndexError
# IndexError happens if character is missing from Pipe name
right_message = 'RIGHT'
wrong_message = 'WRONG'
index_message = 'INDEX'


def check_length_pipe_name(pipe_name: str) -> str:
    """
    Checks length of pipe name based on standard agreed on.
    Important later for different check methods based on length of Pipe name
    :param pipe_name: Host Group name
    :return: status of check
    """
    try:
        if len(pipe_name) == 18:
            return right_message

        elif len(pipe_name) == 20:
            return right_message

        elif len(pipe_name) == 22:
            return right_message

        elif len(pipe_name) == 24:
            return right_message

        else:
            return wrong_message

    except IndexError:
        return index_message


def check_pipe_in_name(pipe_name: str) -> str:
    """
    Check whether Pipe- is in name. Important for gathering
    :param pipe_name:
    :return:
    """
    try:
        name_section: str = pipe_name[0:5]

        if name_section == 'Pipe-':
            return right_message

        else:
            return wrong_message

    except IndexError:
        return 'INDEX'


def check_numbers_in_name(pipe_name: str) -> str:
    """
    Checks to make sure there are three digits following the Pipe-
    Note, the first number of the three digits signify
    :param pipe_name:
    :return:
    """
    try:
        name_section: str = pipe_name[5:8]

        if name_section.isdigit():
            return right_message

        else:
            return wrong_message

    except IndexError:
        return 'INDEX'


def check_paths(pipe_name: str):
    """
    Create path of character analysis based on the number of characters in the pipe name
    :param pipe_name:
    :return:
    """
    try:
        if len(pipe_name) == 18:
            return check_method_1(pipe_name)

        elif len(pipe_name) == 20:
            return check_method_2(pipe_name)

        elif len(pipe_name) == 22:
            return check_method_3(pipe_name)

        elif len(pipe_name) == 24:
            return check_method_4(pipe_name)

        else:
            return wrong_message

    except IndexError:
        return index_message


def check_correct_character(position: int, character_type: str, pipe_name: str) -> list:
    """
    Return type of feedback based on the desired character from agreed standard.
    :param position: index of the pipe name
    :param character_type: type of character of the position ie. letter, digit, etc.
    :param pipe_name:
    :return:
    """
    # Ensures future comparisons
    character_type = character_type.upper()

    if character_type == 'LETTER':
        if pipe_name[position].isalpha() and pipe_name[position].isupper():
            return [position, 'LETTER', 'RIGHT']
        else:
            return [position, 'LETTER', 'WRONG']

    elif character_type == 'DIGIT':
        if pipe_name[position].isdigit():
            return [position, 'DIGIT', 'RIGHT']
        else:
            return [position, 'DIGIT', 'WRONG']

    elif character_type == 'SPACE':
        if pipe_name[position] == ' ':
            return [position, 'SPACE', 'RIGHT']
        else:
            return [position, 'SPACE', 'WRONG']

    elif character_type == '[':
        if pipe_name[position] == '[':
            return [position, '[', 'RIGHT']
        else:
            return [position, '[', 'WRONG']

    elif character_type == '[':
        if pipe_name[position] == '[':
            return [position, '[', 'RIGHT']
        else:
            return [position, '[', 'WRONG']

    elif character_type == '(':
        if pipe_name[position] == '(':
            return [position, '(', 'RIGHT']
        else:
            return [position, '(', 'WRONG']

    elif character_type == ')':
        if pipe_name[position] == ')':
            return [position, ')', 'RIGHT']
        else:
            return [position, ')', 'WRONG']

    elif character_type == 'HYPHEN':
        if pipe_name[position] == '-':
            return [position, 'HYPHEN', 'RIGHT']
        else:
            return [position, 'HYPHEN', 'WRONG']

    elif character_type == 'PERIOD':
        if pipe_name[position] == '.':
            return [position, 'PERIOD', 'RIGHT']
        else:
            return [position, 'PERIOD', 'WRONG']

    else:
        # If character type not present
        return [position, 'INVALID', 'INDEX']


def check_method_1(pipe_name: str) -> list:
    """
    Accounts for 18 characters
    ex. Pipe-616 [0W041-B]
    :param pipe_name:
    :return: results
    """
    return [check_correct_character(8, 'SPACE', pipe_name), check_correct_character(9, '[', pipe_name),
            check_correct_character(10, 'NUMBER', pipe_name), check_correct_character(11, 'LETTER', pipe_name),
            check_correct_character(12, 'NUMBER', pipe_name), check_correct_character(13, 'NUMBER', pipe_name),
            check_correct_character(14, 'NUMBER', pipe_name), check_correct_character(15, 'HYPHEN', pipe_name),
            check_correct_character(16, 'LETTER', pipe_name), check_correct_character(17, ']', pipe_name)]


def check_method_2(pipe_name: str) -> list:
    """
    Accounts for 20 characters
    ex. Pipe-625-A [0W030-B]
    :param pipe_name:
    :return: results
    """
    return [check_correct_character(8, 'HYPHEN', pipe_name), check_correct_character(9, 'LETTER', pipe_name),
            check_correct_character(10, 'SPACE', pipe_name), check_correct_character(11, '[', pipe_name),
            check_correct_character(12, 'NUMBER', pipe_name), check_correct_character(13, 'LETTER', pipe_name),
            check_correct_character(14, 'NUMBER', pipe_name), check_correct_character(15, 'NUMBER', pipe_name),
            check_correct_character(16, 'NUMBER', pipe_name), check_correct_character(17, 'HYPHEN', pipe_name),
            check_correct_character(18, 'LETTER', pipe_name), check_correct_character(19, ']', pipe_name)]


def check_method_3(pipe_name: str) -> list:
    """
    Accounts for 22 characters
    ex. Pipe-735 CPT [0B004-T]
    :param pipe_name: results
    :return: results
    """
    return [check_correct_character(8, 'SPACE', pipe_name), check_correct_character(9, 'LETTER', pipe_name),
            check_correct_character(10, 'LETTER', pipe_name), check_correct_character(11, 'LETTER', pipe_name),
            check_correct_character(12, 'SPACE', pipe_name), check_correct_character(13, '[', pipe_name),
            check_correct_character(14, 'NUMBER', pipe_name), check_correct_character(15, 'LETTER', pipe_name),
            check_correct_character(16, 'NUMBER', pipe_name), check_correct_character(17, 'NUMBER', pipe_name),
            check_correct_character(18, 'NUMBER', pipe_name), check_correct_character(19, 'HYPHEN', pipe_name),
            check_correct_character(20, 'LETTER', pipe_name), check_correct_character(21, ']', pipe_name)]


def check_method_4(pipe_name: str) -> list:
    """
    Accounts for 24 characters
    ex. Pipe-625-A STO [0W030-B]
    :param pipe_name:
    :return: results
    """
    return [check_correct_character(8, 'HYPHEN', pipe_name), check_correct_character(9, 'LETTER', pipe_name),
            check_correct_character(10, 'SPACE', pipe_name), check_correct_character(11, 'LETTER', pipe_name),
            check_correct_character(12, 'LETTER', pipe_name), check_correct_character(13, 'LETTER', pipe_name),
            check_correct_character(14, 'SPACE', pipe_name), check_correct_character(15, '[', pipe_name),
            check_correct_character(16, 'NUMBER', pipe_name), check_correct_character(17, 'LETTER', pipe_name),
            check_correct_character(18, 'NUMBER', pipe_name), check_correct_character(19, 'NUMBER', pipe_name),
            check_correct_character(20, 'NUMBER', pipe_name), check_correct_character(21, 'HYPHEN', pipe_name),
            check_correct_character(22, 'LETTER', pipe_name), check_correct_character(23, ']', pipe_name)]


def main_method(pipe_name: str) -> dict:
    """
    Check main
    :param pipe_name:
    :return:
    """
    return {'length': check_length_pipe_name(pipe_name), 'pipe_in_name': check_pipe_in_name(pipe_name),
            'pipe_number': check_numbers_in_name(pipe_name), 'check_paths': check_paths(pipe_name)}
