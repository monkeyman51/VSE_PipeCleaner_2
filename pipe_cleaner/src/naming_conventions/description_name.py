"""
Check naming convention standard as agreed upon by everyone in VSE. Checks specifically in the ZT Console Server
Host Groups page in the Description fields.

Important for reports and gathering Pipe only data from Console Server.
"""

# Standard feedback for report later
right_message = 'RIGHT'
wrong_message = 'WRONG'


def check_hyphen(description_name: str) -> str:
    """
    Checks if description field has hyphen separating the suppliers and the SKU types.
    :param description_name: description field in Host Group page
    :return: RIGHT or WRONG
    """
    try:
        if ' - ' in description_name:
            return right_message
        else:
            return wrong_message
    except TypeError:
        return wrong_message


def check_opening_square_bracket(description_name: str) -> str:
    """
    Checks for opening square brackets for start of SKU sub types
    :param description_name: description field in Host Group page
    :return: RIGHT or WRONG
    """
    try:
        if '[' in description_name:
            return right_message
        else:
            return wrong_message
    except TypeError:
        pass


def check_closing_square_bracket(description_name: str) -> str:
    """
    Checks for closing square brackets for end of SKU sub types
    :param description_name: description field in Host Group page
    :return: RIGHT or WRONG
    """
    try:
        if ']' in description_name:
            return right_message
        else:
            return wrong_message
    except TypeError:
        pass


def main_method(description_name: str) -> list:
    """
    Checks description field for agreed standard
    :param description_name:
    :return: flags wrong if doesn't go with standard
    """
    return [check_hyphen(description_name), check_opening_square_bracket(description_name),
            check_closing_square_bracket(description_name)]
