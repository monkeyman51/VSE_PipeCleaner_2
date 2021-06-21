"""
Figure out naming convention.
"""
from time import strftime
from pipe_cleaner.src.log_database import access_database_document


def get_year_for_form_number() -> str:
    """
    Form number number requires year for the 2nd and 3rd index position in the form number.
    """
    return strftime('%m/%d/%Y')[8:10]


def convert_site_name_to_site_number_for_form_number(site_name: str) -> str:
    """
    Convert site name into site number for the first index position in the form number.
    """
    if not site_name:
        return '0'

    elif 'KIRKLAND' in site_name.upper():
        return '0'

    elif 'TAIWAN' in site_name.upper():
        return '1'

    else:
        return '0'


def combine_site_and_year_for_form_number(site_number: str, year: str) -> str:
    """
    Form number requires location and time for documentation purposes.  Combining them together for future reference
    will be important for consistency in standardizing the form.
    :param site_number: Tells location
    :param year: Tells time
    :return: ex. 021
    """
    if not site_number:
        return f'N{year}'

    elif not year:
        return f'{site_number}NN'

    elif not site_number and not year:
        return f'NNN'

    else:
        return f'{site_number}{year}'


def set_level_1_schema_for_transaction_database() -> dict:
    """
    Need to establish form number standard on key-value pairs for consistent calls to database later.
    """
    return {'_id': '',
            'request': {},
            'update': {}}


def set_level_2_schema_for_transaction_database() -> dict:
    """
    Establishes main contain required for request and update information.
    """
    return {'basic': {'date_entry': '',
                      'time_entry': '',
                      'verified_by': '',
                      'current_site': '',
                      'form_number': '',
                      'pipe_cleaner_version': ''},
            'fields': {'approved_by': '',
                       'part_number': '',
                       'start_location': '',
                       'end_location': '',
                       'task_name': '',
                       'notes': '',
                       'pipe': '',
                       'purchase_order': ''},
            'status': ''}


def main_method_for_new_form_number() -> None:
    """
    Need to find create a process to identify previous form numbers within database then create a new unique
    form number.  This form number is for update and
    """
    pass
