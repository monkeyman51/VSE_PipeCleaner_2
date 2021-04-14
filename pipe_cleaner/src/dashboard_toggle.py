from xlrd import open_workbook
from pipe_cleaner.src.sheet_3 import cleans_xlrd_cell


def get_component_toggles(document_filepath: str, component_type: str) -> dict:
    """
    Get component toggle information from all_toggles.xlsx
    :param component_type:
    :param document_filepath:
    :return:
    """
    request_type_sheet = open_workbook(document_filepath).sheet_by_name(component_type)

    toggle_data: dict = {}

    for part in range(6, request_type_sheet.nrows):

        # Converts to string type for later parsing
        key = str(request_type_sheet.cell(part, 1))
        value = str(request_type_sheet.cell(part, 2))

        # Cleans xlrd extra information ie. text and number
        cleaned_key = cleans_xlrd_cell(key)
        cleaned_value = cleans_xlrd_cell(value)

        # Prepare key
        prepared_key = cleaned_key.replace(' ', '_').lower()

        if 'empty:' in value == '' or 'END' in value:
            pass
        elif prepared_key == '' or cleaned_key == '':
            pass
        else:
            toggle_data[prepared_key] = cleaned_value

    return toggle_data


def main_method():
    get_components: list = ['DIMM',
                            'NVME',
                            'SSD',
                            'HDD']

    all_toggle_data: dict = {}

    toggle_file_path = 'settings/all_toggles.xlsx'
    for component_type in get_components:
        toggle_data: dict = get_component_toggles(toggle_file_path, component_type)
        all_toggle_data[component_type] = toggle_data

    return all_toggle_data
