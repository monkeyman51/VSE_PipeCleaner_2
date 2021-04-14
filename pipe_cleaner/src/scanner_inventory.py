import pandas as pd
from pipe_cleaner.src.terminal_properties import break_line
from colorama import Fore, Style


def access_inventory(file_path: str, sheet_name: str):
    """
    Access inventory via local file within Pipe Cleaner.
    WARNING: Must update local excel file in order to be up to date
    :param sheet_name: commodity inventory
    :param file_path: file path of inventory maintained by Traci, Bruce, or inventory person
    :return:
    """

    commodity_inventory = pd.read_excel(f'{file_path}', sheet_name=f'{sheet_name}')
    commodity_inventory.to_csv('commodity_inventory.csv', index=False)

    data_frame = pd.read_csv('commodity_inventory.csv')

    return data_frame


def cover_page_from_inventory():
    # sheet = 'cover_page'

    commodity_inventory = pd.read_excel(f'input/crd.xlsx', sheet_name='Cover Page')
    commodity_inventory.to_csv('cover_page.csv', index=False)

    data_frame = pd.read_csv('cover_page.csv')

    return data_frame


def temp_inventory_message() -> None:
    print('  \n Inventory Report:')
    print('   * Software Catalog (Anomalies):')
    print('     - C2360.BS.3A13')
    print('     - C2160.BS.3A13 \n')
    print('   * Hardware Catalog (Anomalies):')
    print('     - None (Still in development)')


def check_part_and_model_numbers(data_frame, number):
    model_number_df = data_frame['Model Number'].to_list()
    part_number_df = data_frame['MFG Part Number'].to_list()
    item_type_df = data_frame['Item Type'].to_list()
    part_number_compare = dict(zip(part_number_df, item_type_df))
    model_number_compare = dict(zip(model_number_df, item_type_df))

    if number in part_number_compare:
        return part_number_compare[number]
    elif number in model_number_compare:
        return model_number_compare[number]
    else:
        print('  Part/Model Number not in Hardware Inventory...')


def get_type(data_frame, part_number):
    """
    :param data_frame:
    :param part_number:
    :return:
    """
    item_type_df = data_frame['Item Type'].to_list()

    part_number_compare = dict(zip(part_number, item_type_df))

    return part_number_compare


def get_supplier(data_frame, part_number):
    """
    :param data_frame:
    :param part_number:
    :return:
    """
    item_type_df = data_frame['Item Supplier'].to_list()

    part_number_compare = dict(zip(part_number, item_type_df))

    return part_number_compare


def get_description(data_frame, part_number):
    """
    :param data_frame:
    :param part_number:
    :return:
    """
    item_type_df = data_frame['Description'].to_list()

    part_number_compare = dict(zip(part_number, item_type_df))
    # check_description(part_number_compare)

    return part_number_compare


def get_quantity(data_frame, part_number):
    """
    :param data_frame:
    :param part_number:
    :return:
    """
    item_type_df = data_frame['Actual Qty available for use'].to_list()

    part_number_compare = dict(zip(part_number, item_type_df))

    return part_number_compare


# def check_description(number_to_description: dict): # TODO
#     """
#     Checking for naming convention in description
#     :param number_to_description:
#     :return:
#     """
#     description_parts = []
#
#     for part_number, description in enumerate(number_to_description):
#         description_parts.clear()
#
#
#         # Check Empty
#         if description == '' or description is None:
#             print(f'   WARNING: {Fore.RED}{part_number} - Missing Description Field{Style.RESET_ALL}')
#
#         elif


def main_method() -> dict:
    """
    Access and scrub data for Excel, Word output for Pipe Cleaner
    :return:
    """
    inventory: dict = {}

    inventory_file_path: str = 'settings/Kirkland_Inventory.xlsx'
    commodity_inventory: str = 'Commodity Inventory'

    print(f'\n  {break_line("Hardware and Software Inventory", "=", " ", "[", "]")}  ')
    print(f'  {break_line("Getting and checking Inventory", " ", " ", " ", " ")}')
    print(f'  {break_line("", "=", "=", "=", "=")}  \n')

    print(f'  NOTE: Software Inventory - Temporarily Off')
    print(f'  NOTE: Hardware Inventory - Copied from Teams | Scanned locally\n')

    print(f'  Extracting and scrubbing {Fore.GREEN}hardware inventory data{Style.RESET_ALL}...\n')

    # Get Data Frame and Part Number for keys in dictionaries
    # ie. item types, suppliers, description, quantity in stock
    data_frame = access_inventory(inventory_file_path, commodity_inventory)
    part_number_df = data_frame['MFG Part Number'].to_list()

    # MFG Part Number to **
    item_type = get_type(data_frame, part_number_df)
    item_supplier = get_supplier(data_frame, part_number_df)
    description = get_description(data_frame, part_number_df)
    quantity = get_quantity(data_frame, part_number_df)

    inventory['item_type'] = item_type
    inventory['item_supplier'] = item_supplier
    inventory['description'] = description
    inventory['quantity'] = quantity

    return inventory

