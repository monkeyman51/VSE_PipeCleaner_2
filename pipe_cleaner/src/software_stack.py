import os
import win32api
import win32net

path_directory = r'\\172.30.1.100\pxe\Kirkland_Lab\Users\Joe_Ton\pipe_cleaner_test'

bios_scan = []
bmc_scan = []

# Files/Folders that don't follow normal naming conventions.
anomalies = []

bios_properties = ('.BS.', '12', '16')
bmc_properties = ('.BC.', '12', '15')


def enter_user_password(user_name, password) -> None:
    use_dict = {'remote': f'{path_directory}',
                'username': user_name,
                'password': password}
    win32net.NetUseAdd(None, 2, use_dict)


def check_commodity_software() -> list:
    """
    Checks if BIOS and BMC are in the folders.
    :return: list of BIOS and BMC
    """
    # Files that match inventory items.
    inventory_items: list = []

    try:
        for file in os.listdir(path_directory):
            if '.BS.' in file:  # Grabs BIOS
                inventory_items.append(file)
            elif '.BC.' in file:  # Grabs BMC
                inventory_items.append(file)
            else:
                anomalies.append(file)  # Place extra in anomalies
    except OSError:
        print('  Access Denied. Need to enter credentials to access Z: Drive files. \n')
        user_name = input('  Z:Drive User Name: ')
        password = input('  Z:Drive Password: ')
        enter_user_password(user_name, password)
        for file in os.listdir(path_directory):
            if '.BS.' in file:  # Grabs BIOS
                inventory_items.append(file)
            elif '.BC.' in file:  # Grabs BMC
                inventory_items.append(file)
            else:
                anomalies.append(file)  # Place extra in anomalies

    return inventory_items


foo = check_commodity_software()
print(anomalies)
