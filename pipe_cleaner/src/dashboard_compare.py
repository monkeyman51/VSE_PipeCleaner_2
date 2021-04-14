"""
Compare data from TRRs to Console Server
"""

from pipe_cleaner.src.dashboard_all_issues import check_missing

import json

all_issues: list = []
total_checks: list = []
missing_tally: list = []
mismatch_tally: list = []
vse_issue_tally: list = []


def ticket_vs_system(ticket_value: str, system_value: str) -> str:
    """
    Compares string value from ADO to Console Server per component given.
    :param ticket_value: Component's value based from ADO
    :param system_value: Same component but value based from Console Server
    :return:
    """
    upper_ticket = ticket_value.upper()
    upper_system = system_value.upper()

    if upper_ticket is None or upper_ticket == '' or upper_system is None or upper_system == '':
        return 'Missing'
    elif upper_ticket != upper_system:
        return 'Mismatch'
    elif upper_ticket == upper_system:
        return 'Match'
    # For any anomalies
    elif 'ERRONEOUS' in upper_ticket or 'ERRONEOUS' in upper_system:
        return 'Other'
    else:
        return 'Other'


def scrub_server_bios(raw_server_bios: str) -> str:
    """
    Ensure clean data for later comparison
    :param raw_server_bios: data from Console Server
    :return: clean data
    """
    parse_server_bios = raw_server_bios.replace('-[', '').replace(']-', '')

    return parse_server_bios


def scrub_server_bmc(raw_server_bmc: str) -> str:
    """
    Ensure clean data for later comparison
    :param raw_server_bmc: data from Console Server
    :return: clean data
    """
    if len(raw_server_bmc) == 4 and raw_server_bmc[1] == '.':
        parse_bmc = raw_server_bmc.replace('.', '')
        return parse_bmc
    elif len(raw_server_bmc) == 16 and raw_server_bmc[1] == '.':
        parse_bmc = raw_server_bmc.replace('.', '')
        return parse_bmc
    else:
        return raw_server_bmc


def scrub_server_cpld(raw_server_cpld: str) -> list:
    """
    Ensure clean data for later comparison
    :param raw_server_cpld: data from Console Server
    :return: clean data
    """
    unique_characters: list = []

    clean_cpld = raw_server_cpld.replace('000000', '').replace('000000', '').replace('000', '')

    for character in clean_cpld:
        if character.isdigit() or character.isalpha():
            unique_characters.append(character)

    check_empty = list(filter(None, unique_characters))

    return check_empty


def scrub_server_os(raw_server_os: str) -> str:
    """
    Ensure clean data for later comparison
    :param raw_server_os: data from Console Server
    :return: clean data
    """
    # Goal = ex. 17763

    if len(raw_server_os) == 10 and raw_server_os[2] == '.' and raw_server_os[4] == '.' and \
            raw_server_os[-5:].isdigit():
        parse_os: str = raw_server_os[-5:]
        return parse_os.strip()

    return raw_server_os


def scrub_server_tpm(raw_server_tpm: str) -> str:
    """
    Ensure clean data for later comparison
    :param raw_server_tpm: data from Console Server
    :return: clean data
    """
    # goal = 7.63

    if len(raw_server_tpm) == 14 and raw_server_tpm[1] == '.':
        return raw_server_tpm[0:2]
    elif len(raw_server_tpm) == 12 and raw_server_tpm[1] == '.':
        return raw_server_tpm[0:2]
    elif len(raw_server_tpm) == 4 and raw_server_tpm[1] == '.':
        return raw_server_tpm[0:2]
    else:
        return raw_server_tpm


def scrub_machine_name(raw_machine_name: str) -> str:
    """
    Ensure clean data for later comparison
    :param raw_machine_name: data from Console Server
    :return: clean data
    """
    return raw_machine_name


def scrub_checked_out_to(raw_checked_out_to: str) -> str:
    """
    Ensure clean data for later comparison
    :param raw_checked_out_to: data from Console Server
    :return: clean data
    """
    if check_missing(raw_checked_out_to) == 'None':
        return 'None'
    else:
        return raw_checked_out_to.replace('.', ' ').title()


def scrub_server_ticket(raw_server_ticket: str) -> str:
    """
    Ensure clean data for later comparison
    :param raw_server_ticket: data from Console Server
    :return: clean data
    """
    if raw_server_ticket.isdigit() is True:
        return raw_server_ticket
    elif check_missing(raw_server_ticket) == 'None':
        return 'None'


def scrub_ticket_bios(raw_ticket_bios):
    """
    Ensure clean data for later comparison
    :param raw_ticket_bios: data from ticket in ADO
    :return: clean data
    """
    return raw_ticket_bios.strip()


def scrub_ticket_bmc(raw_ticket_bmc: str):
    """
    Ensure clean data for later comparison
    :param raw_ticket_bmc: data from ticket in ADO
    :return: clean data
    """

    try:
        clean_bmc = clean_raw_bmc(raw_ticket_bmc)

        if len(clean_bmc) == 16 and '.BC.' in clean_bmc and clean_bmc[-3:] == '.00' and \
                clean_bmc[5] == '.' and clean_bmc[8] == '.':
            if clean_bmc[-1:] == '.':
                cleaner_bmc: str = clean_bmc[:-1]
                parsed_bmc = cleaner_bmc.split('.')[2]
                final_bmc = parsed_bmc[-3:]
                return final_bmc.strip()
            else:
                parsed_bmc = clean_bmc.split('.')[2]
                final_bmc = parsed_bmc[-3:]
                return final_bmc.strip()

        elif len(clean_bmc) == 13 and '.BC.' in clean_bmc and \
                clean_bmc[5] == '.' and clean_bmc[8] == '.':
            parsed_bmc = clean_bmc.split('.')[2]
            final_bmc = parsed_bmc[-3:]
            return final_bmc.strip()

        elif len(clean_bmc) == 11 and '.BC.' in clean_bmc and clean_bmc[-3:] == '.00' and \
                clean_bmc[3] == '.' and clean_bmc[5] == '.':
            parsed_bmc = clean_bmc.split('.')[2]
            final_bmc = parsed_bmc[-3:]
            return final_bmc.strip()
        else:
            return clean_bmc.strip()
    except AttributeError:
        pass


def clean_raw_bmc(raw_ticket_bmc):
    clean_bmc = raw_ticket_bmc.replace(' ', '').strip()
    if clean_bmc[-1] == '.':
        return clean_bmc[0:-1]
    else:
        return clean_bmc


def scrub_ticket_cpld(raw_ticket_cpld: str) -> list:
    """
    Ensure clean data for later comparison
    :param raw_ticket_cpld: data from ticket in ADO
    :return: clean data
    """
    unique_characters: list = []

    try:
        clean_cpld = raw_ticket_cpld.replace('000000', '').replace('000000', '').replace('000', '').replace('.00.', '')

        for character in clean_cpld:
            if character.isdigit() or character.isalpha():
                parse_character = str(character).replace('â', '').replace('Â', '')
                unique_characters.append(parse_character)

        check_empty = list(filter(None, unique_characters))

        return check_empty
    except AttributeError:
        pass


def scrub_ticket_os(raw_ticket_os: str):
    """
    Ensure clean data for later comparison
    :param raw_ticket_os: data from ticket in ADO
    :return: clean data
    """
    # ex. 17763
    upper_ticket_os = raw_ticket_os.upper()

    if raw_ticket_os[0:3].isdigit() and 'DATACENTER' in upper_ticket_os and 'VERSION' in upper_ticket_os and \
            '(' in raw_ticket_os and ')' in raw_ticket_os:
        parse_ticket_os = raw_ticket_os.replace(')', '')[-5:]
        return parse_ticket_os
    elif '2019' in raw_ticket_os and 'DATACENTER' in upper_ticket_os and '-' in upper_ticket_os:
        return '17763'
    elif '2019' in raw_ticket_os and 'DATACENTER' in upper_ticket_os:
        return '17763'
    elif '2019' in raw_ticket_os:
        return '17763'

    return raw_ticket_os.strip()


def scrub_ticket_tpm(raw_ticket_tpm: str):
    """
    Ensure clean data for later comparison
    :param raw_ticket_tpm: data from ticket in ADO
    :return: clean data
    """
    # goal = ex. 5.62
    clean_tpm = raw_ticket_tpm.replace(' ', '')

    if len(clean_tpm) == 11 and clean_tpm[1] == '.':
        return clean_tpm[0:3]
    elif len(clean_tpm) == 14 and clean_tpm[1] == '.':
        return clean_tpm[0:3]
    elif len(clean_tpm) == 12 and clean_tpm[1] == '.':
        return clean_tpm[0:3]
    elif len(clean_tpm) == 4 and clean_tpm[1] == '.':
        return clean_tpm[0:3]
    else:
        return clean_tpm


def scrub_ticket_boot_drive(raw_ticket_boot_drive: str) -> str:
    """
    Make sure clean data for later comparison
    :param raw_ticket_boot_drive:
    :return:
    """
    return raw_ticket_boot_drive.strip()


def scrub_ticket_toolkit(raw_ticket_toolkit: str) -> str:
    """
    Make sure clean data for later comparison
    :param raw_ticket_toolkit:
    :return:
    """
    return raw_ticket_toolkit.strip()


def scrub_ticket_reference_test_plans(raw_ticket_scrub_ticket_reference_test_plans: str):
    """
    Make sure clean data for later comparison
    :param raw_ticket_scrub_ticket_reference_test_plans:
    :return:
    """
    return raw_ticket_scrub_ticket_reference_test_plans


def scrub_ticket_chipset_driver(raw_ticket_chipset_driver: str):
    """
    Make sure clean data for later comparison
    :param raw_ticket_chipset_driver:
    :return:
    """
    return raw_ticket_chipset_driver


def scrub_ticket_server_processors(raw_ticket_server_processors: str):
    """
    Make sure clean data for later comparison
    :param raw_ticket_server_processors:
    :return:
    """
    return raw_ticket_server_processors


def scrub_ticket_fpga_release(raw_ticket_fpga_release: str):
    """
    Make sure clean data for later comparison
    :param raw_ticket_fpga_release:
    :return:
    """
    return raw_ticket_fpga_release


def scrub_ticket_nic_firmware(raw_ticket_nic_firmware: str):
    """
    Make sure clean data for later comparison
    :param raw_ticket_nic_firmware:
    :return:
    """
    return raw_ticket_nic_firmware


def scrub_ticket_nic_pxe(raw_ticket_nic_pxe: str):
    """
    Make sure clean data for later comparison
    :param raw_ticket_nic_pxe:
    :return:
    """
    return raw_ticket_nic_pxe


def scrub_ticket_nic_uefi(raw_ticket_nic_uefi: str):
    """
    Make sure clean data for later comparison
    :param raw_ticket_nic_uefi:
    :return:
    """
    return raw_ticket_nic_uefi


def scrub_ticket_nic_driver(raw_ticket_nic_driver: str):
    """
    Make sure clean data for later comparison
    :param raw_ticket_nic_driver:
    :return:
    """
    return raw_ticket_nic_driver


def scrub_ticket_rm_firmware(raw_ticket_rm_firmware: str):
    """
    Make sure clean data for later comparison
    :param raw_ticket_rm_firmware:
    :return:
    """
    return raw_ticket_rm_firmware


def get_ticket_ssd_data(table_data: dict) -> dict:
    """
    Get SSD data for later comparison
    :param table_data:
    :return:
    """
    if table_data == 'None' or table_data == '':
        return {}
    else:
        ssd_table_data: list = [str(table_data.get('qcl_ssd_1', '')), str(table_data.get('qcl_ssd_2', '')),
                                str(table_data.get('qcl_ssd_3', '')), str(table_data.get('qcl_ssd_4', '')),
                                str(table_data.get('qcl_ssd_5', '')), str(table_data.get('qcl_ssd_6', '')),
                                str(table_data.get('qcl_ssd_7', '')), str(table_data.get('qcl_ssd_8', '')),
                                str(table_data.get('qcl_ssd_9', '')), str(table_data.get('qcl_ssd_10', ''))]
        return clean_data(ssd_table_data)


def get_ticket_nvme_data(table_data: dict) -> dict:
    """
    Get NVMe data for later comparison
    :param table_data:
    :return:
    """
    if table_data == 'None' or table_data == '':
        return {}
    else:
        nvme_table_data: list = [str(table_data.get('qcl_nvme_1', '')), str(table_data.get('qcl_nvme_2', '')),
                                 str(table_data.get('qcl_nvme_3', '')), str(table_data.get('qcl_nvme_4', '')),
                                 str(table_data.get('qcl_nvme_5', '')), str(table_data.get('qcl_nvme_6', '')),
                                 str(table_data.get('qcl_nvme_7', '')), str(table_data.get('qcl_nvme_8', '')),
                                 str(table_data.get('qcl_nvme_9', '')), str(table_data.get('qcl_nvme_10', ''))]

        return clean_data(nvme_table_data)


def get_ticket_hdd_data(table_data: dict) -> dict:
    """
    Get HDD data for later comparison
    :param table_data:
    :return:
    """
    if table_data == 'None' or table_data == '':
        return {}
    else:
        hdd_table_data: list = [str(table_data.get('qcl_hdd_1', '')), str(table_data.get('qcl_hdd_2', '')),
                                str(table_data.get('qcl_hdd_3', '')), str(table_data.get('qcl_hdd_4', '')),
                                str(table_data.get('qcl_hdd_5', '')), str(table_data.get('qcl_hdd_6', '')),
                                str(table_data.get('qcl_hdd_7', '')), str(table_data.get('qcl_hdd_8', '')),
                                str(table_data.get('qcl_hdd_9', '')), str(table_data.get('qcl_hdd_10', ''))]

        return clean_data(hdd_table_data)


def manual_clean_data(data: list):
    """
    Manual remove information from Part Number and Version
    :param data:
    :return:
    """
    container_1: list = []
    excess_parts: list = ['3.,', '3.84TB,', 'HYNIX,', 'TOSHIBA', 'SAMSUNG', 'SKHYNIX', 'PM9833.', '84TB',
                          'WESTERNDIGITAL', 'MICRON5200PRO', 'MICRON', 'PE6010', 'PE6011', 'SEAGATEEXOS',
                          'SEAGATE', 'PM963', 'PM983', 'PE6011', 'PE4010', '32GB', '960GB', '12TB', 'P4511',
                          'INTEL', 'SSD', 'AMD', 'HDD', 'NVME', 'DIMM', 'PM883', 'SHYNIX', 'HGST', 'HYNIX']

    # If data doesn't have anything
    # Doesn't have to go through iteration wasting time
    if len(data) == 0:
        return data
    else:
        for sql_part in data:
            single_container: list = []
            initial = 0

            while initial < len(excess_parts):
                excess_part = excess_parts[initial]

                if len(single_container) == 0:
                    if excess_part in sql_part:
                        clean = str(sql_part).replace(excess_part, '')
                        single_container.append(clean)
                    else:
                        single_container.append(sql_part)

                elif len(single_container) == 1:
                    if excess_part in single_container[0]:
                        clean = str(single_container[0]).replace(excess_part, '')
                        single_container.clear()
                        single_container.append(clean)

                else:
                    pass
                initial += 1
            container_1.append(single_container[0])

    return container_1


def separate_firmware_data(all_qualified_components: list) -> dict:
    """
    Seprates the firmware from the model numbers / part numbers.
    :param all_qualified_components:
    :return:
    """
    container: dict = {}
    model_numbers: list = []
    firmware_versions: list = []

    if len(all_qualified_components) == 0:
        return container
    else:
        for qualified_component in all_qualified_components:
            if '/FW;' in qualified_component:
                model_number = str(qualified_component).split('/FW;')[0]
                firmware = str(qualified_component).split('/FW;')[-1]
                model_numbers.append(model_number)
                firmware_versions.append(firmware)

            elif '/FW:' in qualified_component:
                model_number = str(qualified_component).split('/FW:')[0]
                firmware = str(qualified_component).split('/FW:')[-1]
                model_numbers.append(model_number)
                firmware_versions.append(firmware)
            else:
                model_numbers.append(qualified_component)

    container['model_numbers'] = model_numbers
    container['firmware_versions'] = firmware_versions

    return container


def separate_model_numbers(separated_firmware: list):
    """

    :param separated_firmware:
    :return:
    """
    container: list = []

    # Remove Empties
    scrub_for_empty: list = [x for x in separated_firmware if x.strip()]

    if len(scrub_for_empty) == 0:
        return container
    else:
        for qualified_component in scrub_for_empty:
            if '(#)' in qualified_component:
                cleaner_data = str(qualified_component).replace('(#)', '')
                container.append(cleaner_data)

            elif '(#' in qualified_component:
                cleaner_data = str(qualified_component).replace(')', '')
                model_number = str(cleaner_data).split('(#')[0]
                part_number = str(cleaner_data).split('(#')[1]
                container.append(model_number)
                container.append(part_number)
            else:
                container.append(qualified_component)

        return container


def clean_data(data: list) -> dict:
    """
    Clean DIMM, NVMe, SSD, HDD
    :param data:
    :return:
    """
    # Remove Unnecessary Data
    none_text: list = [str(x).replace('None', '') for x in data]
    none_upper: list = [str(x).replace('NONE', '') for x in none_text]
    double_hyphen: list = [str(x).replace('--', '') for x in none_upper]

    # Remove Empty Data - For Reducing Iteration load
    remove_empty: list = [x for x in double_hyphen if x.strip()]
    # Remove Spaces - Easier to separate /FW:
    remove_spaces: list = [str(x).replace(' ', '') for x in remove_empty]
    # # Remove based on non-model or non-firmware parts
    unnecessary: list = manual_clean_data(remove_spaces)

    # Separate Firmware
    separated_firmware: dict = separate_firmware_data(unnecessary)  # TODO - TESTING
    # Unpack Model Numbers and Firmware Versions
    model_numbers: list = separated_firmware.get('model_numbers')
    firmware_versions: list = separated_firmware.get('firmware_versions')

    model_numbers = []
    firmware_versions = []

    # Separate Model Number from Part Number
    separated_numbers: list = separate_model_numbers(model_numbers)

    return {'part_numbers': separated_numbers, 'firmware_versions': firmware_versions}


def get_ticket_dimm_data(table_data: dict):
    """
    Get DIMM data for later comparison
    :param table_data:
    :return:
    """

    if table_data == 'None' or table_data == '':
        return 'None'
    else:
        dimm_data: list = [str(table_data.get('qcl_dimm_1', '')), str(table_data.get('qcl_dimm_2', '')),
                           str(table_data.get('qcl_dimm_3', '')), str(table_data.get('qcl_dimm_4', '')),
                           str(table_data.get('qcl_dimm_5', '')), str(table_data.get('qcl_dimm_6', '')),
                           str(table_data.get('qcl_dimm_7', '')), str(table_data.get('qcl_dimm_8', '')),
                           str(table_data.get('qcl_dimm_9', '')), str(table_data.get('qcl_ssd_10', ''))]
        return clean_data(dimm_data)


def get_system_data(component: str, system_data) -> list:
    """

    """
    container: list = []
    # Account for capitalization in code
    component_upper = component.upper()
    component_lower = component.lower()

    clean_system_data = str(system_data).replace("'", '"')
    system_json = json.loads(clean_system_data)

    # testing = 'MODEL'
    # if component == testing and len(system_json) != 0:
    #     print(f'System Data ( {testing} ): {system_data}')
    #     print(f'System JSON ( {testing} ): {system_json}')

    if len(system_json) != 0:
        for unique_data in system_json:
            if 'COUNT' in component_upper:
                container.append(unique_data.get(component_lower, 'None'))

            elif 'MODEL' in component_upper:
                unique_data.get(component_lower, 'None')
                container.append(unique_data.get(component_lower, 'None'))

            elif 'FIRMWARE' in component_upper:
                container.append(unique_data.get(component_lower, 'None'))

            elif 'SIZE' in component_upper:
                container.append(unique_data.get(component_lower, 'None'))

            elif 'PART' in component_upper:
                container.append(unique_data.get(component_lower, 'None'))
    else:
        return container

    # if component == testing:
    #     print(f'System GET  ( {testing} ): {container}\n')

    return container


def check_reference_test_plans(ticket_table_data: dict) -> str:
    """
    Check to see if reference test plans are misspelled which conflict with getting the data from the table data
    :param ticket_table_data:
    :return:
    """
    possible_reference_test_plans: list = []
    # Different ways Reference Test Plans are written
    call_01: str = ticket_table_data.get('reference_test_plans')
    call_02: str = ticket_table_data.get('reference_test_plan')
    call_03: str = ticket_table_data.get('reference_testplans')
    call_04: str = ticket_table_data.get('reference_testplan')

    possible_reference_test_plans.append(call_01)
    possible_reference_test_plans.append(call_02)
    possible_reference_test_plans.append(call_03)
    possible_reference_test_plans.append(call_04)

    try:
        return possible_reference_test_plans[0]
    except IndexError:
        return 'None'


def clean_pipe_data(pipe_name: str, pipe_data: dict, ado_data: dict):
    """
    Get clean relevant information from system data from pipe for later comparison.
    :param pipe_name:
    :param ado_data:
    :param pipe_data:
    :return:
    """
    all_pipe_data: dict = {}

    # Tech vs TRR
    total_tally_count: int = 0
    match_tally_count: int = 0
    mismatch_tally_count: int = 0
    missing_tally_count: int = 0
    vse_tally_count: int = 0
    other_tally_count: int = 0

    # Ticket (TRR) Review
    ticket_total_tally: int = 0
    ticket_match_tally: int = 0
    ticket_missing_tally: int = 0

    for index, system in enumerate(pipe_data, start=1):

        if system == 'pipe_inventory':
            pass

        else:
            system_data: dict = pipe_data[system]

            raw_server_ticket: str = system_data.get('ticket').strip()
            raw_server_bios: str = system_data.get('server_bios')
            raw_server_bmc: str = system_data.get('server_bmc')
            raw_server_cpld: str = system_data.get('server_cpld')
            raw_server_os: str = system_data.get('server_os')
            raw_server_tpm: str = system_data.get('server_tpm')
            raw_machine_name: str = system_data.get('machine_name')
            raw_checked_out_to: str = system_data.get('checked_out_to')

            raw_unique_disks_count: list = get_system_data('COUNT', system_data.get('unique_disks'))
            raw_unique_disks_model: list = get_system_data('MODEL', system_data.get('unique_disks'))
            raw_unique_disks_firmware: list = get_system_data('FIRMWARE', system_data.get('unique_disks'))
            raw_unique_nvmes_count: list = get_system_data('COUNT', system_data.get('unique_nvmes'))
            raw_unique_nvmes_model: list = get_system_data('MODEL', system_data.get('unique_nvmes'))
            raw_unique_nvmes_firmware: list = get_system_data('FIRMWARE', system_data.get('unique_nvmes'))
            raw_unique_dimms_count: list = get_system_data('COUNT', system_data.get('unique_dimms'))
            raw_unique_dimms_size: list = get_system_data('SIZE', system_data.get('unique_dimms'))
            raw_unique_dimms_part: list = get_system_data('PART', system_data.get('unique_dimms'))

            # Scrub for ticket comparison later
            clean_server_bios: str = scrub_server_bios(raw_server_bios)
            clean_server_bmc: str = scrub_server_bmc(raw_server_bmc)
            clean_server_cpld: list = scrub_server_cpld(raw_server_cpld)
            clean_server_os: str = scrub_server_os(raw_server_os)
            clean_server_tpm: str = scrub_server_tpm(raw_server_tpm)
            clean_server_ticket: str = scrub_server_ticket(raw_server_ticket)
            clean_machine_name: str = scrub_machine_name(raw_machine_name)
            clean_checked_out_to: str = scrub_checked_out_to(raw_checked_out_to)

            # Store Pipe Data
            all_system_data: dict = {}
            current_ticket_data: dict = ado_data.get(clean_server_ticket, 'None')
            ticket_state: str = ado_data.get('ticket_state', 'None')
            raw_ticket_title: str = ado_data.get(clean_server_ticket, {}).get('title')

            # Store Ticket
            try:
                ticket_table_data: dict = current_ticket_data.get('table_data', 'None')
                ticket_qcl_parts: dict = current_ticket_data.get('qcl_parts', 'None')
                all_system_data['ticket_number'] = clean_server_ticket

                # Clean Data for easier comparison later
                all_system_data['table_ssd_data']: dict = get_ticket_ssd_data(ticket_table_data)
                all_system_data['table_nvme_data']: dict = get_ticket_nvme_data(ticket_table_data)
                all_system_data['table_hdd_data']: dict = get_ticket_hdd_data(ticket_table_data)
                all_system_data['table_dimm_data']: dict = get_ticket_dimm_data(ticket_table_data)

                all_system_data['raw_unique_disks_count']: str = raw_unique_disks_count
                all_system_data['raw_unique_disks_model']: str = raw_unique_disks_model
                all_system_data['raw_unique_disks_firmware']: str = raw_unique_disks_firmware
                all_system_data['raw_unique_nvmes_count']: str = raw_unique_nvmes_count
                all_system_data['raw_unique_nvmes_model']: str = raw_unique_nvmes_model
                all_system_data['raw_unique_nvmes_firmware']: str = raw_unique_nvmes_firmware
                all_system_data['raw_unique_dimms_count']: str = raw_unique_dimms_count
                all_system_data['raw_unique_dimms_size']: str = raw_unique_dimms_size
                all_system_data['raw_unique_dimms_part']: str = raw_unique_dimms_part

                raw_ticket_bios: str = ticket_table_data.get('server_bios')
                raw_ticket_bmc: str = ticket_table_data.get('server_bmc')
                raw_ticket_cpld: str = ticket_table_data.get('server_cpld')
                raw_ticket_os: str = ticket_table_data.get('server_os').strip()
                raw_ticket_tpm: str = ticket_table_data.get('server_tpm')
                raw_ticket_boot_drive: str = ticket_table_data.get('server_boot_drive')
                raw_ticket_toolkit: str = ticket_table_data.get('toolkit')
                raw_ticket_reference_test_plans: str = check_reference_test_plans(ticket_table_data)
                raw_ticket_chipset_driver: str = ticket_table_data.get('server_chipset_driver')
                raw_ticket_server_processors: str = ticket_table_data.get('server_processors')
                raw_ticket_fpga_release: str = ticket_table_data.get('server_fpga_release_package')
                raw_ticket_nic_firmware: str = ticket_table_data.get('server_nic_firmware')
                raw_ticket_nic_pxe: str = ticket_table_data.get('server_nic_pxe')
                raw_ticket_nic_uefi: str = ticket_table_data.get('server_nic_uefi')
                raw_ticket_nic_driver: str = ticket_table_data.get('server_nic_driver')
                raw_ticket_rm_firmware: str = ticket_table_data.get('rack_manager_firmware')
                raw_ticket_request_type: str = ticket_table_data.get('request_type')
                raw_ticket_target_configuration: str = ticket_table_data.get('target_configuration')
                raw_ticket_part_number: str = ticket_table_data.get('part_number')
                raw_ticket_supplier: str = ticket_table_data.get('supplier')
                raw_ticket_description: str = ticket_table_data.get('description')
                raw_ticket_datasheet: str = ticket_table_data.get('datasheet')
                raw_ticket_diagnostic_utility: str = ticket_table_data.get('diagnostic_utility')
                raw_ticket_firmware_update_utility: str = ticket_table_data.get('firmware_update_utility')
                raw_ticket_firmware: str = ticket_table_data.get('firmware')
                raw_ticket_firmware_n_1: str = ticket_table_data.get('firmware_n-1')

                # Scrub for system comparison later
                clean_ticket_bios: str = scrub_ticket_bios(raw_ticket_bios)
                clean_ticket_bmc: str = scrub_ticket_bmc(raw_ticket_bmc)
                clean_ticket_cpld: list = scrub_ticket_cpld(raw_ticket_cpld)
                clean_ticket_os: str = scrub_ticket_os(raw_ticket_os)
                clean_ticket_tpm: str = scrub_ticket_tpm(raw_ticket_tpm)
                clean_ticket_boot_drive: str = scrub_ticket_boot_drive(raw_ticket_boot_drive)
                clean_ticket_toolkit: str = scrub_ticket_toolkit(raw_ticket_toolkit)
                clean_ticket_reference_test_plans: str = scrub_ticket_reference_test_plans(
                    raw_ticket_reference_test_plans)
                clean_ticket_chipset_driver: str = scrub_ticket_chipset_driver(raw_ticket_chipset_driver)
                clean_ticket_server_processors: str = scrub_ticket_server_processors(raw_ticket_server_processors)
                clean_ticket_fpga_release: str = scrub_ticket_fpga_release(raw_ticket_fpga_release)
                clean_ticket_nic_firmware: str = scrub_ticket_nic_firmware(raw_ticket_nic_firmware)
                clean_ticket_nic_pxe: str = scrub_ticket_nic_pxe(raw_ticket_nic_pxe)
                clean_ticket_nic_uefi: str = scrub_ticket_nic_uefi(raw_ticket_nic_uefi)
                clean_ticket_nic_driver: str = scrub_ticket_nic_driver(raw_ticket_nic_driver)
                clean_ticket_rm_firmware: str = scrub_ticket_rm_firmware(raw_ticket_rm_firmware)
                clean_ticket_request_type: str = scrub_ticket_rm_firmware(raw_ticket_request_type)
                clean_ticket_target_configuration: str = scrub_ticket_rm_firmware(raw_ticket_target_configuration)
                clean_ticket_part_number: str = scrub_ticket_rm_firmware(raw_ticket_part_number)
                clean_ticket_supplier: str = scrub_ticket_rm_firmware(raw_ticket_supplier)
                clean_ticket_description: str = scrub_ticket_rm_firmware(raw_ticket_description)
                clean_ticket_datasheet: str = scrub_ticket_rm_firmware(raw_ticket_datasheet)
                clean_ticket_diagnostic_utility: str = scrub_ticket_rm_firmware(raw_ticket_diagnostic_utility)
                clean_ticket_firmware_update_utility: str = scrub_ticket_rm_firmware(raw_ticket_firmware_update_utility)
                clean_ticket_firmware: str = scrub_ticket_rm_firmware(raw_ticket_firmware)
                clean_ticket_firmware_n_1: str = scrub_ticket_rm_firmware(raw_ticket_firmware_n_1)

                # Store original server data into all system data structure
                all_system_data['original_ticket_title'] = raw_ticket_title
                all_system_data['original_server_bios'] = raw_server_bios
                all_system_data['original_server_bmc'] = raw_server_bmc
                all_system_data['original_server_cpld'] = raw_server_cpld
                all_system_data['original_server_os'] = raw_server_os
                all_system_data['original_server_tpm'] = raw_server_tpm
                all_system_data['original_server_boot_drive'] = 'None'
                all_system_data['original_server_toolkit'] = 'None'
                all_system_data['original_server_reference_test_plans'] = 'None'
                all_system_data['original_server_chipset_driver'] = 'None'
                all_system_data['original_server_server_processors'] = 'None'
                all_system_data['original_server_fpga_release'] = 'None'
                all_system_data['original_server_nic_firmware'] = 'None'
                all_system_data['original_server_nic_pxe'] = 'None'
                all_system_data['original_server_nic_uefi'] = 'None'
                all_system_data['original_server_nic_driver'] = 'None'
                all_system_data['original_server_rm_firmware'] = 'None'
                all_system_data['original_server_request_type'] = 'None'
                all_system_data['original_server_target_configuration'] = 'None'
                all_system_data['original_server_part_number'] = 'None'
                all_system_data['original_server_supplier'] = 'None'
                all_system_data['original_server_description'] = 'None'
                all_system_data['original_server_datasheet'] = 'None'
                all_system_data['original_server_diagnostic_utility'] = 'None'
                all_system_data['original_server_firmware_update_utility'] = 'None'
                all_system_data['original_server_firmware'] = 'None'
                all_system_data['original_server_firmware_n_1'] = 'None'

                # Store original ticket data into all system data structure
                all_system_data['original_ticket_bios'] = raw_ticket_bios
                all_system_data['original_ticket_bmc'] = raw_ticket_bmc
                all_system_data['original_ticket_cpld'] = raw_ticket_cpld
                all_system_data['original_ticket_os'] = raw_ticket_os
                all_system_data['original_ticket_boot_drive'] = raw_ticket_boot_drive
                all_system_data['original_ticket_toolkit'] = raw_ticket_toolkit
                all_system_data['original_ticket_reference_test_plans'] = raw_ticket_reference_test_plans
                all_system_data['original_ticket_chipset_driver'] = raw_ticket_chipset_driver
                all_system_data['original_ticket_server_processors'] = raw_ticket_server_processors
                all_system_data['original_ticket_fpga_release'] = raw_ticket_fpga_release
                all_system_data['original_ticket_nic_firmware'] = raw_ticket_nic_firmware
                all_system_data['original_ticket_nic_pxe'] = raw_ticket_nic_pxe
                all_system_data['original_ticket_nic_uefi'] = raw_ticket_nic_uefi
                all_system_data['original_ticket_nic_driver'] = raw_ticket_nic_driver
                all_system_data['original_ticket_rm_firmware'] = raw_ticket_rm_firmware
                all_system_data['original_ticket_request_type'] = raw_ticket_request_type
                all_system_data['original_ticket_target_configuration'] = raw_ticket_target_configuration
                all_system_data['original_ticket_part_number'] = raw_ticket_part_number
                all_system_data['original_ticket_supplier'] = raw_ticket_supplier
                all_system_data['original_ticket_description'] = raw_ticket_description
                all_system_data['original_ticket_datasheet'] = raw_ticket_datasheet
                all_system_data['original_ticket_diagnostic_utility'] = raw_ticket_diagnostic_utility
                all_system_data['original_ticket_firmware_update_utility'] = raw_ticket_firmware_update_utility
                all_system_data['original_ticket_firmware'] = raw_ticket_firmware
                all_system_data['original_ticket_firmware_n_1'] = raw_ticket_firmware_n_1

                # Store parsed server data into all system data structure
                all_system_data['parsed_server_bios'] = clean_server_bios
                all_system_data['parsed_server_bmc'] = clean_server_bmc
                all_system_data['parsed_server_cpld'] = clean_server_cpld
                all_system_data['parsed_server_os'] = clean_server_os
                all_system_data['parsed_server_tpm'] = clean_server_tpm
                all_system_data['parsed_server_boot_drive'] = 'None'
                all_system_data['parsed_server_toolkit'] = 'None'
                all_system_data['parsed_server_reference_test_plans'] = 'None'
                all_system_data['parsed_server_chipset_driver'] = 'None'
                all_system_data['parsed_server_server_processors'] = 'None'
                all_system_data['parsed_server_fpga_release'] = 'None'
                all_system_data['parsed_server_nic_firmware'] = 'None'
                all_system_data['parsed_server_nic_pxe'] = 'None'
                all_system_data['parsed_server_nic_uefi'] = 'None'
                all_system_data['parsed_server_nic_driver'] = 'None'
                all_system_data['parsed_server_rm_firmware'] = 'None'
                all_system_data['parsed_server_request_type'] = 'None'
                all_system_data['parsed_server_target_configuration'] = 'None'
                all_system_data['parsed_server_part_number'] = 'None'
                all_system_data['parsed_server_supplier'] = 'None'
                all_system_data['parsed_server_description'] = 'None'
                all_system_data['parsed_server_datasheet'] = 'None'
                all_system_data['parsed_server_diagnostic_utility'] = 'None'
                all_system_data['parsed_server_firmware_update_utility'] = 'None'
                all_system_data['parsed_server_firmware'] = 'None'
                all_system_data['parsed_server_firmware_n_1'] = 'None'

                # Store parsed ticket data into all system data structure
                all_system_data['parsed_ticket_bios'] = clean_ticket_bios
                all_system_data['parsed_ticket_bmc'] = clean_ticket_bmc
                all_system_data['parsed_ticket_cpld'] = clean_ticket_cpld
                all_system_data['parsed_ticket_os'] = clean_ticket_os
                all_system_data['parsed_ticket_tpm'] = clean_ticket_tpm
                all_system_data['parsed_ticket_boot_drive'] = clean_ticket_boot_drive
                all_system_data['parsed_ticket_toolkit'] = clean_ticket_toolkit
                all_system_data['parsed_ticket_reference_test_plans'] = clean_ticket_reference_test_plans
                all_system_data['parsed_ticket_chipset_driver'] = clean_ticket_chipset_driver
                all_system_data['parsed_ticket_server_processors'] = clean_ticket_server_processors
                all_system_data['parsed_ticket_fpga_release'] = clean_ticket_fpga_release
                all_system_data['parsed_ticket_nic_firmware'] = clean_ticket_nic_firmware
                all_system_data['parsed_ticket_nic_pxe'] = clean_ticket_nic_pxe
                all_system_data['parsed_ticket_nic_uefi'] = clean_ticket_nic_uefi
                all_system_data['parsed_ticket_nic_driver'] = clean_ticket_nic_driver
                all_system_data['parsed_ticket_rm_firmware'] = clean_ticket_rm_firmware
                all_system_data['parsed_ticket_request_type'] = clean_ticket_request_type
                all_system_data['parsed_ticket_target_configuration'] = clean_ticket_target_configuration
                all_system_data['parsed_ticket_part_number'] = clean_ticket_part_number
                all_system_data['parsed_ticket_supplier'] = clean_ticket_supplier
                all_system_data['parsed_ticket_description'] = clean_ticket_description
                all_system_data['parsed_ticket_datasheet'] = clean_ticket_datasheet
                all_system_data['parsed_ticket_diagnostic_utility'] = clean_ticket_diagnostic_utility
                all_system_data['parsed_ticket_firmware'] = clean_ticket_firmware
                all_system_data['parsed_ticket_firmware_update_utility'] = clean_ticket_firmware_update_utility
                all_system_data['parsed_ticket_firmware_n_1'] = clean_ticket_firmware_n_1

                # Add State of the Ticket
                all_system_data['ticket_state'] = ticket_state
                all_system_data['ticket_title'] = raw_ticket_title

                # Compare system to ticket
                comparison_data: dict = component_comparison(pipe_name, clean_machine_name, clean_checked_out_to,
                                                             raw_server_ticket, all_system_data,
                                                             clean_ticket_request_type, ado_data, pipe_data)

                # Unpack comparing of system vs ticket data for tally count
                total_tally_count += comparison_data.get('total_count')
                match_tally_count += comparison_data.get('match_count')
                mismatch_tally_count += comparison_data.get('mismatch_count')
                missing_tally_count += comparison_data.get('missing_count')
                vse_tally_count += comparison_data.get('vse_count')
                other_tally_count += comparison_data.get('other_count')

                # Stores all system data into into all pipe data
                all_pipe_data[clean_machine_name] = all_system_data
            except AttributeError:
                pass

    # Store Tally Count
    all_pipe_data['total_tally'] = total_tally_count
    all_pipe_data['match_tally'] = match_tally_count
    all_pipe_data['mismatch_tally'] = mismatch_tally_count
    all_pipe_data['missing_tally'] = missing_tally_count
    all_pipe_data['vse_tally'] = vse_tally_count
    all_pipe_data['other_tally'] = other_tally_count

    # Store Ticket Review Count
    all_pipe_data['ticket_total_tally'] = ticket_total_tally
    all_pipe_data['ticket_match_tally'] = ticket_match_tally
    all_pipe_data['ticket_missing_tally'] = ticket_missing_tally

    return all_pipe_data


def is_version_in_qcl(trr_qualified_components: list, machine_versions: list) -> bool:
    if len(machine_versions) == 1:
        for version in machine_versions:
            for qualified_component in trr_qualified_components:
                if version in qualified_component:
                    return True
        else:
            return False

    elif len(machine_versions) >= 2:
        count: int = 0
        for version in machine_versions:
            for qualified_component in trr_qualified_components:
                if version in qualified_component:
                    count += 1

        if count >= len(machine_versions):
            return True
        else:
            return False


def is_commodity_version_in_trr(trr_qualified_commodity_list: list, machine_commodity_versions: list) -> bool:
    """

    """
    if len(machine_commodity_versions) == 1:
        comparison = list(set(trr_qualified_commodity_list).
                          intersection(machine_commodity_versions))

        if len(comparison) == 1:
            return True
        else:
            return False

    elif len(machine_commodity_versions) >= 2:
        count: int = 0
        for part_number in machine_commodity_versions:
            for qcl_part in trr_qualified_commodity_list:
                if part_number in qcl_part:
                    count += 1
        if count == len(machine_commodity_versions):
            return True
        else:
            return False


def get_dimm_part_numbers(pipe_data: dict, machine_name: str) -> list:
    """
    Returns list of unique dimm part numbers in the given machine name.
    """
    dimm_parts: list = []

    for component in pipe_data[machine_name]['unique_dimms']:
        dimm_parts.append(component['part'])

    return dimm_parts


def get_disk_part_numbers(pipe_data: dict, machine_name: str) -> list:
    disk_parts: list = []
    for component in pipe_data[machine_name]['unique_disks']:
        part_number: str = component['model']
        
        if not part_number or \
                part_number != 'Unknown' or \
                part_number == '':
            clean_part: str = get_disk_part_number(part_number)
            disk_parts.append(clean_part)
            
    return disk_parts


def get_nvme_part_numbers(pipe_data: dict, machine_name: str) -> list:
    versions: list = []
    for component in pipe_data[machine_name]['unique_nvmes']:
        part_number: str = component['model']

        if not part_number or \
                part_number != 'Unknown' or \
                part_number == '':
            clean_part: str = get_disk_part_number(part_number)
            versions.append(clean_part)

    return versions


def get_nvme_firmware(pipe_data: dict, machine_name: str) -> list:
    versions: list = []
    for component in pipe_data[machine_name]['unique_nvmes']:
        part_number: str = component['firmware']

        if not part_number or \
                part_number != 'Unknown' or \
                part_number == '':
            clean_part: str = get_disk_part_number(part_number)
            versions.append(clean_part)

    return versions


def get_empty_list(dimm_parts) -> list:
    for item in dimm_parts:
        item.strip()
        if not item.strip():
            dimm_parts.remove(item)
    return dimm_parts


def get_commodity_firmware(pipe_data: dict, machine_name: str) -> list:
    dimm_parts: list = []

    for component in pipe_data[machine_name]['unique_nvmes']:
        disk_firmware: str = component['firmware']

        if not disk_firmware:
            pass

        elif 'Unknown' not in disk_firmware:
            clean_part: str = get_disk_part_number(disk_firmware)
            dimm_parts.append(clean_part)

    return dimm_parts


def component_comparison(pipe_name: str, machine_name: str, checked_out_to: str, raw_server_ticket: str,
                         all_system_data: dict, clean_ticket_request_type: str, ado_data: dict, pipe_data) -> dict:
    """

    """
    total_count: int = 0
    match_count: int = 0
    mismatch_count: int = 0
    missing_count: int = 0
    vse_count: int = 0
    other_count: int = 0

    # SYSTEM - Software Stack
    original_ticket_title: str = all_system_data.get('original_ticket_title')
    original_system_bios: str = all_system_data.get('original_server_bios')
    original_system_bmc: str = all_system_data.get('original_server_bmc')
    original_system_cpld: str = all_system_data.get('original_server_cpld')
    original_system_os: str = all_system_data.get('original_server_os')
    original_system_boot_drive: str = all_system_data.get('None', 'None')
    original_system_toolkit: str = all_system_data.get('None', 'None')
    original_system_reference_test_plans: str = all_system_data.get('None', 'None')
    original_system_chipset_driver: str = all_system_data.get('None', 'None')
    original_system_server_processors: str = all_system_data.get('None', 'None')
    original_system_fpga_release: str = all_system_data.get('None', 'None')
    original_system_nic_firmware: str = all_system_data.get('None', 'None')
    original_system_nic_pxe: str = all_system_data.get('None', 'None')
    original_system_nic_uefi: str = all_system_data.get('None', 'None')
    original_system_nic_driver: str = all_system_data.get('None', 'None')
    original_system_rm_firmware: str = all_system_data.get('None', 'None')
    original_system_request_type: str = all_system_data.get('None', 'None')
    original_system_target_configuration: str = all_system_data.get('None', 'None')
    original_system_part_number: str = all_system_data.get('None', 'None')
    original_system_supplier: str = all_system_data.get('None', 'None')
    original_system_description: str = all_system_data.get('None', 'None')
    original_system_datasheet: str = all_system_data.get('None', 'None')
    original_system_diagnostic_utility: str = all_system_data.get('None', 'None')
    original_system_firmware_update_utility: str = all_system_data.get('None', 'None')
    original_system_firmware: str = all_system_data.get('None', 'None')
    original_system_firmware_n_1: str = all_system_data.get('None', 'None')
    original_system_title: str = all_system_data.get('None', 'None')

    # SYSTEM - Commodities
    # system_disk_part_numbers = all_system_data.get('raw_unique_disks_model', 'None')
    # system_disk_firmware_versions = all_system_data.get('raw_unique_disks_firmware', 'None')
    # system_nvme_part_numbers = all_system_data.get('raw_unique_nvmes_model', 'None')
    # system_nvme_firmware_versions = all_system_data.get('raw_unique_nvmes_firmware', 'None')
    # system_dimm_part_numbers = all_system_data.get('raw_unique_dimms_part', 'None')

    # Unpack data from ticket
    original_checked_out_to: str = all_system_data.get('original_checked_out_to')
    original_ticket_bios: str = all_system_data.get('original_ticket_bios')
    original_ticket_bmc: str = all_system_data.get('original_ticket_bmc')
    original_ticket_cpld: str = all_system_data.get('original_ticket_cpld')
    original_ticket_os: str = all_system_data.get('original_ticket_os')
    original_ticket_boot_drive: str = all_system_data.get('original_ticket_boot_drive')
    original_ticket_toolkit: str = all_system_data.get('original_ticket_toolkit')
    original_ticket_reference_test_plans: str = all_system_data.get('original_ticket_reference_test_plans')
    original_ticket_chipset_driver: str = all_system_data.get('original_ticket_chipset_driver')
    original_ticket_server_processors: str = all_system_data.get('original_ticket_server_processors')
    original_ticket_fpga_release: str = all_system_data.get('original_ticket_fpga_release')
    original_ticket_nic_firmware: str = all_system_data.get('original_ticket_nic_firmware')
    original_ticket_nic_pxe: str = all_system_data.get('original_ticket_nic_pxe')
    original_ticket_nic_uefi: str = all_system_data.get('original_ticket_nic_uefi')
    original_ticket_nic_driver: str = all_system_data.get('original_ticket_nic_driver')
    original_ticket_rm_firmware: str = all_system_data.get('original_ticket_rm_firmware')
    original_ticket_request_type: str = all_system_data.get('original_ticket_request_type')
    original_ticket_target_configuration: str = all_system_data.get('original_ticket_target_configuration')
    original_ticket_part_number: str = all_system_data.get('original_ticket_part_number')
    original_ticket_supplier: str = all_system_data.get('original_ticket_supplier')
    original_ticket_description: str = all_system_data.get('original_ticket_description')
    original_ticket_datasheet: str = all_system_data.get('original_ticket_datasheet')
    original_ticket_diagnostic_utility: str = all_system_data.get('original_ticket_diagnostic_utility')
    original_ticket_firmware_update_utility: str = all_system_data.get('original_ticket_firmware_update_utility')
    original_ticket_firmware: str = all_system_data.get('original_ticket_firmware')
    original_ticket_firmware_n_1: str = all_system_data.get('original_ticket_firmware_n_1')
    original_ticket_title: str = all_system_data.get('original_ticket_title')

    # original_nvme_data: dict = all_system_data.get('table_nvme_data')
    # original_dimm_data: dict = all_system_data.get('table_dimm_data')
    # original_ssd_data: dict = all_system_data.get('table_ssd_data')
    # original_hdd_data: dict = all_system_data.get('table_hdd_data')

    # TICKET - Unpack Part Numbers and Firmware Versions
    # ticket_nvme_firmware_versions: list = original_nvme_data.get('firmware_versions')
    # ticket_ssd_firmware_versions: list = original_ssd_data.get('firmware_versions')
    # ticket_hdd_firmware_versions: list = original_hdd_data.get('firmware_versions')
    # ticket_disk_combine_firmware_versions: list = ticket_ssd_firmware_versions + ticket_hdd_firmware_versions
    # ticket_nvme_part_numbers: list = original_nvme_data.get('part_numbers')
    # ticket_dimm_part_numbers: list = original_dimm_data.get('part_numbers')
    # ticket_ssd_part_numbers: list = original_ssd_data.get('part_numbers')
    # ticket_hdd_part_numbers: list = original_hdd_data.get('part_numbers')
    # ticket_disk_combine_part_numbers: list = ticket_ssd_part_numbers + ticket_hdd_part_numbers

    # SYSTEM - Unpack data
    parsed_system_bios: str = all_system_data.get('parsed_server_bios')
    parsed_system_bmc: str = all_system_data.get('parsed_server_bmc')
    parsed_system_cpld: list = all_system_data.get('parsed_server_cpld')
    parsed_system_os: str = all_system_data.get('parsed_server_os')
    parsed_system_boot_drive: str = all_system_data.get('parsed_server_boot_drive')
    parsed_system_toolkit: str = all_system_data.get('parsed_server_toolkit')
    parsed_system_reference_test_plans: str = all_system_data.get('parsed_server_reference_test_plans')
    parsed_system_chipset_driver: str = all_system_data.get('parsed_server_chipset_driver')
    parsed_system_server_processors: str = all_system_data.get('parsed_server_server_processors')
    parsed_system_fpga_release: str = all_system_data.get('parsed_server_fpga_release')
    parsed_system_nic_firmware: str = all_system_data.get('parsed_server_nic_firmware')
    parsed_system_nic_pxe: str = all_system_data.get('parsed_server_nic_pxe')
    parsed_system_nic_uefi: str = all_system_data.get('parsed_server_nic_uefi')
    parsed_system_nic_driver: str = all_system_data.get('parsed_server_nic_driver')
    parsed_system_rm_firmware: str = all_system_data.get('parsed_server_rm_firmware')
    parsed_system_request_type: str = all_system_data.get('parsed_server_request_type')
    parsed_system_target_configuration: str = all_system_data.get('parsed_server_target_configuration')
    parsed_system_part_number: str = all_system_data.get('parsed_server_part_number')
    parsed_system_supplier: str = all_system_data.get('parsed_server_supplier')
    parsed_system_description: str = all_system_data.get('parsed_server_description')
    parsed_system_datasheet: str = all_system_data.get('parsed_server_datasheet')
    parsed_system_diagnostic_utility: str = all_system_data.get('parsed_server_diagnostic_utility')
    parsed_system_firmware_update_utility: str = all_system_data.get('parsed_server_firmware_update_utility')
    parsed_system_firmware: str = all_system_data.get('parsed_server_firmware')
    parsed_system_firmware_n_1: str = all_system_data.get('parsed_server_firmware_n_1')
    parsed_system_title: str = all_system_data.get('parsed_server_title')

    # Unpack data from ticket
    parsed_ticket_bios: str = all_system_data.get('parsed_ticket_bios')
    parsed_ticket_bmc: str = all_system_data.get('parsed_ticket_bmc')
    parsed_ticket_cpld: str = all_system_data.get('parsed_ticket_cpld')
    parsed_ticket_os: str = all_system_data.get('parsed_ticket_os')
    parsed_ticket_boot_drive: str = all_system_data.get('parsed_ticket_boot_drive')
    parsed_ticket_toolkit: str = all_system_data.get('parsed_ticket_toolkit')
    parsed_ticket_reference_test_plans: str = all_system_data.get('parsed_ticket_reference_test_plans')
    parsed_ticket_chipset_driver: str = all_system_data.get('parsed_ticket_chipset_driver')
    parsed_ticket_server_processors: str = all_system_data.get('parsed_ticket_server_processors')
    parsed_ticket_fpga_release: str = all_system_data.get('parsed_ticket_fpga_release')
    parsed_ticket_nic_firmware: str = all_system_data.get('parsed_ticket_nic_firmware')
    parsed_ticket_nic_pxe: str = all_system_data.get('parsed_ticket_nic_pxe')
    parsed_ticket_nic_uefi: str = all_system_data.get('parsed_ticket_nic_uefi')
    parsed_ticket_nic_driver: str = all_system_data.get('parsed_ticket_nic_driver')
    parsed_ticket_rm_firmware: str = all_system_data.get('parsed_ticket_rm_firmware')
    parsed_ticket_request_type: str = all_system_data.get('parsed_ticket_request_type')
    parsed_ticket_target_configuration: str = all_system_data.get('parsed_ticket_target_configuration')
    parsed_ticket_part_number: str = all_system_data.get('parsed_ticket_part_number')
    parsed_ticket_supplier: str = all_system_data.get('parsed_ticket_supplier')
    parsed_ticket_description: str = all_system_data.get('parsed_ticket_description')
    parsed_ticket_datasheet: str = all_system_data.get('parsed_ticket_datasheet')
    parsed_ticket_diagnostic_utility: str = all_system_data.get('parsed_ticket_diagnostic_utility')
    parsed_ticket_firmware_update_utility: str = all_system_data.get('parsed_ticket_firmware_update_utility')
    parsed_ticket_firmware: str = all_system_data.get('parsed_ticket_firmware')
    parsed_ticket_firmware_n_1: str = all_system_data.get('parsed_ticket_firmware_n_1')
    parsed_ticket_title: str = all_system_data.get('parsed_ticket_title')

    # COMPARISON - System vs Ticket - Software Stack
    machine_info: dict = store_machine_basic_information(checked_out_to, machine_name, pipe_name, raw_server_ticket)

    compare_dimm_part_number: str = check_system_dimm_in_qcl(ado_data, pipe_data, machine_info)

    compare_disk_part_number: str = check_disk_part_number_in_qcl(ado_data, pipe_data, machine_info)

    compare_disk_firmware: str = check_disk_firmware_in_qcl(ado_data, pipe_data, machine_info)

    compare_nvme_part_number: str = check_nvme_part_number_in_qcl(ado_data, pipe_data, machine_info)

    compare_nvme_firmware: str = check_nvme_firmware_in_qcl(ado_data, pipe_data, machine_info)

    compare_missing_ticket: str = compare_system_and_ticket(checked_out_to, pipe_name, machine_name, raw_server_ticket,
                                                            'Missing Ticket',
                                                            'Missing TRR Field',
                                                            '',
                                                            '',
                                                            '',
                                                            clean_ticket_request_type)

    compare_bios: str = compare_system_and_ticket(checked_out_to, pipe_name, machine_name, raw_server_ticket,
                                                  'BIOS',
                                                  original_system_bios,
                                                  original_ticket_bios,
                                                  parsed_system_bios,
                                                  parsed_ticket_bios,
                                                  clean_ticket_request_type)

    compare_bmc: str = compare_system_and_ticket(checked_out_to, pipe_name, machine_name, raw_server_ticket,
                                                 'BMC',
                                                 original_system_bmc,
                                                 original_ticket_bmc,
                                                 parsed_system_bmc,
                                                 parsed_ticket_bmc,
                                                 clean_ticket_request_type)

    compare_cpld: str = compare_system_and_ticket(checked_out_to, pipe_name, machine_name, raw_server_ticket,
                                                  'CPLD',
                                                  original_system_cpld,
                                                  original_ticket_cpld,
                                                  parsed_system_cpld,
                                                  parsed_ticket_cpld,
                                                  clean_ticket_request_type)

    compare_os: str = compare_system_and_ticket(checked_out_to, pipe_name, machine_name, raw_server_ticket,
                                                'OS',
                                                original_system_os,
                                                original_ticket_os,
                                                parsed_system_os,
                                                parsed_ticket_os,
                                                clean_ticket_request_type)

    compare_boot_drive: str = compare_system_and_ticket(checked_out_to, pipe_name, machine_name, raw_server_ticket,
                                                        'Boot Drive',
                                                        original_system_boot_drive,
                                                        'Boot Drive - Missing in TRR',
                                                        parsed_system_boot_drive,
                                                        parsed_ticket_boot_drive,
                                                        clean_ticket_request_type)

    compare_toolkit: str = compare_system_and_ticket(checked_out_to, pipe_name, machine_name, raw_server_ticket,
                                                     'Toolkit',
                                                     original_system_toolkit,
                                                     'Toolkit - Missing in TRR',
                                                     parsed_system_toolkit,
                                                     parsed_ticket_toolkit,
                                                     clean_ticket_request_type)

    compare_chipset_driver: str = compare_system_and_ticket(checked_out_to, pipe_name, machine_name, raw_server_ticket,
                                                            'Chipset Driver',
                                                            original_system_chipset_driver,
                                                            'Chipset Driver - Missing in TRR',
                                                            parsed_system_chipset_driver,
                                                            parsed_ticket_chipset_driver,
                                                            clean_ticket_request_type)

    compare_server_processors: str = compare_system_and_ticket(checked_out_to, pipe_name, machine_name,
                                                               raw_server_ticket, 'Processors',
                                                               original_system_server_processors,
                                                               'Processors - Missing in TRR',
                                                               parsed_system_server_processors,
                                                               parsed_ticket_server_processors,
                                                               clean_ticket_request_type)

    compare_fpga_release: str = compare_system_and_ticket(checked_out_to, pipe_name, machine_name, raw_server_ticket,
                                                          'FPGA Release',
                                                          original_system_fpga_release,
                                                          'FPGA Release - Missing in TRR',
                                                          parsed_system_fpga_release,
                                                          parsed_ticket_fpga_release,
                                                          clean_ticket_request_type)

    compare_nic_firmware: str = compare_system_and_ticket(checked_out_to, pipe_name, machine_name, raw_server_ticket,
                                                          'NIC Firmware',
                                                          original_system_nic_firmware,
                                                          'NIC Firmware - Missing in TRR',
                                                          parsed_system_nic_firmware,
                                                          parsed_ticket_nic_firmware,
                                                          clean_ticket_request_type)

    compare_nic_pxe: str = compare_system_and_ticket(checked_out_to, pipe_name, machine_name, raw_server_ticket,
                                                     'NIC PXE',
                                                     original_system_nic_pxe,
                                                     'NIC PXE - Missing in TRR',
                                                     parsed_system_nic_pxe,
                                                     parsed_ticket_nic_pxe,
                                                     clean_ticket_request_type)

    compare_nic_uefi: str = compare_system_and_ticket(checked_out_to, pipe_name, machine_name, raw_server_ticket,
                                                      'NIC UEFI',
                                                      original_system_nic_uefi,
                                                      'NIC UEFI - Missing in TRR',
                                                      parsed_system_nic_uefi,
                                                      parsed_ticket_nic_uefi,
                                                      clean_ticket_request_type)

    compare_nic_driver: str = compare_system_and_ticket(checked_out_to, pipe_name, machine_name, raw_server_ticket,
                                                        'NIC Driver',
                                                        original_system_nic_driver,
                                                        'NIC Driver - Missing in TRR',
                                                        parsed_system_nic_driver,
                                                        parsed_ticket_nic_driver,
                                                        clean_ticket_request_type)

    compare_rm_firmware: str = compare_system_and_ticket(checked_out_to, pipe_name, machine_name, raw_server_ticket,
                                                         'RM Firmware',
                                                         original_system_rm_firmware,
                                                         'Rack Manager Firmware - Missing in TRR',
                                                         parsed_system_rm_firmware,
                                                         parsed_ticket_rm_firmware,
                                                         clean_ticket_request_type)

    compare_request_type: str = compare_system_and_ticket(checked_out_to, pipe_name, machine_name, raw_server_ticket,
                                                          'Request Type',
                                                          original_system_request_type,
                                                          'Request TYpe - Missing in TRR',
                                                          parsed_system_request_type,
                                                          parsed_ticket_request_type,
                                                          clean_ticket_request_type)

    compare_target_configuration: str = compare_system_and_ticket(checked_out_to, pipe_name, machine_name,
                                                                  raw_server_ticket,
                                                                  'Target Config',
                                                                  original_system_target_configuration,
                                                                  'Target Configuration - Missing in TRR',
                                                                  parsed_system_target_configuration,
                                                                  parsed_ticket_target_configuration,
                                                                  clean_ticket_request_type)

    compare_part_number: str = compare_system_and_ticket(checked_out_to, pipe_name, machine_name, raw_server_ticket,
                                                         'Part Number',
                                                         original_system_part_number,
                                                         'Part Number - Missing in TRR',
                                                         parsed_system_part_number,
                                                         parsed_ticket_part_number,
                                                         clean_ticket_request_type)

    compare_supplier: str = compare_system_and_ticket(checked_out_to, pipe_name, machine_name, raw_server_ticket,
                                                      'Supplier',
                                                      original_system_supplier,
                                                      'Supplier - Missing in TRR',
                                                      parsed_system_supplier,
                                                      parsed_ticket_supplier,
                                                      clean_ticket_request_type)

    compare_description: str = compare_system_and_ticket(checked_out_to, pipe_name, machine_name, raw_server_ticket,
                                                         'Description',
                                                         original_system_description,
                                                         'Description - Missing in TRR',
                                                         parsed_system_description,
                                                         parsed_ticket_description,
                                                         clean_ticket_request_type)

    compare_datasheet: str = compare_system_and_ticket(checked_out_to, pipe_name, machine_name, raw_server_ticket,
                                                       'Datasheet',
                                                       original_system_datasheet,
                                                       'Datasheet - Missing in TRR',
                                                       parsed_system_datasheet,
                                                       parsed_ticket_datasheet,
                                                       clean_ticket_request_type)

    compare_diagnostic_utility: str = compare_system_and_ticket(checked_out_to, pipe_name, machine_name,
                                                                raw_server_ticket,
                                                                'Dia. Utility',
                                                                original_system_diagnostic_utility,
                                                                'Diagnostic Utility - Missing in TRR',
                                                                parsed_system_diagnostic_utility,
                                                                parsed_ticket_diagnostic_utility,
                                                                clean_ticket_request_type)

    compare_firmware_update_utility: str = compare_system_and_ticket(checked_out_to, pipe_name, machine_name,
                                                                     raw_server_ticket, 'FM Update Util.',
                                                                     original_system_firmware_update_utility,
                                                                     'Firmware Update Utility - Missing in TRR',
                                                                     parsed_system_firmware_update_utility,
                                                                     parsed_ticket_firmware_update_utility,
                                                                     clean_ticket_request_type)

    compare_firmware: str = compare_system_and_ticket(checked_out_to, pipe_name, machine_name, raw_server_ticket,
                                                      'Firmware',
                                                      original_system_firmware,
                                                      'Firmware - Missing in TRR',
                                                      parsed_system_firmware,
                                                      parsed_ticket_firmware,
                                                      clean_ticket_request_type)

    compare_firmware_n_1: str = compare_system_and_ticket(checked_out_to, pipe_name, machine_name, raw_server_ticket,
                                                          'Firmware N-1',
                                                          original_system_firmware_n_1,
                                                          'Firmware N-1 - Missing in TRR',
                                                          parsed_system_firmware_n_1,
                                                          parsed_ticket_firmware_n_1,
                                                          clean_ticket_request_type)

    # Store compare results in list to make it easier to iterate through later
    all_comparison: list = [compare_missing_ticket, compare_bios, compare_bmc, compare_cpld, compare_os,
                            compare_boot_drive, compare_toolkit, compare_chipset_driver, compare_server_processors,
                            compare_fpga_release, compare_nic_firmware, compare_nic_pxe, compare_nic_uefi,
                            compare_nic_driver, compare_rm_firmware, compare_request_type, compare_target_configuration,
                            compare_part_number, compare_supplier, compare_description, compare_datasheet,
                            compare_diagnostic_utility, compare_firmware_update_utility, compare_firmware,
                            compare_firmware_n_1, compare_dimm_part_number, compare_disk_part_number,
                            compare_disk_firmware, compare_nvme_part_number, compare_nvme_firmware]

    # Tally Result
    for comparison_result in all_comparison:
        total_count += 1
        if comparison_result == 'MATCH':
            match_count += 1

        elif comparison_result == 'MISMATCH':
            mismatch_count += 1

        elif comparison_result == 'MISSING':
            missing_count += 1

        elif comparison_result == 'VSE':
            vse_count += 1

        else:
            other_count += 1

    # Store Results in structure
    pipe_number_summary: dict = {'total_count': total_count,
                                 'match_count': match_count,
                                 'mismatch_count': mismatch_count,
                                 'missing_count': missing_count,
                                 'vse_count': other_count,
                                 'other_count': other_count}

    return pipe_number_summary


def store_machine_basic_information(checked_out_to: str, machine_name: str, pipe_name: str, raw_server_ticket: str):
    return {'checked_out_to': checked_out_to,
            'machine_name': machine_name,
            'pipe_name': pipe_name,
            'ticket': raw_server_ticket}


def check_system_dimm_in_qcl(azure_devops_data: dict, pipe_data: dict, machine_info: dict) -> str:
    """
    Checks machine data to QCL within TRR
    :return: MATCH or MISMATCH
    """
    add_to_total_checks()

    machine_name: str = machine_info['machine_name']
    ticket: str = machine_info['ticket']

    machine_dimm_part_numbers: list = get_dimm_part_numbers(pipe_data, machine_name)
    trr_qualified_components: list = get_trr_qualified_components(azure_devops_data, ticket)

    is_dimm_in_qcl_parts: bool = is_version_in_qcl(trr_qualified_components, machine_dimm_part_numbers)

    if not is_dimm_in_qcl_parts:
        version_statement: str = 'DIMM P/N - Not in TRR QCL'
        difference_statement: str = get_difference_statement(trr_qualified_components, machine_dimm_part_numbers)

        return add_to_all_issues(version_statement, difference_statement, machine_info)

    else:
        return 'MATCH'


def add_to_all_issues(version_statement: str, difference_statement: str, machine_info: dict) -> str:

    checked_out_to: str = machine_info['checked_out_to']
    pipe_name: str = machine_info['pipe_name']
    machine_name: str = machine_info['machine_name']
    ticket: str = machine_info['ticket']

    difference_statement: str = check_difference_statement(difference_statement, version_statement)
    system_component: str = get_system_component(version_statement)

    all_issues.append(issue_report(checked_out_to,
                                   pipe_name,
                                   machine_name,
                                   ticket,
                                   system_component,
                                   'MISMATCH_qcl',
                                   difference_statement,
                                   version_statement,
                                   difference_statement,
                                   version_statement,
                                   'Mismatch',
                                   'Machine'))
    return 'MISMATCH'


def get_system_component(version_statement: str):
    if 'DIMM P/N' in version_statement:
        return 'DIMM P/N'

    elif 'Disk P/N' in version_statement:
        return 'Disk P/N'

    elif 'Disk F/W' in version_statement:
        return 'Disk F/W'

    elif 'NVMe P/N' in version_statement:
        return 'NVMe P/N'

    elif 'NVMe F/W' in version_statement:
        return 'NVMe F/W'


def check_difference_statement(difference_statement: str, version_statement: str) -> str:
    if 'Empty' in difference_statement:
        if 'DIMM P/N' in version_statement:
            return 'No DIMM Part Number'

        elif 'Disk P/N' in version_statement:
            return 'No Disk Part Number'

        elif 'Disk F/W' in version_statement:
            return 'No Disk Firmware'

        elif 'NVMe P/N' in version_statement:
            return 'No NVMe Part Number'

        elif 'NVMe F/W' in version_statement:
            return 'No NVMe Firmware'

    else:
        return f'Mismatch - {difference_statement}'


def get_trr_qualified_components(azure_devops_data: dict, ticket: str) -> list:
    trr_qualified_components: list = azure_devops_data[ticket]['qcl_parts']

    flat_list: list = []
    for item in trr_qualified_components:
        if isinstance(item, list):
            for sub_item in item:
                flat_list.append(sub_item)
        else:
            flat_list.append(item)

    return list(set(flat_list))


def get_matched_components(ado_qcl_parts, system_dimm_parts):
    matched_component: list = []
    for component in system_dimm_parts:
        for qcl_part in ado_qcl_parts:
            if component in qcl_part:
                if not component:
                    pass
                else:
                    matched_component.append(component)

    return matched_component


def get_disk_part_number(unique_disk) -> str:
    split_parts: list = unique_disk.replace('  ', ' ').replace('_', ' ').split(' ')

    if len(split_parts) >= 3:
        for item in split_parts:
            if 'Micron' in item:
                return split_parts[-1].strip()
        else:
            return unique_disk.strip()

    count: int = 0
    for part_split in split_parts:
        if part_split.isalpha():
            count += 1

    if count == len(split_parts):
        return unique_disk.strip()

    else:
        for part_split in split_parts:
            if not part_split.isalpha():
                if not part_split:
                    return part_split
                else:
                    return part_split
        else:
            return unique_disk


def check_disk_part_number_in_qcl(azure_devops_data: dict, pipe_data: dict, machine_info: dict) -> str:
    """
    Checks machine data to QCL within TRR
    :return: MATCH or MISMATCH
    """
    add_to_total_checks()

    machine_name: str = machine_info['machine_name']
    ticket: str = machine_info['ticket']

    system_disk_parts: list = get_disk_part_numbers(pipe_data, machine_name)
    trr_qualified_components: list = get_trr_qualified_components(azure_devops_data, ticket)

    is_disk_in_qcl_parts: bool = is_version_in_qcl(trr_qualified_components, system_disk_parts)

    if not is_disk_in_qcl_parts:
        difference_statement: str = get_difference_statement(trr_qualified_components, system_disk_parts)
        version_statement: str = 'Disk P/N - Not in TRR QCL'

        return add_to_all_issues(version_statement, difference_statement, machine_info)

    else:
        return 'MATCH'


def check_nvme_part_number_in_qcl(azure_devops_data: dict, pipe_data: dict, machine_info: dict) -> str:
    """
    Ensures current system's NVMe part number is in assigned TRR from Console Server's ticket field.
    Assigned TRR (ex. 421697) should have part number. If not, will be added to issues as potential problems.
    
    :return: MATCH or MISMATCH - Added for overall tally later based on output
    """
    add_to_total_checks()

    machine_name: str = machine_info['machine_name']
    ticket: str = machine_info['ticket']

    commodity_versions: list = get_nvme_part_numbers(pipe_data, machine_name)
    trr_qualified_components: list = get_trr_qualified_components(azure_devops_data, ticket)

    machine_versions: list = get_system_nvme_firmwares(commodity_versions)

    is_disk_in_qcl_parts: bool = is_version_in_qcl(trr_qualified_components, machine_versions)

    if not is_disk_in_qcl_parts:
        version_statement: str = 'NVMe P/N - Not in TRR QCL'
        difference_statement: str = get_difference_statement(trr_qualified_components, commodity_versions)

        return add_to_all_issues(version_statement, difference_statement, machine_info)

    else:
        return 'MATCH'


def check_disk_firmware_in_qcl(azure_devops_data: dict, pipe_data: dict, machine_info: dict) -> str:
    """
    Ensures current system's disk (SSD or HDD) firmware is in assigned TRR from Console Server's ticket field.
    Assigned TRR (ex. 421697) should have firmware. If not, will be added to issues as potential problems.
    
    :return: MATCH or MISMATCH - Added for overall tally later based on output
    """
    add_to_total_checks()

    machine_name: str = machine_info['machine_name']
    ticket: str = machine_info['ticket']

    commodity_firmware: list = get_commodity_firmware(pipe_data, machine_name)
    trr_qualified_components: list = get_trr_qualified_components(azure_devops_data, ticket)

    if not is_version_in_qcl(trr_qualified_components, commodity_firmware):
        version_statement: str = 'Disk F/W - Not in TRR QCL'
        difference_statement: str = get_difference_statement(trr_qualified_components, commodity_firmware)

        return add_to_all_issues(version_statement, difference_statement, machine_info)

    else:
        return 'MATCH'


def get_difference_statement(trr_qualified_components: list, commodity_versions: list) -> str:
    matched_component: list = get_matched_components(trr_qualified_components, commodity_versions)
    difference = list(set(commodity_versions) - set(matched_component))

    if not difference:
        return 'Empty'
    else:
        return ', '.join(difference)


def add_to_total_checks() -> None:
    """
    Adds to total checks being done to ensure Console Server and TRR compares correctly.
    """
    total_checks.append(1)


def check_nvme_firmware_in_qcl(azure_devops_data: dict, pipe_data: dict, machine_info: dict):
    add_to_total_checks()

    ticket: str = machine_info['ticket']
    machine_name: str = machine_info['machine_name']

    commodity_versions: list = get_nvme_firmware(pipe_data, machine_name)
    trr_qualified_components: list = get_trr_qualified_components(azure_devops_data, ticket)

    if not is_version_in_qcl(trr_qualified_components, commodity_versions):
        version_statement: str = 'NVMe F/W - Not in TRR QCL'
        difference_statement: str = get_difference_statement(trr_qualified_components, commodity_versions)

        add_to_all_issues(version_statement, difference_statement, machine_info)

    else:
        return 'MATCH'


def get_system_nvme_firmwares(system_nvme_firmwares):
    clean_items: list = []
    for item in system_nvme_firmwares:
        if not item:
            pass
        else:
            clean_items.append(item)
    return clean_items


def compare_system_and_ticket(checked_out_to: str, pipe_name: str, clean_machine_name: str, server_ticket: str,
                              component: str, original_system, original_ticket, parsed_system, parsed_ticket,
                              ticket_request_type: str) -> str:
    """
    Compares string value from ADO to Console Server per component given.
    :param checked_out_to: checked out to in Console Server
    :param ticket_request_type:
    :param parsed_ticket:
    :param parsed_system:
    :param original_ticket:
    :param original_system:
    :param pipe_name:
    :param server_ticket:
    :param clean_machine_name:
    :param component:
    :return:
    """
    total_checks.append(1)

    clean_system_value = parsed_system
    clean_ticket_value = parsed_ticket
    generation = clean_machine_name[5]
    clean_request_type = ticket_request_type.upper().replace(' TEST', '').replace('TEST', '')

    if component.upper() == 'CHECKED OUT TO' and check_missing(checked_out_to) == 'None':
        vse_issue_tally.append(1)
        new_issue: dict = issue_report(checked_out_to,
                                       pipe_name,
                                       clean_machine_name,
                                       server_ticket,
                                       component,
                                       'VSE',
                                       original_system,
                                       original_ticket,
                                       clean_system_value,
                                       clean_ticket_value,
                                       'Missing',
                                       'Console Server')
        all_issues.append(new_issue)
        return 'VSE'

    elif component.upper() == 'MISSING TICKET':
        vse_issue_tally.append(1)
        # new_issue: dict = issue_report(checked_out_to, pipe_name, clean_machine_name, server_ticket, component,
        #                                'VSE', original_system, original_ticket, clean_system_value,
        #                                clean_ticket_value, 'Missing', 'Console Server')
        # all_issues.append(new_issue)
        return 'VSE'

    elif component.upper() == 'CPLD':
        try:
            same_characters: list = []
            for system_character in clean_system_value:
                for ticket_character in clean_ticket_value:
                    if system_character == ticket_character:
                        same_characters.append(system_character)
            unique_characters = list(set(same_characters))

            if len(unique_characters) == 2 or clean_system_value == clean_ticket_value:
                return 'MATCH'

            else:
                # For All Issues in Main Dashboard later
                new_issue: dict = issue_report(checked_out_to, pipe_name, clean_machine_name, server_ticket, component,
                                               'MISMATCH', original_system, original_ticket, clean_system_value,
                                               clean_ticket_value, 'Mismatch', 'Console Server')
                all_issues.append(new_issue)
                mismatch_tally.append(1)
                return 'MISMATCH'

        except TypeError:
            return 'OTHER'
        except IndexError:
            return 'OTHER'

    elif component.upper() == 'BMC':

        if parsed_system in parsed_ticket or original_system in original_ticket:
            return 'MATCH'

        else:
            # For All Issues in Main Dashboard later
            new_issue: dict = issue_report(checked_out_to, pipe_name, clean_machine_name, server_ticket, component,
                                           'MISMATCH', original_system, original_ticket, clean_system_value,
                                           clean_ticket_value, 'Mismatch', 'Console Server')
            all_issues.append(new_issue)
            mismatch_tally.append(1)
            return 'MISMATCH'

    elif 'Title (' in component:
        if 'Request Type' in component:
            request_type: str = get_component_from_request_type(original_ticket)
            test_component: str = get_test_component_from_title(original_system)
            check_result: bool = check_request_type_in_title(original_system, original_ticket)

            if check_result is False:
                new_issue: dict = issue_report(checked_out_to, pipe_name, clean_machine_name, server_ticket, component,
                                               'MISMATCH', '',
                                               f'[{request_type}] not in TRR Title, {test_component} Instead',
                                               '', '', 'Mismatch', 'TRR')
                all_issues.append(new_issue)
                mismatch_tally.append(1)
                return 'MISMATCH'
        # TODO
        # print(component.upper())
        # print(f'ticket_title: {original_system}')

    elif generation == '5' and component == 'RM Firmware':
        return 'MATCH'

    elif clean_request_type == 'DIMM' and component == 'Dia. Utility':
        return 'MATCH'

    elif clean_request_type == 'DIMM' and component == 'FM Update Util.':
        return 'MATCH'

    elif clean_request_type == 'DIMM' and component == 'Firmware':
        return 'MATCH'

    elif clean_request_type == 'DIMM' and component == 'Firmware N-1':
        return 'MATCH'

    # Mostly for PM Review, looks for missing information on the table data
    elif component == 'Boot Drive' or component == 'Toolkit' or component == 'Test Plans' or \
            component == 'Chipset Driver' or component == 'Processors' or component == 'FPGA Release' or \
            component == 'NIC Firmware' or component == 'NIC PXE' or component == 'NIC UEFI' or \
            component == 'NIC Driver' or component == 'PSU Firmware' or component == 'RM Firmware' or \
            component == 'Request Type' or component == 'Target Config' or component == 'Part Number' or \
            component == 'Supplier' or component == 'Description' or component == 'Datasheet' or \
            component == 'Dia. Utility' or component == 'FM Update Util.' or component == 'Firmware' or \
            component == 'Firmware N-1':

        if clean_ticket_value is None or clean_ticket_value == '' or clean_ticket_value == 'None':
            new_issue: dict = issue_report(checked_out_to, pipe_name, clean_machine_name, server_ticket, component,
                                           'MISSING', original_system, original_ticket, clean_system_value,
                                           clean_ticket_value, 'Missing', 'TRR')
            all_issues.append(new_issue)
            missing_tally.append(1)
            return 'MISSING'

    else:
        if clean_ticket_value is None or clean_ticket_value == '' or clean_system_value is None or \
                clean_system_value == '' or clean_system_value == 'None' or clean_ticket_value == 'None':

            # For All Issues in Main Dashboard later
            new_issue: dict = issue_report(checked_out_to, pipe_name, clean_machine_name, server_ticket, component,
                                           'MISSING', original_system, original_ticket, clean_system_value,
                                           clean_ticket_value, 'Missing', 'TRR')
            all_issues.append(new_issue)
            missing_tally.append(1)
            return 'MISSING'

        elif parsed_system != parsed_ticket:

            # For All Issues in Main Dashboard later
            new_issue: dict = issue_report(checked_out_to, pipe_name, clean_machine_name, server_ticket, component,
                                           'MISMATCH', original_system, original_ticket, clean_system_value,
                                           clean_ticket_value, 'Mismatch', 'Comparison')
            all_issues.append(new_issue)
            mismatch_tally.append(1)
            return 'MISMATCH'
        elif clean_system_value == clean_ticket_value:
            return 'MATCH'

        # For any anomalies
        elif 'ERRONEOUS' in clean_system_value or 'ERRONEOUS' in clean_ticket_value:
            return 'OTHER'
        else:
            return 'OTHER'


def issue_report(username: str, pipe_name: str, machine_name: str, ticket_id: str, system_component: str,
                 issue_state: str,
                 original_system_data: str, original_ticket_data: str,
                 parsed_system_data: str, parsed_ticket_data: str, reason: str, section: str) -> dict:
    """
    Returns in callable dictionary format
    :param section: area of concern ex. Console Server, CRD, TRR, Z: Drive, etc.
    :param reason:
    :param username: checked out to in Console Server
    :param pipe_name: Host Group Name
    :param machine_name: System Name
    :param ticket_id: TRR
    :param system_component: Commodity
    :param issue_state: Missing or Mismatch
    :param original_system_data: Raw Data from System
    :param original_ticket_data: Raw Data from Ticket
    :param parsed_system_data: Cleaned Data from System
    :param parsed_ticket_data: Cleaned Data from System
    :return:
    """
    return {'username': username.strip(),
            'pipe_name': pipe_name.strip(),
            'machine_name': machine_name.strip(),
            'ticket_id': ticket_id.strip(),
            'issue_state': issue_state.strip(),
            'system_component': system_component.strip(),
            'original_system_data': original_system_data.strip(),
            'original_ticket_data': original_ticket_data.strip(),
            'parsed_system_data': parsed_system_data.strip(),
            'parsed_ticket_data': parsed_ticket_data.strip(),
            'reason': reason.strip(),
            'section': section.strip()}


def get_component_from_request_type(request_type: str) -> str:
    """
    Starts process for checking if Request Title is in the TRR title in ADO
    :param request_type: component being tested
    :return:
    """
    if 'SSD' in request_type.upper():
        return 'SSD'
    elif 'HDD' in request_type.upper():
        return 'HDD'
    elif 'NVME' in request_type.upper():
        return 'NVME'
    elif 'DIMM' in request_type.upper():
        return 'DIMM'
    else:
        return 'NONE'


def get_test_component_from_title(request_type: str) -> str:
    """
    Starts process for checking if Request Title is in the TRR title in ADO
    :param request_type: component being tested
    :return:
    """
    if 'SSD' in request_type.upper():
        return '[SSD]'
    elif 'HDD' in request_type.upper():
        return '[HDD]'
    elif 'NVME' in request_type.upper():
        return '[NVME]'
    elif 'DIMM' in request_type.upper():
        return '[DIMM]'
    else:
        return 'NONE'


def check_request_type_in_title(ticket_title: str, request_type: str) -> bool:
    """
    Starts process for checking if Request Title is in the TRR title in ADO
    :param ticket_title: TRR title containing the QCL workflows
    :param request_type: component being tested
    :return:
    """
    if 'SSD' in request_type.upper() and '[SSD]' in ticket_title.upper():
        return True
    elif 'HDD' in request_type.upper() and '[HDD]' in ticket_title.upper():
        return True
    elif 'NVME' in request_type.upper() and '[NVME]' in ticket_title.upper():
        return True
    elif 'DIMM' in request_type.upper() and '[DIMM]' in ticket_title.upper():
        return True
    else:
        return False


def get_all_issues():
    """
    Get All issues compiled from going through all information
    :return:
    """
    return all_issues


def get_total_checks():
    """
    Get All processed compiled from going through all information
    :return:
    """
    return str(sum(total_checks))


def get_missing_tally() -> str:
    """
    Get All processed compiled from going through all information
    :return:
    """
    return str(sum(missing_tally))


def get_mismatch_tally() -> str:
    """
    Get All processed compiled from going through all information
    :return:
    """
    return str(sum(mismatch_tally))


def get_vse_issues_count() -> str:
    """
    Get All processed compiled from going through all information
    :return:
    """
    return str(sum(vse_issue_tally))


def main_method(console_server_data: dict, ado_data: dict):
    for pipe_name in console_server_data:
        if 'Pipe-' in pipe_name and 'OFFLINE' not in pipe_name:

            current_pipe: dict = console_server_data.get(pipe_name)
            pipe_data: dict = current_pipe.get('pipe_data')
            pipe_unique_tickets: list = current_pipe.get('group_unique_tickets')

            # try-except accounts for Pipes with no data
            try:
                if len(pipe_unique_tickets) == 0 or pipe_unique_tickets == 'None':
                    pass
                else:
                    compare_data: dict = clean_pipe_data(pipe_name, pipe_data, ado_data)
                    console_server_data[pipe_name]['compare_data']: dict = compare_data
            except TypeError:
                pass
    console_server_data['host_groups_data']['vse_log'] = get_vse_issues_count()

    return console_server_data
