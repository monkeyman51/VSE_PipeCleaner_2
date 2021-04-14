from pipe_cleaner.src.credentials import Path
# from main import get_input
from time import strftime
from json import loads
import sys
import os


rights = []
wrongs = []
empty = []
unavailable = []
total = []
right_components = {}
wrong_components = {}
empty_component = {}


def get_json(pipe_num, trr_id, host_id):
    with open(f'{Path.info}{str(trr_id)}/final.json') as f:
        trr = loads(f.read())
    with open(f'{Path.info}{str(host_id)}.json') as f:
        system = loads(f.read())
    organize_report(pipe_num, trr_id, trr, system)


def organize_report(pipe_num, trr_id, trr, system):
    intro(pipe_num, trr_id, trr, system)
    bios_version(trr, system, pipe_num)
    bmc_version(trr, system, pipe_num)
    cpld_version(trr, system, pipe_num)
    tpm_version(trr, system, pipe_num)
    nic_firmware(trr, system, pipe_num)
    hdd_version(trr, system, pipe_num)
    hdd_firmware(trr, system, pipe_num)
    ssd_version(trr, system, pipe_num)
    ssd_firmware(trr, system, pipe_num)
    dimm_version(trr, system, pipe_num)
    nvme_version(trr, system, pipe_num)
    nvme_firmware(trr, system, pipe_num)
    os_version(trr, system, pipe_num)


    intro_two(pipe_num, trr_id, trr, system)
    hardware_configuration(trr, system, pipe_num)
    results(system, rights, wrongs, total, pipe_num)
    os_version(trr, system, pipe_num)
    bios_version(trr, system, pipe_num)
    bmc_version(trr, system, pipe_num)
    cpld_version(trr, system, pipe_num)
    tpm_version(trr, system, pipe_num)
    nic_firmware(trr, system, pipe_num)
    hdd_version(trr, system, pipe_num)
    hdd_firmware(trr, system, pipe_num)
    ssd_version(trr, system, pipe_num)
    ssd_firmware(trr, system, pipe_num)
    dimm_version(trr, system, pipe_num)
    nvme_version(trr, system, pipe_num)
    nvme_firmware(trr, system, pipe_num)
    # check_sku(trr)

    # end(system)


def replace_company_info(string):
    string_upper = str(string).upper()
    if 'SAMSUNG' in string_upper:
        string_upper = string_upper.replace('SAMSUNG', '')
    if 'SEAGATE' in str(string_upper):
        string_upper = string_upper.replace('SEAGATE', '')
    if 'SKHYNIX' in str(string_upper):
        string_upper = string_upper.replace('SKHYNIX', '')
    if 'INTEL' in str(string_upper):
        string_upper = string_upper.replace('INTEL', '')
    if 'P4511' in str(string_upper):
        string_upper = string_upper.replace('P4511', '')
    if 'WD' in str(string_upper):
        string_upper = string_upper.replace('WD', '')
    if 'MICRON' in str(string_upper):
        string_upper = string_upper.replace('MICRON', '')
    if '12TB' in str(string_upper):
        string_upper = string_upper.replace('12TB', '')
    if 'PE6011' in str(string_upper):
        string_upper = string_upper.replace('PE6011', '')
    if 'PM983' in str(string_upper):
        string_upper = string_upper.replace('PM983', '')
    if 'PM883' in str(string_upper):
        string_upper = string_upper.replace('PM883', '')
    if 'HYNIX' in str(string_upper):
        string_upper = string_upper.replace('HYNIX', '')
    if 'SE4011' in str(string_upper):
        string_upper = string_upper.replace('SE4011', '')
    if 'PE4010' in str(string_upper):
        string_upper = string_upper.replace('PE4010', '')
    if '32GB' in str(string_upper):
        string_upper = string_upper.replace('32GB', '')
    if 'PM983' in str(string_upper):
        string_upper = string_upper.replace('PM983', '')
    if 'LITEON' in str(string_upper):
        string_upper = string_upper.replace('LITEON', '')
    if 'PM963' in str(string_upper):
        string_upper = string_upper.replace('PM963', '')
    if 'TOSHIBA' in str(string_upper):
        string_upper = string_upper.replace('TOSHIBA', '')
    if '4TB' in str(string_upper):
        string_upper = string_upper.replace('4TB', '')
    if '8TB' in str(string_upper):
        string_upper = string_upper.replace('8TB', '')
    if '16GB' in str(string_upper):
        string_upper = string_upper.replace('32GB', '')
    if '960GB' in str(string_upper):
        string_upper = string_upper.replace('960GB', '')
    if '_5200_' in str(string_upper):
        string_upper = string_upper.replace('_5200_', '')
    return string_upper


def get_model(string):
    string_upper = str(string)
    string_upper = string_upper.upper()
    if '(' in string:
        string = str(string).split('(')[0]
        return string
    if '/' in string:
        string = str(string).split('/')[0]
        return string
    if 'FW:' in string_upper:
        string_upper = str(string_upper).split('FW:')[0]
    return string_upper


def get_firmware(string):
    string_upper = str(string)
    string_upper = string_upper.upper()
    if 'FW:' in string_upper:
        string_upper = str(string_upper).split('FW:')[-1]
        return string_upper
    elif '/' in string_upper:
        string_upper = str(string_upper).split('/')[-1]
        return string_upper


def replace_unnecesarry(string):
    unnecessary = str(string)
    if ' ' in unnecessary:
        unnecessary = unnecessary.replace(' ', '')
    if 'FW:' in unnecessary:
        unnecessary = unnecessary.replace('FW:', '')
    if ')' in unnecessary:
        unnecessary = unnecessary.replace(')', '')
    if '(' in unnecessary:
        unnecessary = unnecessary.replace('(', '')
    if '#' in unnecessary:
        unnecessary = unnecessary.replace('#', '')
    if '960GB' in unnecessary:
        unnecessary = unnecessary.replace('960GB', '')
    return unnecessary


def model(string):
    replace = replace_company_info(string)
    get = get_model(replace)
    unnecesarry = replace_unnecesarry(get)
    return unnecesarry


def firmware(string):
    replace = replace_company_info(string)
    get = get_firmware(replace)
    unnecesarry = replace_unnecesarry(get)
    return unnecesarry


def check_qcl(trr, component, parsed_system):
    component_string = str(component)
    component_upper = component_string.upper()
    part_number = get_part_number(trr)
    trr_firmware = get_trr_firmware(trr)
    for item in trr:
        item_string = str(item)
        item_upper = item_string.upper()
        if 'QCL' in item_upper and component_upper in item_upper:
            if parsed_system in model(trr[item]):
                return ('   Match')
            elif parsed_system in firmware(trr[item]):
                return ('   Match')
            elif parsed_system.split('-')[-1] in model(trr[item]):
                return ('   Match')
            elif parsed_system != model(trr[item]) and parsed_system != model(trr[item]):
                return ('>> MISMATCHED')
    return ('>> MISMATCHED')


def qcl_versions(trr, component, name, parsed_system, state, pipe_num):
    component_string = str(component)
    component_upper = component_string.upper()
    request = get_request_type(trr)
    part_number = get_part_number(trr)

    short_name = name[-3::]

    with open(f'{Path.reports}/pipe_{pipe_num}/{name}/text_{short_name}.txt', 'a') as f:
        f.write(f'  {state} - {component} Versions: \n')
        f.write(f'     - TRR QCL: \n')
        for item in trr:
            item_string = str(item)
            item_upper = item_string.upper()
            if 'QCL' in item_upper and component_upper in item_upper:
                if parsed_system in model(trr[item]):
                    f.write(f'       -- {model(trr[item])} << MATCHED \n')
                else:
                    f.write(f'       -- {model(trr[item])} \n')
        f.write(f'     - System:  \n')
        if model(part_number) in model(parsed_system):
            f.write(f'       -- {parsed_system}  \n')
        else:
            f.write(f'       -- {parsed_system} \n')
        f.write(' ' + '-' * 118 + ' ' + '\n')
        empty_component[f' - {component} (Version)'] = f'System ({parsed_system}) | TRR '
        calculate_results(tally=0)


def qcl_firmware(trr, component, name, parsed_system, state, pipe_num):
    component_string = str(component)
    component_upper = component_string.upper()
    request = get_request_type(trr)
    trr_firmware = get_trr_firmware(trr)

    short_name = name[-3::]

    with open(f'{Path.reports}/pipe_{pipe_num}/{name}/text_{short_name}.txt', 'a') as f:
        f.write(f'  {state} - {component} Firmwares: \n')
        f.write(f'     - TRR QCL: \n')
        for item in trr:
            item_string = str(item)
            item_upper = item_string.upper()
            if 'QCL' in item_upper and component_upper in item_upper:
                if parsed_system in firmware(trr[item]):
                    f.write(f'       -- {firmware(trr[item])} << MATCHED \n')
                else:
                    f.write(f'       -- {firmware(trr[item])} \n')
        f.write(f'     - System:  \n')
        if firmware(trr_firmware) in firmware(parsed_system):
            f.write(f'       -- {parsed_system}  \n')
        else:
            f.write(f'       -- {parsed_system} \n')
        f.write(' ' + '-' * 118 + ' ' + '\n')
        empty_component[f' - {component} (Version)'] = f'System ({parsed_system}) | TRR '
        calculate_results(tally=0)


def get_nic_firmware(trr):
    for item in trr:
        if 'NIC' in item and 'Firmware' in item:
            nic = trr[item]
            return nic


def check_empty(parsed_trr, parsed_system):
    parsed_trr = str(parsed_trr)
    parsed_system = str(parsed_system)
    if parsed_trr == "" or parsed_system == "":
        return True
    elif parsed_trr.isspace() or parsed_system.isspace():
        return True
    else:
        return False

def is_not_blank(string):
    return bool(string and string.strip())


def get_raw_trr(trr, component):
    for item in trr:
        upper_component = str(component).upper()
        upper_item = str(item).upper()
        if upper_component in upper_item:
            raw_trr = trr[item]
            return str(raw_trr)


def return_trr_bios(trr, component):
    for item in trr:
        upper_component = str(component).upper()
        upper_item = str(item).upper()
        if upper_component in upper_item:
            raw_trr = trr[item]
            raw_trr = str(raw_trr).split(' ')[0]
            return raw_trr


def return_trr_os(trr, component):
    for item in trr:
        upper_component = ' ' + str(component).upper()
        upper_item = str(item).upper()
        if upper_component in upper_item and 'Server' in upper_item:
            raw_trr = trr[item]
            raw_trr = str(raw_trr).split(' ')
            return raw_trr


def engineering_group(name, parsed_trr, parsed_system, pipe_num):

    short_name = name[-3::]

    if parsed_trr != parsed_system:
        with open(f'{Path.reports}/pipe_{pipe_num}/{name}/text_{short_name}.txt', 'a') as f:
            f.write('   Reasons: \n')
            if parsed_trr[:5:] != parsed_system[:5:]:
                f.write(f'   - Systems: ({parsed_trr[:5:]}) | System ({parsed_system[:5:]}) \n')
            if parsed_trr[-3::] != parsed_system[-3::]:
                f.write(f'   - Engineering Groups: ({parsed_trr[-3::]}) | System ({parsed_system[-3::]}) \n')


def check_version(trr, component):
    for item in trr:
        item = str(item)
        if component.upper() in item.upper():
            if is_not_blank(trr[item]):
                raw_trr = (trr[item])
                return raw_trr
        else:
            raw_trr = 'No Matched Component'
            return raw_trr


def check_nvme_firmware(trr, component):
    for item in trr:
        item = str(item)
        if component.upper() in item.upper():
            if is_not_blank(trr[item]):
                raw_trr = (trr[item])
                raw_trr = raw_trr[-8::]
                return raw_trr
        else:
            raw_trr = 'No Matched Component'
            return raw_trr

# def results(system, rights, wrongs, total):
#     name = system['machine_name']
#     with open(f'{Path.reports}{name}.txt', 'a') as f:
#         f.write(f'  Matched Results: {sum(rights)} out of {sum(total)} \n')
#         for key, value in right_components.items():
#             statement = f' {key} \n'
#             f.write(statement)
#         # f.write(' ' + '-' * 59 + ' ' + '\n')
#         f.write(f'  Mismatched Results: {abs(sum(wrongs))} out of {sum(total)} \n')
#         for key, value in wrong_components.items():
#             statement = f' {key} \n'
#             f.write(statement)
#         # f.write(' ' + '-' * 59 + ' ' + '\n')
#         f.write(f'  Missing Info Results: {abs(sum(empty))} out of {sum(total)} \n')
#         for key, value in empty_component.items():
#             statement = f' {key} \n'
#             f.write(statement)
#         f.write('*' * 120 + '\n')


def results(system, rights, wrongs, total, pipe_num):

    name = system['machine_name']
    short_name = name[-3::]

    with open(f'{Path.reports}/pipe_{pipe_num}/{name}/text_{short_name}.txt', 'a') as f:
        f.write(f'  NOTE: Future Feature... \n')
        f.write(f'  Results: {sum(rights)} out of {sum(total)} matched. \n')
        for key, value in right_components.items():
            statement = f' {key}: {value} \n'
            f.write(statement)
        f.write(' ' + '-' * 118 + ' ' + '\n')
        f.write(f'  Results: {abs(sum(wrongs))} out of {sum(total)} are mismatched. \n')
        for key, value in wrong_components.items():
            statement = f' {key}: {value} \n'
            f.write(statement)
        f.write(' ' + '-' * 118 + ' ' + '\n')
        f.write(f'  Results: {abs(sum(empty))} out of {sum(total)} are missing information. \n')
        for key, value in empty_component.items():
            statement = f' {key}: {value} \n'
            f.write(statement)
        f.write('-' * 120 + '\n')


def append_qcl_missing_versions(component, name, raw_system, parsed_trr, parsed_system, list, pipe_num):

    short_name = name[-3::]

    with open(f'{Path.reports}/pipe_{pipe_num}/{name}/text_{short_name}.txt', 'a') as f:
        f.write(f'  >> MISSING INFORMATION - {component} Versions: \n')
        f.write(f'     - TRR QCL: \n')
        for item in list:
            f.write(f'       - : {item}\n')
        f.write(f'     - System: {raw_system} \n')
        f.write(' ' + '-' * 118 + ' ' + '\n')
        empty_component[f' - {component} (Version)'] = f'System ({parsed_system}) | TRR ({parsed_trr})'
        calculate_results(tally=0)


def append_qcl_matched_versions(component, name, raw_system, parsed_trr, parsed_system, trr, pipe_num):

    short_name = name[-3::]

    with open(f'{Path.reports}/pipe_{pipe_num}/{name}/text_{short_name}.txt', 'a') as f:
        f.write(f'  >> MATCHED - {component} Versions: \n')
        f.write(f'     - TRR QCL: \n')
        for item in trr:
            item_string = str(item)
            item_upper = item_string.upper()
            component_string = str(component)
            component_upper = component_string.upper()
            if 'QCL' in item_upper and component_upper in item_upper:
                for qcl in item_upper:
                    if parsed_system in qcl:
                        f.write(model(trr[item]) + ' << MATCHED')
                    else:
                        f.write(model(trr[item]))
        f.write(f'     - System: {raw_system} \n')
        f.write(' ' + '-' * 118 + ' ' + '\n')
        empty_component[f' - {component} (Version)'] = f'System ({parsed_system}) | TRR ({parsed_trr})'
        calculate_results(tally=0)


def append_qcl_mismatched_versions(component, name, raw_system, parsed_trr, parsed_system, trr, pipe_num):

    short_name = name[-3::]

    with open(f'{Path.reports}/pipe_{pipe_num}/{name}/text_{short_name}.txt', 'a') as f:
        f.write(f'  >> MISMATCHED - {component} Versions: \n')
        f.write(f'     - TRR QCL: \n')
        for item in trr:
            item_string = str(item)
            item_upper = item_string.upper()
            component_string = str(component)
            component_upper = component_string.upper()
            if 'QCL' in item_upper and component_upper in item_upper:
                for qcl in item_upper:
                    if parsed_system in qcl:
                        f.write(model(trr[item]) + ' << MATCHED')
                    else:
                        f.write(model(trr[item]))
        f.write(f'     - System: {raw_system} \n')
        f.write(' ' + '-' * 118 + ' ' + '\n')
        empty_component[f' - {component} (Version)'] = f'System ({parsed_system}) | TRR ({parsed_trr})'
        calculate_results(tally=0)


def append_qcl_except_mismatched_versions(component, name, raw_system, parsed_trr, parsed_system, trr, pipe_num):

    short_name = name[-3::]

    with open(f'{Path.reports}/pipe_{pipe_num}/{name}/text_{short_name}.txt', 'a') as f:
        f.write(f'  >> MATCHED - {component} Versions: \n')
        f.write(f'     - TRR QCL: \n')
        for item in trr:
            item_string = str(item)
            item_upper = item_string.upper()
            component_string = str(component)
            component_upper = component_string.upper()
            if 'QCL' in item_upper and component_upper in item_upper:
                for qcl in item_upper:
                    if parsed_system in qcl:
                        f.write(model(trr[item]) + ' << MATCHED')
                    else:
                        f.write(model(trr[item]))
        f.write(f'     - System: {raw_system} \n')
        f.write(' ' + '-' * 118 + ' ' + '\n')
        empty_component[f' - {component} (Version)'] = f'System ({parsed_system}) | TRR ({parsed_trr})'
        calculate_results(tally=0)


def append_empty_version(component, name, raw_trr, raw_system, parsed_trr, parsed_system, pipe_num):

    short_name = name[-3::]

    with open(f'{Path.reports}/pipe_{pipe_num}/{name}/text_{short_name}.txt', 'a') as f:
        f.write(f'  >> MISSING INFORMATION - {component} Versions: \n')
        f.write(f'     - TRR: {raw_trr} \n')
        f.write(f'     - System: {raw_system} \n')
        f.write(' ' + '-' * 118 + ' ' + '\n')
        empty_component[f' - {component} (Version)'] = f'System ({parsed_system}) | TRR ({parsed_trr})'
        calculate_results(tally=0)


def append_match_version(component, name, raw_trr, raw_system, parsed_trr, parsed_system, pipe_num):

    short_name = name[-3::]

    with open(f'{Path.reports}/pipe_{pipe_num}/{name}/text_{short_name}.txt', 'a') as f:
        f.write(f'     Match - {component} Versions: \n')
        f.write(f'     - TRR: {raw_trr} \n')
        f.write(f'     - System: {raw_system} \n')
        f.write(' ' + '-' * 118 + ' ' + '\n')
        right_components[f' - {component} (Version)'] = f'System ({parsed_system}) | TRR ({parsed_trr})'
        calculate_results(tally=1)


def append_mismatch_version(component, name, raw_trr, raw_system, parsed_trr, parsed_system, pipe_num):

    short_name = name[-3::]

    with open(f'{Path.reports}/pipe_{pipe_num}/{name}/text_{short_name}.txt', 'a') as f:
        f.write(f'  >> MISMATCH - {component} Versions: \n')
        f.write(f'     - TRR: {raw_trr} \n')
        f.write(f'     - System: {raw_system} \n')
        f.write(' ' + '-' * 118 + ' ' + '\n')
        wrong_components[f' - {component} (Version)'] = f'System ({parsed_system}) | TRR ({parsed_trr})'
        calculate_results(tally=-1)


def append_except_mismatch_version(component, name, raw_trr, raw_system, parsed_trr, parsed_system, pipe_num):

    short_name = name[-3::]

    with open(f'{Path.reports}/pipe_{pipe_num}/{name}/text_{short_name}.txt', 'a') as f:
        f.write(f'  >> MISMATCH - {component} Versions: \n')
        f.write(f'     - TRR: {raw_trr} \n')
        f.write(f'     - System: {raw_system} \n')
        f.write(' ' + '-' * 118 + ' ' + '\n')
        wrong_components[f' - {component} (Version)'] = f'System ({parsed_system}) | TRR ({parsed_trr})'
        calculate_results(tally=-1)


def append_empty_tpm_version(component, name, raw_trr, raw_system, parsed_trr, parsed_system, pipe_num):

    short_name = name[-3::]

    with open(f'{Path.reports}/pipe_{pipe_num}/{name}/text_{short_name}.txt', 'a') as f:
        f.write(f'  >> MISSING INFORMATION - {component} Versions: \n')
        f.write(f'     - TRR: {raw_trr} \n')
        f.write(f'     - System: {raw_system} \n')
        f.write(f'     NOTE: DO NOT update TPM, might brick blade \n')
        f.write(' ' + '-' * 118 + ' ' + '\n')
        empty_component[f' - {component} (Version)'] = f'System ({parsed_system}) | TRR ({parsed_trr})'
        calculate_results(tally=0)


def append_match_tpm_version(component, name, raw_trr, raw_system, parsed_trr, parsed_system, pipe_num):

    short_name = name[-3::]

    with open(f'{Path.reports}/pipe_{pipe_num}/{name}/text_{short_name}.txt', 'a') as f:
        f.write(f'     Match - {component} Versions: \n')
        f.write(f'     - TRR: {raw_trr} \n')
        f.write(f'     - System: {raw_system} \n')
        f.write(f'     NOTE: DO NOT update TPM, might brick blade \n')
        f.write(' ' + '-' * 118 + ' ' + '\n')
        right_components[f' - {component} (Version)'] = f'System ({parsed_system}) | TRR ({parsed_trr})'
        calculate_results(tally=1)


def append_mismatch_tpm_version(component, name, raw_trr, raw_system, parsed_trr, parsed_system, pipe_num):

    short_name = name[-3::]

    with open(f'{Path.reports}/pipe_{pipe_num}/{name}/text_{short_name}.txt', 'a') as f:
        f.write(f'  >> MISMATCH - {component} Versions: \n')
        f.write(f'     - TRR: {raw_trr} \n')
        f.write(f'     - System: {raw_system} \n')
        f.write(f'     NOTE: DO NOT update TPM, might brick blade \n')
        f.write(' ' + '-' * 118 + ' ' + '\n')
        wrong_components[f' - {component} (Version)'] = f'System ({parsed_system}) | TRR ({parsed_trr})'
        calculate_results(tally=-1)


def append_except_mismatch_tpm_version(component, name, raw_trr, raw_system, parsed_trr, parsed_system, pipe_num):

    short_name = name[-3::]

    with open(f'{Path.reports}/pipe_{pipe_num}/{name}/text_{short_name}.txt', 'a') as f:
        f.write(f'  >> MISMATCH - {component} Versions: \n')
        f.write(f'     - TRR: {raw_trr} \n')
        f.write(f'     - System: {raw_system} \n')
        f.write(f'     NOTE: DO NOT update TPM, might brick blade \n')
        f.write(' ' + '-' * 118 + ' ' + '\n')
        wrong_components[f' - {component} (Version)'] = f'System ({parsed_system}) | TRR ({parsed_trr})'
        calculate_results(tally=-1)


def append_empty_bios_version(component, name, raw_trr, raw_system, parsed_trr, parsed_system, pipe_num):

    short_name = name[-3::]

    with open(f'{Path.reports}/pipe_{pipe_num}/{name}/text_{short_name}.txt', 'a') as f:
        f.write(f'  >> MISSING INFORMATION - {component} Versions: \n')
        f.write(f'     - TRR: {raw_trr} \n')
        f.write(f'     - System: {raw_system} \n')
        f.write(' ' + '-' * 118 + ' ' + '\n')
        empty_component[f' - {component} (Version)'] = f'System ({parsed_system}) | TRR ({parsed_trr})'
        calculate_results(tally=0)


def append_match_bios_version(component, name, raw_trr, raw_system, parsed_trr, parsed_system, pipe_num):

    short_name = name[-3::]

    with open(f'{Path.reports}/pipe_{pipe_num}/{name}/text_{short_name}.txt', 'a') as f:
        f.write(f'     Match - {component} Versions: \n')
        f.write(f'     - TRR: {raw_trr} \n')
        f.write(f'     - System: {raw_system} \n')
        f.write(' ' + '-' * 118 + ' ' + '\n')
        right_components[f' - {component} (Version)'] = f'System ({parsed_system}) | TRR ({parsed_trr})'
        calculate_results(tally=1)


def append_mismatch_bios_version(component, name, raw_trr, raw_system, parsed_trr, parsed_system, pipe_num):

    short_name = name[-3::]

    with open(f'{Path.reports}/pipe_{pipe_num}/{name}/text_{short_name}.txt', 'a') as f:
        f.write(f'  >> MISMATCH - {component} Versions: \n')
        f.write(f'     - TRR: {raw_trr} \n')
        f.write(f'     - System: {raw_system} \n')
        if parsed_trr != parsed_system:
            f.write('     REASONS: \n')
            if parsed_trr[:5:] != parsed_system[:5:]:
                f.write(f'     - Systems: TRR ({parsed_trr[:5:]}) | System ({parsed_system[:5:]}) \n')
            if parsed_trr[9:-4:] != parsed_system[9:-4:]:
                f.write(f'     - Firmware Versions: TRR ({parsed_trr[9:-4:]}) | System ({parsed_system[9:-4:]}) \n')
            if parsed_trr[-3::] != parsed_system[-3::]:
                f.write(f'     - Engineering Groups: TRR ({parsed_trr[-3::]}) | System ({parsed_system[-3::]}) \n')
        f.write(' ' + '-' * 118 + ' ' + '\n')
        wrong_components[f' - {component} (Version)'] = f'System ({parsed_system}) | TRR ({parsed_trr})'
        calculate_results(tally=-1)


def append_except_mismatch_bios_version(component, name, raw_trr, raw_system, parsed_trr, parsed_system, pipe_num):

    short_name = name[-3::]

    with open(f'{Path.reports}/pipe_{pipe_num}/{name}/text_{short_name}.txt', 'a') as f:
        f.write(f'  >> MISMATCH - {component} Versions: \n')
        f.write(f'     - TRR: {raw_trr} \n')
        f.write(f'     - System: {raw_system} \n')
        engineering_group(name, parsed_trr, parsed_system, pipe_num)
        f.write(' ' + '-' * 118 + ' ' + '\n')
        wrong_components[f' - {component} (Version)'] = f'System ({parsed_system}) | TRR ({parsed_trr})'
        calculate_results(tally=-1)


def append_empty_bmc_version(component, name, raw_trr, raw_system, parsed_trr, parsed_system, pipe_num):

    short_name = name[-3::]

    with open(f'{Path.reports}/pipe_{pipe_num}/{name}/text_{short_name}.txt', 'a') as f:
        f.write(f'  >> MISSING INFORMATION - {component} Versions: \n')
        f.write(f'     - TRR: {raw_trr} \n')
        f.write(f'     - System: {raw_system} \n')
        f.write(' ' + '-' * 118 + ' ' + '\n')
        empty_component[f' - {component} (Version)'] = f'System ({parsed_system}) | TRR ({parsed_trr})'
        calculate_results(tally=0)


def append_match_bmc_version(component, name, raw_trr, raw_system, parsed_trr, parsed_system, pipe_num):

    short_name = name[-3::]

    with open(f'{Path.reports}/pipe_{pipe_num}/{name}/text_{short_name}.txt', 'a') as f:
        f.write(f'     Match - {component} Versions: \n')
        f.write(f'     - TRR: {raw_trr} \n')
        f.write(f'     - System: {raw_system} \n')
        f.write(' ' + '-' * 118 + ' ' + '\n')
        right_components[f' - {component} (Version)'] = f'System ({parsed_system}) | TRR ({parsed_trr})'
        calculate_results(tally=1)


def append_mismatch_bmc_version(component, name, raw_trr, raw_system, parsed_trr, parsed_system, pipe_num):

    short_name = name[-3::]

    with open(f'{Path.reports}/pipe_{pipe_num}/{name}/text_{short_name}.txt', 'a') as f:
        f.write(f'  >> MISMATCH - {component} Versions: \n')
        f.write(f'     - TRR: {raw_trr} \n')
        f.write(f'     - System: {raw_system} \n')
        if parsed_trr != parsed_system:
            f.write('     REASON: \n')
            if parsed_trr[10:-3:] != parsed_system:
                f.write(f'     - Firmware Versions: TRR ({parsed_trr[10:-3:]}) | System ({parsed_system}) \n')
        f.write(' ' + '-' * 118 + ' ' + '\n')
        wrong_components[f' - {component} (Version)'] = f'System ({parsed_system}) | TRR ({parsed_trr})'
        calculate_results(tally=-1)


def append_except_mismatch_bmc_version(component, name, raw_trr, raw_system, parsed_trr, parsed_system, pipe_num):

    short_name = name[-3::]

    with open(f'{Path.reports}/pipe_{pipe_num}/{name}/text_{short_name}.txt', 'a') as f:
        f.write(f'  >> MISMATCH - {component} Versions: \n')
        f.write(f'     - TRR: {raw_trr} \n')
        f.write(f'     - System: {raw_system} \n')
        engineering_group(name, parsed_trr, parsed_system, pipe_num)
        f.write(' ' + '-' * 118 + ' ' + '\n')
        wrong_components[f' - {component} (Version)'] = f'System ({parsed_system}) | TRR ({parsed_trr})'
        calculate_results(tally=-1)


def append_empty_firmware(component, name, raw_trr, raw_system, parsed_trr, parsed_system, pipe_num):

    short_name = name[-3::]

    with open(f'{Path.reports}/pipe_{pipe_num}/{name}/text_{short_name}.txt', 'a') as f:
        f.write(f'  >> MISSING INFORMATION - {component} Firmwares: \n')
        f.write(f'     - TRR: {raw_trr} \n')
        f.write(f'     - System: {raw_system} \n')
        f.write(' ' + '-' * 118 + ' ' + '\n')
        empty_component[f' - {component} (Firmware)'] = f'System ({parsed_system}) | TRR ({parsed_trr})'
        calculate_results(tally=0)


def append_match_firmware(component, name, raw_trr, raw_system, parsed_trr, parsed_system, pipe_num):

    short_name = name[-3::]

    with open(f'{Path.reports}/pipe_{pipe_num}/{name}/text_{short_name}.txt', 'a') as f:
        f.write(f'     Match - {component} Firmwares: \n')
        f.write(f'     - TRR: {raw_trr} \n')
        f.write(f'     - System: {raw_system} \n')
        f.write(' ' + '-' * 118 + ' ' + '\n')
        right_components[f' - {component} (Firmware)'] = f'System ({parsed_system}) | TRR ({parsed_trr})'
        calculate_results(tally=1)


def append_mismatch_firmware(component, name, raw_trr, raw_system, parsed_trr, parsed_system, pipe_num):

    short_name = name[-3::]

    with open(f'{Path.reports}/pipe_{pipe_num}/{name}/text_{short_name}.txt', 'a') as f:
        f.write(f'  >> MISMATCH - {component} Firmwares: \n')
        f.write(f'     - TRR: {raw_trr} \n')
        f.write(f'     - System: {raw_system} \n')
        f.write(' ' + '-' * 118 + ' ' + '\n')
        wrong_components[f' - {component} (Firmware)'] = f'System ({parsed_system}) | TRR ({parsed_trr})'
        calculate_results(tally=-1)


def append_except_mismatch_firmware(component, name, raw_trr, raw_system, parsed_trr, parsed_system, pipe_num):

    short_name = name[-3::]

    with open(f'{Path.reports}/pipe_{pipe_num}/{name}/text_{short_name}.txt', 'a') as f:
        f.write(f'  >> MISMATCH - {component} Firmwares: \n')
        f.write(f'     - TRR: {raw_trr} \n')
        f.write(f'     - System: {raw_system} \n')
        f.write(' ' + '-' * 118 + ' ' + '\n')
        wrong_components[f' - {component} (Firmware)'] = f'System ({parsed_system}) | TRR ({parsed_trr})'
        calculate_results(tally=-1)


def append_os_empty(name, raw_trr, raw_system, parsed_trr, parsed_system, pipe_num):

    short_name = name[-3::]

    with open(f'{Path.reports}/pipe_{pipe_num}/{name}/text_{short_name}.txt', 'a') as f:
        f.write('  >> MISSING INFORMATION - OS Versions: \n')
        f.write(f'    - TRR: {raw_trr} \n')
        f.write(f'    - System: {raw_system} \n ')
        f.write(' ' + '-' * 118 + ' ' + '\n')
        empty_component[' - OS (Version)'] = f'System ({parsed_system}) | TRR ({parsed_trr})'
        calculate_results(tally=0)


def append_os_match(name, raw_trr, raw_system, parsed_trr, parsed_system, pipe_num):

    short_name = name[-3::]

    with open(f'{Path.reports}/pipe_{pipe_num}/{name}/text_{short_name}.txt', 'a') as f:
        f.write('     Match - OS Versions: \n')
        f.write(f'     - TRR: {raw_trr} \n')
        f.write(f'     - System: {raw_system} \n ')
        f.write(' ' + '-' * 118 + ' ' + '\n')
        right_components[' - OS (Version)'] = f'System ({parsed_system}) | TRR ({parsed_trr})'
        calculate_results(tally=1)


def append_os_mismatch(name, raw_trr, raw_system, parsed_trr, parsed_system, pipe_num):

    short_name = name[-3::]

    with open(f'{Path.reports}/pipe_{pipe_num}/{name}/text_{short_name}.txt', 'a') as f:
        f.write('  >> MISMATCH - OS Versions: \n')
        f.write(f'     - TRR: {raw_trr} \n')
        f.write(f'     - System: {raw_system} \n ')
        f.write(' ' + '-' * 118 + ' ' + '\n')
        wrong_components[' - OS (Version)'] = f'System ({parsed_system}) | TRR ({parsed_trr})'
        calculate_results(tally=-1)


def append_except_os_mismatch(name, raw_trr, raw_system, parsed_trr, parsed_system, pipe_num):

    short_name = name[-3::]

    with open(f'{Path.reports}/pipe_{pipe_num}/{name}/text_{short_name}.txt', 'a') as f:
        f.write('  >> MISMATCH - OS Versions: \n')
        f.write(f'     - TRR: {raw_trr} \n')
        f.write(f'     - System: {raw_system} \n ')
        f.write(' ' + '-' * 118 + ' ' + '\n')
        wrong_components[' - OS (Version)'] = f'System ({parsed_system}) | TRR ({parsed_trr})'
        calculate_results(tally=-1)


def get_target_configuration(trr):
    for item in trr:
        item_upper = str(item)
        item_upper = item_upper.upper()
        if 'TARGET' in item_upper and 'CONFIGURATION' in item_upper:
            return trr[item]


def get_request_type(trr):
    for item in trr:
        item_upper = str(item)
        item_upper = item_upper.upper()
        if 'REQUEST' in item_upper and 'TYPE' in item_upper:
            return trr[item]


def get_part_number(trr):
    for item in trr:
        item_upper = str(item)
        item_upper = item_upper.upper()
        if 'PART' in item_upper and 'NUMBER' in item_upper:
            return trr[item]


def get_trr_description(trr):
    for item in trr:
        item_upper = str(item)
        item_upper = item_upper.upper()
        if 'DESCRIPTION' == item_upper:
            return trr[item]


def get_trr_firmware(trr):
    for item in trr:
        item_upper = str(item)
        item_upper = item_upper.upper()
        if 'FIRMWARE' == item_upper:
            return trr[item]


def write_pipe_folder(pipe_num):
    try:
        os.mkdir(f'reports/pipe_{pipe_num}')
    except FileExistsError:
        pass


def write_machine_folder(pipe_num, name):
    try:
        os.mkdir(f'reports/pipe_{pipe_num}/{name}')
    except FileExistsError:
        pass


def intro(pipe_num, trr_id, trr, system):
    target = get_target_configuration(trr)
    request = get_request_type(trr)
    part_number = get_part_number(trr)
    trr_description = replace_company_info(get_trr_description(trr))
    trr_firmware = get_trr_firmware(trr)
    rack = system['location']
    try:
        name = system['machine_name']
        short_name = name[-3::]
        date_time = strftime('  Time: %m/%d/%Y - %I:%M %p')
        write_pipe_folder(pipe_num)
        write_machine_folder(pipe_num, name)
        with open(f'{Path.reports}/pipe_{pipe_num}/{name}/text_{short_name}.txt', 'w') as f:
            f.write('-' * 120 + '\n')
            f.write(f'  TRR Info: {trr_id} \n')
            f.write(f'  System Info: {name} \n')
            f.write(f'{date_time} \n')
            f.write('  Location: VSE0 - Kirkland \n')
            f.write(' ' + '-' * 118 + ' ' + '\n')
            f.write(f'  Pipe: (Future Feature) \n')
            f.write(f'  Rack: {rack} \n')
            f.write('  Engineer: (Future Feature) \n')
            f.write('  Technician: (Future Feature) \n')
            f.write('  Progress: (Future Feature) \n')
            f.write(' ' + '-' * 118 + ' ' + '\n')
            f.write(f'  Request Type: {request} \n')
            f.write(f'  Target Configuration: {target} \n')
            f.write(f'  Part Number: {model(part_number)} \n')
            f.write(f'  Description: {trr_description} \n')
            f.write(f'  Firmware: {trr_firmware} \n')
            f.write('-' * 120 + '\n \n')
    except KeyError:
        pass


def intro_two(pipe_num, trr_id, trr, system):
    target = get_target_configuration(trr)
    request = get_request_type(trr)
    part_number = get_part_number(trr)
    trr_description = replace_company_info(get_trr_description(trr))
    trr_firmware = get_trr_firmware(trr)
    rack = system['location']
    try:
        name = system['machine_name']
        short_name = name[-3::]
        time = strftime('%I:%M %p')
        date = strftime('%m/%d/%Y')
        with open(f'{Path.reports}/pipe_{pipe_num}/{name}/text_{short_name}.txt', 'w') as f:
            f.write('-' * 120 + '\n')
            f.write(f'  Pipe: N/A' + ' ' * 97 + f'{date} \n')
            f.write(f'  TRR: {trr_id}' + ' ' * 97 + f'{time} \n')
            f.write(f'  System: {str(name)[-7::]}' + ' ' * 88 + 'VSE0-Kirkland \n')
            f.write(' ' + '-' * 118 + ' ' + '\n')
            # f.write(f'  Rack: {rack} \n')
            f.write('  Progress: N/A ' + ' ' * 80 + 'N/A % - TRR Completion \n')
            f.write('  Engineer: N/A ' + ' ' * 80 + 'N/A % - Hardware Setup \n')
            f.write('  Technician: N/A ' + ' ' * 78 + 'N/A % - Software Setup \n')
            f.write(' ' + '-' * 118 + ' ' + '\n')
            f.write(f'  Request Type: {request} \n')
            f.write(f'  Part Number: {model(part_number)} \n')
            # f.write(f'  Description: {trr_description} \n')
            f.write(f'  Firmware: {trr_firmware} \n')
            f.write(f'  Target Configuration: {target} \n')
            f.write(' ' + '-' * 118 + ' ' + '\n')
    except KeyError:
        pass


def bios_version(trr, system, pipe_num):
    try:
        component = 'BIOS'
        name = str(system['machine_name'])
        raw_trr = str(return_trr_bios(trr, component))
        raw_system = str(system['dmi']['bios']['version'])
        parsed_trr = raw_trr
        parsed_system = str(system['dmi']['bios']['version'])
        if check_empty(parsed_trr, parsed_system) == True:
            append_empty_bios_version(component, name, raw_trr, raw_system, parsed_trr, parsed_system, pipe_num)
        elif parsed_trr in parsed_system or parsed_system in parsed_trr:
            append_match_bios_version(component, name, raw_trr, raw_system, parsed_trr, parsed_system, pipe_num)
        else:
            append_mismatch_bios_version(component, name, raw_trr, raw_system, parsed_trr, parsed_system, pipe_num)
    except KeyError:
        append_except_mismatch_bios_version(component, name, raw_trr, raw_system, parsed_trr, parsed_system, pipe_num)


def bmc_version(trr, system, pipe_num):
    try:
        component = 'BMC'
        name = str(system['machine_name'])
        raw_trr = get_raw_trr(trr, component)
        raw_system = str(system['bmc']['mc']['firmware'])
        parsed_trr = str(raw_trr)
        parsed_system = str(system['bmc']['mc']['firmware']).replace('.', '')
        if check_empty(parsed_trr, parsed_system) == True:
            append_empty_bmc_version(component, name, raw_trr, raw_system, parsed_trr, parsed_system, pipe_num)
        elif parsed_trr in parsed_system or parsed_system in parsed_trr:
            append_match_bmc_version(component, name, raw_trr, raw_system, parsed_trr, parsed_system, pipe_num)
        else:
            append_mismatch_bmc_version(component, name, raw_trr, raw_system, parsed_trr, parsed_system, pipe_num)
    except KeyError:
        append_except_mismatch_bmc_version(component, name, raw_trr, raw_system, parsed_trr, parsed_system, pipe_num)


def cpld_version(trr, system, pipe_num):
    # try:
    component = 'CPLD'
    name = str(system['machine_name'])
    raw_trr = str(get_raw_trr(trr, component))
    raw_system = 'Unavailable'
    parsed_trr = raw_trr
    parsed_system = 'Unavailable'

    name = system['machine_name']
    short_name = name[-3::]

    with open(f'{Path.reports}/pipe_{pipe_num}/{name}/text_{short_name}.txt', 'a') as f:
        f.write(f'     Unavailable - {component} Firmwares: \n')
        f.write(f'     - TRR: {raw_trr} \n')
        f.write(f'     - System: Unavailable \n')
        f.write(f'     REASON: \n')
        f.write(f'     - Not yet available \n')
        f.write(' ' + '-' * 118 + ' ' + '\n')
    #     if check_empty(parsed_trr, parsed_system) == True:
    #         append_empty_version(component, name, raw_trr, raw_system, parsed_trr, parsed_system)
    #     elif parsed_trr in parsed_system or parsed_system in parsed_trr:
    #         append_match_version(component, name, raw_trr, raw_system, parsed_trr, parsed_system)
    #     else:
    #         append_mismatch_version(component, name, raw_trr, raw_system, parsed_trr, parsed_system)
    # except KeyError:
    #     append_except_mismatch_version(component, name, raw_trr, raw_system, parsed_trr, parsed_system)


def nic_firmware(trr, system, pipe_num):
    # try:
    component = 'NIC'
    name = str(system['machine_name'])
    raw_trr = str(get_nic_firmware(trr))
    raw_system = 'Unavailable'
    parsed_trr = raw_trr
    parsed_system = 'Unavailable'

    name = system['machine_name']
    short_name = name[-3::]

    with open(f'{Path.reports}/pipe_{pipe_num}/{name}/text_{short_name}.txt', 'a') as f:
        f.write(f'     Unavailable - {component} Firmwares: \n')
        f.write(f'     - TRR: {raw_trr} \n')
        f.write(f'     - System: Unavailable \n')
        f.write(f'     REASON: \n')
        f.write(f'     - Not yet available \n')
        f.write(' ' + '-' * 118 + ' ' + '\n')
    #     if check_empty(parsed_trr, parsed_system) == True:
    #         append_empty_version(component, name, raw_trr, raw_system, parsed_trr, parsed_system)
    #     elif parsed_trr in parsed_system or parsed_system in parsed_trr:
    #         append_match_version(component, name, raw_trr, raw_system, parsed_trr, parsed_system)
    #     else:
    #         append_mismatch_version(component, name, raw_trr, raw_system, parsed_trr, parsed_system)
    # except KeyError:
    #     append_except_mismatch_version(component, name, raw_trr, raw_system, parsed_trr, parsed_system)


def hdd_version(trr, system, pipe_num):
    try:
        component = "HDD"
        name = str(system['machine_name'])
        parsed_system = model(system['disk']['unique_disks']['0']['model'])
        state = check_qcl(trr, component, parsed_system)
        qcl_versions(trr, component, name, parsed_system, state, pipe_num)
    except KeyError:
        pass


def hdd_firmware(trr, system, pipe_num):
    try:
        component = "HDD"
        name = str(system['machine_name'])
        parsed_system = system['disk']['unique_disks']['0']['firmware']
        state = check_qcl(trr, component, parsed_system)
        qcl_firmware(trr, component, name, parsed_system, state, pipe_num)
    except KeyError:
        pass


def ssd_version(trr, system, pipe_num):
    try:
        component = "SSD"
        name = str(system['machine_name'])
        parsed_system = model(system['disk']['unique_disks']['0']['model'])
        state = check_qcl(trr, component, parsed_system)
        qcl_versions(trr, component, name, parsed_system, state, pipe_num)
    except KeyError:
        pass


def ssd_firmware(trr, system, pipe_num):
    try:
        component = "SSD"
        name = str(system['machine_name'])
        parsed_system = str(system['disk']['unique_disks']['0']['firmware'])
        state = check_qcl(trr, component, parsed_system)
        qcl_firmware(trr, component, name, parsed_system, state, pipe_num)
    except KeyError:
        pass


def dimm_version(trr, system, pipe_num):
    try:
        component = "DIMM"
        name = str(system['machine_name'])
        parsed_system = str(system['dmi']['unique_dimms']['0']['part'])
        state = check_qcl(trr, component, parsed_system)
        qcl_versions(trr, component, name, parsed_system, state, pipe_num)
    except KeyError:
        pass


def nvme_version(trr, system, pipe_num):
    try:
        component = "NVMe"
        name = str(system['machine_name'])
        parsed_system = str(model(system['nvme']['unique_nvmes']['0']['model']))
        state = check_qcl(trr, component, parsed_system)
        qcl_versions(trr, component, name, parsed_system, state, pipe_num)
    except KeyError:
        pass


def nvme_firmware(trr, system, pipe_num):
    try:
        component = "NVMe"  # stay
        name = str(system['machine_name'])  # stay
        parsed_system = str(system['nvme']['unique_nvmes']['0']['firmware'])
        state = check_qcl(trr, component, parsed_system)
        qcl_firmware(trr, component, name, parsed_system, state, pipe_num)
    except KeyError:
        pass


def tpm_version(trr, system, pipe_num):
    try:
        component = 'TPM'
        name = str(system['machine_name'])
        raw_trr = str(get_raw_trr(trr, component))
        raw_system = str(system['tpm']['version'])
        parsed_trr = raw_trr
        parsed_system = str(system['tpm']['version'][:4:])
        if check_empty(parsed_trr, parsed_system) == True:
            append_empty_tpm_version(component, name, raw_trr, raw_system, parsed_trr, parsed_system, pipe_num)
        elif parsed_trr in parsed_system or parsed_system in parsed_trr:
            append_match_tpm_version(component, name, raw_trr, raw_system, parsed_trr, parsed_system, pipe_num)
        else:
            append_mismatch_tpm_version(component, name, raw_trr, raw_system, parsed_trr, parsed_system, pipe_num)
    except KeyError:
        append_except_mismatch_tpm_version(component, name, raw_trr, raw_system, parsed_trr, parsed_system, pipe_num)


def get_os_version(trr):
    for item in trr:
        item_upper = str(item)
        item_upper = item_upper.upper()
        if 'SERVER' in item_upper and 'OS' in item_upper:
            if 'BI' not in item_upper:
                raw = trr[item]
                parsed = replace_unnecesarry(raw.split(' ')[-1])
                return parsed


def os_version(trr, system, pipe_num):
    try:
        name = str(system['machine_name'])
        parsed_trr = str(get_os_version(trr))
        parsed = str(system['platform']['version'])
        parsed_system = str(replace_unnecesarry(parsed.split('.')[-1]))
        if check_empty(parsed_trr, parsed_system) == True:
            append_os_empty(name, parsed_trr, parsed_system, parsed_trr, parsed_system, pipe_num)
        elif parsed_trr in parsed_system or parsed_system in parsed_trr or '2019' in parsed_trr:
            append_os_match(name, parsed_trr, parsed_system, parsed_trr, parsed_system, pipe_num)
        else:
            append_os_mismatch(name, parsed_trr, parsed_system, parsed_trr, parsed_system, pipe_num)
    except KeyError:
        append_except_os_mismatch(name, parsed_trr, parsed_system, parsed_trr, parsed_system, pipe_num)


def calculate_results(tally):
    if tally == 0:
        tally = + 1
        empty.append(tally)
    elif tally > 0:
        rights.append(tally)
    else:
        wrongs.append(tally)
    add_total = + 1
    total.append(add_total)


def hardware_configuration(trr, system, pipe_num):
    try:
        name = system['machine_name']
        short_name = name[-3::]
        hdd_name = system['disk']['unique_disks']['0']['model']
        hdd_count = system['disk']['unique_disks']['0']['count']
        dimm_name = system['dmi']['unique_dimms']['0']['part']
        dimm_count = system['dmi']['unique_dimms']['0']['count']
        nvme_name = system['nvme']['unique_nvmes']['0']['model']
        nvme_count = system['nvme']['unique_nvmes']['0']['count']
        rack = system['location']
        target = get_target_configuration(trr)

        with open(f'{Path.reports}/pipe_{pipe_num}/{name}/text_{short_name}.txt', 'a') as f:
            f.write(f'  SKU: N/A \n')
            f.write(f'  Rack Location: {rack} \n')
            # f.write(f'  Target Configuration: {target} \n')
            f.write(f'  - HDD/ SSD Name: {hdd_name} - {hdd_count} count' + ' ' * 51 + '\n')
            f.write(f'  - DIMM Name: {dimm_name} - {dimm_count} count' + ' ' * 59 + '\n')
            f.write(f'  - NVMe Name: {nvme_name} - {nvme_count} count' + ' ' * 49 + '\n')
            f.write('-' * 120 + '\n')
    except KeyError:
        pass


def end(system):
    name = system['machine_name']
    print('\n')
    input(f'\n   {name} created inside the reports folder. \n   Press enter to exit... ')
    # os.system(f'notepad.exe reports/{name}.txt')
    sys.exit('   Exiting...')
    # os.open(f'reports/{name}.txt')


def check_sku(trr):
    whole = []
    x = 0

    target = str(get_target_configuration(trr))
    target = str(target.replace(']', ''))
    target = target.split('[')

    while x < len(target):
        if ' ' in target[x]:
            split = str(target[x]).split(' ')
            for item in split:
                whole.append(item)
        else:
            whole.append(target[x])
        x += 1

    whole = list(filter(None, whole))
    # print(whole)

    x = 0

    # print(whole)
    #
    # while x < len(whole):
    #     if whole[x] in new_data[0: None]:
    #         print(x)
    #     x += 1

    # for part in whole:
        # if 'Gen' in part:
        #     print(part)


    with open('setup/skudoc.csv', 'r') as f:

        for configuration in f:
            x = 0
            while x < 20:
                name = configuration.split(',')[x]