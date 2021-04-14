import pandas as pd


all_dict = {}
gen_list = []
azure_list = []
blade_list = []
server_list = []
bios_list = []
bmc_list = []
tpm_list = []
cpld_list = []
chipset_driver_list = []
server_processor_list = []
fpga_release_package_list = []
fpga_hyperblaster_list = []
fpga_hip_list = []
fpga_filter_list = []
ftdi_port_list = []
ftdi_bus_list = []
nic_firmware_list = []
nic_pxe_list = []
nic_uefi_list = []
nic_driver_list = []
boot_drive_list = []
nvme_list = []
hdd_list = []
psu_list = []


def create_csv_from_crd():
    crd = 'crd'

    cover_page_df = pd.read_excel(f'input/{crd}.xlsx', sheet_name='Cover Page')
    cover_page_df.to_csv('pipe_cleaner/crd_info/cover_page.csv', index=False)

    configuration_df = pd.read_excel(f'input/{crd}.xlsx', sheet_name='FW-SW Configuration')
    configuration_df.to_csv('pipe_cleaner/crd_info/fw-sw_configuration.csv', index=False)


def strip_cell(csv_cell: str) -> str:
    """
    Strips csv cell of unnecessary unicode information including newlines, system types, and commas.

    :param csv_cell: cell within the csv created by Pandas
    :type csv_cell: str
    :return: stripped cell
    :rtype: str
    """
    stripped_cell = str(csv_cell).replace(',', ' ')
    stripped_cell = stripped_cell.replace('BMC', '')
    stripped_cell = stripped_cell.replace('JBOF/F2010 ', '')
    stripped_cell = stripped_cell.replace('C2010 ', '')
    stripped_cell = stripped_cell.replace('F2010 ', '')
    stripped_cell = stripped_cell.replace(' ', '')

    # Azure
    stripped_cell = stripped_cell.replace('AZURE', 'Azure ')
    stripped_cell = stripped_cell.replace('XIO', 'XIO ')
    stripped_cell = stripped_cell.replace('STORAGE', 'Storage ')
    stripped_cell = stripped_cell.replace('GEN', 'Gen')

    # Blades
    stripped_cell = stripped_cell.replace('STORAGE', 'Storage ')
    stripped_cell = stripped_cell.replace('UTILITY', 'Utility ')
    stripped_cell = stripped_cell.replace('BLADEMBBIOSC2010.BS.3F42.GN1', 'Blade')

    # BIOS
    stripped_cell = stripped_cell.replace('BSL', '')
    stripped_cell = stripped_cell.replace('BIOS', '')
    stripped_cell = stripped_cell.replace('MB"', '')
    stripped_cell = stripped_cell.replace('CMCMC_', '')

    # BMC
    stripped_cell = stripped_cell.replace('SERVERSOFTWARE(MB)"', '')

    # TPM
    stripped_cell = stripped_cell.replace('TPM', '')

    # CPLD
    stripped_cell = stripped_cell.replace('CPLD', '')
    stripped_cell = stripped_cell.replace('(Storage )', '')
    stripped_cell = stripped_cell.replace('(Utility )', '')

    # Chipset Driver
    stripped_cell = stripped_cell.replace('INTELCHIPSETOSDRIVER', '')

    # FPGA Package
    stripped_cell = stripped_cell.replace('FPGALONGSPEAK', '')

    # NIC Firmware
    stripped_cell = stripped_cell.replace('MELLANOXNICCARD-FWREVCX4ONLP', '')

    # NIC PXE
    stripped_cell = stripped_cell.replace('MELLANOXNICCARD-PXEREV', '')

    # NVMe
    stripped_cell = stripped_cell.replace('NVME', '')
    stripped_cell = stripped_cell.replace('SSD', '')
    stripped_cell = stripped_cell.replace('SKHYNIXPE4010', '')

    # HDD
    stripped_cell = stripped_cell.replace('HDD', '')
    stripped_cell = stripped_cell.replace('SATA', '')
    stripped_cell = stripped_cell.replace('(10TB)', '')
    stripped_cell = stripped_cell.replace('HITACHI', '')
    stripped_cell = stripped_cell.replace('/', '')
    stripped_cell = stripped_cell.replace('0F27479', '')
    stripped_cell = stripped_cell.replace('FW:', '')
    stripped_cell = stripped_cell.replace('LHMST2J0', '')
    stripped_cell = stripped_cell.replace('512E-SAGE', '')
    stripped_cell = stripped_cell.replace('MOUNTAIN', '')
    stripped_cell = stripped_cell.replace('(B821G3-CM)', '')

    # Previous Work
    stripped_cell = stripped_cell.replace('SAMSUNG', '')
    stripped_cell = stripped_cell.replace('SEAGATE', '')
    stripped_cell = stripped_cell.replace('SKHYNIX', '')
    stripped_cell = stripped_cell.replace('INTEL', '')
    stripped_cell = stripped_cell.replace('P4511', '')
    stripped_cell = stripped_cell.replace('WD', '')
    stripped_cell = stripped_cell.replace('MICRON', '')
    stripped_cell = stripped_cell.replace('12TB', '')
    stripped_cell = stripped_cell.replace('PE6011', '')
    stripped_cell = stripped_cell.replace('PM983', '')
    stripped_cell = stripped_cell.replace('PM883', '')
    stripped_cell = stripped_cell.replace('HYNIX', '')
    stripped_cell = stripped_cell.replace('SE4011', '')
    stripped_cell = stripped_cell.replace('PE4010', '')
    stripped_cell = stripped_cell.replace('32GB', '')
    stripped_cell = stripped_cell.replace('PM983', '')
    stripped_cell = stripped_cell.replace('LITEON', '')
    stripped_cell = stripped_cell.replace('PM963', '')
    stripped_cell = stripped_cell.replace('4TB', '')
    stripped_cell = stripped_cell.replace('8TB', '')
    stripped_cell = stripped_cell.replace('32GB', '')
    stripped_cell = stripped_cell.replace('960GB', '')
    stripped_cell = stripped_cell.replace('_5200_', '')

    stripped_cell = stripped_cell.rstrip('\n')

    return stripped_cell


def get_azure_from_csv(component_1: str, component_2: str, component_row: str):
    upper_component_1 = str(component_1).upper()
    upper_component_2 = str(component_2).upper()
    upper_row = str(component_row).upper()
    lower_component = str(component_1).lower()

    if 'AZURE' in upper_component_1 and 'GEN' in upper_component_2 \
            and upper_component_1 in upper_row and upper_component_2 in upper_row:
        azure_list.append(lower_component)


def get_components_from_csv(component_1: str, component_2: str, component_row: str, component_list: list):
    """
    Gets the components' values from the created SW-FW configuration CSV file. Appends to all components within Data.
    :param component_2: 
    :param component_1: 
    :param component_row: row within the data from the CSV file
    :param component_list: list used to store component data
    :return:
    """

    # Changes the component and row to all caps in order to compare since Python is case sensitive
    upper_component_1 = str(component_1).upper()
    upper_component_2 = str(component_2).upper()
    upper_row = str(component_row).upper()
    lower_component = str(component_1).lower()

    if 'GEN' in upper_component_1 and '.' in upper_component_2 \
            and upper_component_1 in upper_row and upper_component_2 in upper_row:
        component_list.append(strip_cell(upper_row)[::])
    if 'AZURE' in upper_component_1 and 'GEN' in upper_component_2 \
            and upper_component_1 in upper_row and upper_component_2 in upper_row:
        component_list.append(strip_cell(upper_row)[::])
    if 'SERVER' in upper_component_1 and 'SERVER' in upper_component_2 \
            and upper_component_1 in upper_row and upper_component_2 in upper_row:
        component_list.append(strip_cell(upper_row)[::])
    if 'BIOS' in upper_component_1 and 'BS.' in upper_component_2 \
            and upper_component_1 in upper_row and upper_component_2 in upper_row:
        component_list.append(strip_cell(upper_row)[::])
    if 'BMC' in upper_component_1 and 'BC.' in upper_component_2 \
            and upper_component_1 in upper_row and upper_component_2 in upper_row:
        component_list.append(strip_cell(upper_row)[::])
    if 'TPM' in upper_component_1 and 'TPM' in upper_component_2 \
            and upper_component_1 in upper_row and upper_component_2 in upper_row:
        component_list.append(strip_cell(upper_row)[::])
    if 'CPLD' in upper_component_1 and 'CPLD' in upper_component_2 \
            and upper_component_1 in upper_row and upper_component_2 in upper_row:
        component_list.append(strip_cell(upper_row)[:11:])
    if 'CHIPSET' in upper_component_1 and 'CHIPSET' in upper_component_2 \
            and upper_component_1 in upper_row and upper_component_2 in upper_row:
        component_list.append(strip_cell(upper_row)[::])
    if 'FPGA' in upper_component_1 and 'FPGA' in upper_component_2 \
            and upper_component_1 in upper_row and upper_component_2 in upper_row:
        component_list.append(strip_cell(upper_row)[::])
    if 'HYPERBLASTER' in upper_component_1 and 'DLL' in upper_component_2 \
            and upper_component_1 in upper_row and upper_component_2 in upper_row:
        component_list.append(strip_cell(upper_row)[::])
    if 'FPGA' in upper_component_1 and 'HIP' in upper_component_2 \
            and upper_component_1 in upper_row and upper_component_2 in upper_row:
        component_list.append(strip_cell(upper_row)[::])
    if 'FPGA' in upper_component_1 and 'FILTER' in upper_component_2 \
            and upper_component_1 in upper_row and upper_component_2 in upper_row:
        component_list.append(strip_cell(upper_row)[::])
    if 'FTDI' in upper_component_1 and 'PORT' in upper_component_2 \
            and upper_component_1 in upper_row and upper_component_2 in upper_row:
        component_list.append(strip_cell(upper_row)[::])
    if 'FTDI' in upper_component_1 and 'BUS' in upper_component_2 \
            and upper_component_1 in upper_row and upper_component_2 in upper_row:
        component_list.append(strip_cell(upper_row)[::])
    if 'NIC' in upper_component_1 and 'FW' in upper_component_2 \
            and upper_component_1 in upper_row and upper_component_2 in upper_row:
        component_list.append(strip_cell(upper_row)[:10:])
    if 'NIC' in upper_component_1 and 'PXE' in upper_component_2 \
            and upper_component_1 in upper_row and upper_component_2 in upper_row:
        component_list.append(strip_cell(upper_row)[:8:])
    if 'NIC' in upper_component_1 and 'UEFI' in upper_component_2 \
            and upper_component_1 in upper_row and upper_component_2 in upper_row:
        component_list.append(strip_cell(upper_row)[:8:])
    if 'NIC' in upper_component_1 and 'DRIVER' in upper_component_2 \
            and upper_component_1 in upper_row and upper_component_2 in upper_row:
        component_list.append(strip_cell(upper_row)[:8:])
    if 'BOOT' in upper_component_1 and 'BOOT' in upper_component_2 \
            and upper_component_1 in upper_row and upper_component_2 in upper_row:
        component_list.append(strip_cell(upper_row)[:8:])
    if 'NVME' in upper_component_1 and 'NVME' in upper_component_2 \
            and upper_component_1 in upper_row and upper_component_2 in upper_row:
        component_list.append(strip_cell(upper_row)[:18:])
    if 'HDD' in upper_component_1 and 'HDD' in upper_component_2 \
            and upper_component_1 in upper_row and upper_component_2 in upper_row:
        component_list.append(strip_cell(upper_row)[:24:])
    if 'PSU' in upper_component_1 and 'PSU' in upper_component_2 \
            and upper_component_1 in upper_row and upper_component_2 in upper_row:
        component_list.append(strip_cell(upper_row)[::])

    count = 0
    while count < len(component_list):
        component_count = component_list[count]
        all_dict.setdefault(lower_component, []).append(component_count)
        count += 1


def get_cover_page():
    try:
        with open('cover_page.csv', 'r') as cover_file:
            for row in cover_file:
                print(row)
                get_components_from_csv('gen', '.', row, gen_list)
    except UnicodeDecodeError:
        pass


def get_configruation_sheet():
    with open('pipe_cleaner/crd_info/fw-sw_configuration.csv', 'r') as configuration_file:
        for row in configuration_file:
            get_components_from_csv('azure', 'gen', row, azure_list)
            get_components_from_csv('blade', 'blade', row, blade_list)
            get_components_from_csv('server', 'server', row, blade_list)
            get_components_from_csv('bios', 'BS.', row, bios_list)
            get_components_from_csv('bmc', 'BC.', row, bmc_list)
            get_components_from_csv('tpm', 'tpm', row, tpm_list)
            get_components_from_csv('cpld', 'cpld', row, cpld_list)
            get_components_from_csv('chipset', 'chipset', row, chipset_driver_list)  # Chipset Driver
            get_components_from_csv('processor', 'processor', row, server_processor_list)  # Processor Driver
            get_components_from_csv('fpga', 'fpga', row, fpga_release_package_list)  # FPGA Package
            get_components_from_csv('hyperblaster', 'dll', row, fpga_hyperblaster_list)  # FPGA Package
            get_components_from_csv('fpga', 'hip', row, fpga_hip_list)  # FPGA Package
            get_components_from_csv('fpga', 'filter', row, fpga_filter_list)  # FPGA Package
            get_components_from_csv('ftdi', 'port', row, ftdi_port_list)  # FPGA Package
            get_components_from_csv('ftdi', 'bus', row, ftdi_bus_list)  # FPGA Package
            get_components_from_csv('nic', 'fw', row, nic_firmware_list)
            get_components_from_csv('nic', 'pxe', row, nic_pxe_list)
            get_components_from_csv('nic', 'uefi', row, nic_uefi_list)
            get_components_from_csv('nic', 'driver', row, nic_driver_list)
            get_components_from_csv('nvme', 'nvme', row, nvme_list)
            get_components_from_csv('hdd', 'hdd', row, hdd_list)
            get_components_from_csv('psu', 'psu', row, psu_list)
            get_components_from_csv('boot', 'boot', row, boot_drive_list)


def get_gen():
    if not gen_list:
        return 'NONE'
    else:
        all_dict['gen'] = gen_list
        return gen_list


def get_azure():
    if not azure_list:
        return 'NONE'
    else:
        all_dict['azure'] = azure_list
        return azure_list


def get_blade():
    if not blade_list:
        return 'NONE'
    else:
        all_dict['blade'] = blade_list
        return blade_list


def get_server():
    if not server_list:
        return 'NONE'
    else:
        all_dict['server'] = server_list
        return server_list


def get_bios():
    if not bios_list:
        return 'NONE'
    else:
        all_dict['bsl bios'] = bios_list
        return bios_list


def get_bmc():
    if not bmc_list:
        return 'NONE'
    else:
        all_dict['bmc'] = bmc_list
        return bmc_list


def get_tpm():
    if not tpm_list:
        return 'NONE'
    else:
        all_dict['tpm'] = tpm_list
        return tpm_list


def get_cpld():
    if not cpld_list:
        return 'NONE'
    else:
        all_dict['cpld'] = cpld_list
        return cpld_list


def get_chipset():
    if not chipset_driver_list:
        return 'NONE'
    else:
        all_dict['chipset_driver'] = chipset_driver_list
        return chipset_driver_list


def get_processor():
    if not server_processor_list:
        return 'NONE'
    else:
        all_dict['processor'] = server_processor_list
        return server_processor_list


def get_fpga_release():
    if not fpga_release_package_list:
        return 'NONE'
    else:
        all_dict['fpga_release'] = fpga_release_package_list
        return fpga_release_package_list


def get_fpga_hyperblaster():
    if not fpga_hyperblaster_list:
        return 'NONE'
    else:
        all_dict['fpga_hyperblaster'] = fpga_hyperblaster_list
        return fpga_hyperblaster_list


def get_fpga_hip():
    if not fpga_hip_list:
        return 'NONE'
    else:
        all_dict['fpga_hip'] = fpga_hip_list
        return fpga_hip_list


def get_fpga_filter():
    if not fpga_filter_list:
        return 'NONE'
    else:
        all_dict['fpga_filter'] = fpga_filter_list
        return fpga_filter_list


def get_ftdi_port():
    if not ftdi_port_list:
        return 'NONE'
    else:
        all_dict['ftdi_port'] = ftdi_port_list
        return ftdi_port_list


def get_ftdi_bus():
    if not ftdi_bus_list:
        return 'NONE'
    else:
        all_dict['ftdi_bus'] = ftdi_bus_list
        return ftdi_bus_list


def get_nic():
    if not nic_firmware_list:
        return 'NONE'
    else:
        all_dict['nic_fw'] = nic_firmware_list
        return nic_firmware_list


def get_nic_pxe():
    if not nic_pxe_list:
        return 'NONE'
    else:
        all_dict['nic_pxe'] = nic_pxe_list
        return nic_pxe_list


def get_nic_uefi():
    if not nic_uefi_list:
        return 'NONE'
    else:
        all_dict['nic_uefi'] = nic_uefi_list
        return nic_uefi_list


def get_boot_drive():
    if not boot_drive_list:
        return 'NONE'
    else:
        all_dict['boot_drive'] = boot_drive_list
        return boot_drive_list


def get_nvme_pn():
    if not nvme_list:
        return 'NONE'
    else:
        all_dict['nvme_pn'] = nvme_list
        return nvme_list


def get_hdd_pn():
    if not hdd_list:
        return 'NONE'
    else:
        all_dict['hdd_pn'] = hdd_list
        return hdd_list


def get_psu():
    if not psu_list:
        return 'NONE'
    else:
        all_dict['psu'] = psu_list
        return psu_list


def get_boot():
    if not boot_drive_list:
        return 'NONE'
    else:
        all_dict['boot'] = boot_drive_list
        return boot_drive_list


def create_csv():
    create_csv_from_crd()
    get_configruation_sheet()
