from pipe_cleaner.src.credentials import Path
from json import loads

# Tally for comparing JSON files
match = []
mismatch = []
missing = []
software_tally = []
hardware_tally = []

new_configuration_names = []

components_list = [
    'Target Type',
    'Target Configuration',
    'Part Number',
    'Supplier',
    'Description',
    'Mixed Required',
    'Datasheet',
    'Toolkit',
    'Firmware',
    'Firmware N-1',
    'Diagnostic Utility',
    'Firmware Update Utility',
    'Reference Specifications',
    'Reference Test Plans',
    'Reference Test Data',
    'Reference Configuration',
    'Known Issue',

    'Component',
    'Server BIOS',
    'Server BMC',
    'Server CPLD',
    'Server OS',
    'Server Chipset Driver',
    'Server Partition Volume',
    'Server Boot Drive',
    'Server Motherboard PN',
    'Server Processors',
    'Server TPM',
    'FPGA Release Package',
    'FPGA Board PN',
    'FPGA Active Image',
    'FPGA Inactive Images',
    'Hyperblaster DLL',
    'FPGA HIP Driver',
    'FPGA Filter Driver',
    'FTDI Port Driver',
    'FTDI Bus Driver',
    'Server NIC Firmware',
    'Server NIC PXE',
    'Server NIC Driver',

    'QCL NVME 1',
    'QCL NVME 2',
    'QCL NVME 3',
    'QCL NVME 4',
    'QCL NVME 5',
    'QCL NVME 6',
    'QCL NVME 7',
    'QCL NVME 8',
    'QCL NVME 9',
    'QCL NVME 10',

    'QCL SSD 1',
    'QCL SSD 2',
    'QCL SSD 3',
    'QCL SSD 4',
    'QCL SSD 5',
    'QCL SSD 6',
    'QCL SSD 7',
    'QCL SSD 8',
    'QCL SSD 9',
    'QCL SSD 10',

    'QCL HDD 1',
    'QCL HDD 2',
    'QCL HDD 3',
    'QCL HDD 4',
    'QCL HDD 5',
    'QCL HDD 6',
    'QCL HDD 7',
    'QCL HDD 8',
    'QCL HDD 9',
    'QCL HDD 10',

    'QCL DIMM 1',
    'QCL DIMM 2',
    'QCL DIMM 3',
    'QCL DIMM 4',
    'QCL DIMM 5',
    'QCL DIMM 6',
    'QCL DIMM 7',
    'QCL DIMM 8',
    'QCL DIMM 9',
    'QCL DIMM 10',

    'Sever PSU Firmware',
    'Rack Manager PN',
    'Manager Switch Firmware',
    'PMDU',
    'Switch PN',

    'Chassis PSU Firmware',
    'Chassis PSU PN',
    'Chassis Manager Service',
    'Chassis Manager OS',
    'Chassis Manager BIOS',
    'Chassis Manager TPM',
    'Chassis Manager PN',
]


def parse_configuration_names(list_position, components_list):
    """
    Parses the first, second, and last terms for each configuration name and all caps them for later parsing.
    Returns a new list of configuration names.
    :return: three all caps terms
    """

    for item in components_list:
        first = str(item).split(' ')[0]
        try:
            second = str(item).split(' ')[1]
        except IndexError:
            second = first
        last = str(item).split(' ')[-1]
        together = f'{first.upper()} {second.upper()} {last.upper()}'
        new_configuration_names.append(together)

        first_upper = first.upper()
        second_upper = second.upper()
        last_upper = last.upper()

        items = []

        items.append(first_upper)
        items.append(second_upper)
        items.append(last_upper)

        return items


def set_white_lines(wb, ws):
    print('\n  - setting up white lines... ')
    letter_one = 'E'
    letter_two = 'F'
    letter_three = 'G'

    structure = Structure(wb)

    def number(num):
        start = initial + num
        return start

    ws.write(f'{letter_one}{number(1)}', 'Status', structure.teal_middle)
    ws.write(f'{letter_two}{number(0)}', f'Test Run Request', structure.teal_middle)
    ws.write(f'{letter_three}{number(0)}', f'Kirkland System', structure.teal_middle)

    ws.write(f'{letter_one}{number(12)}', '', structure.white)
    ws.write(f'{letter_two}{number(12)}', '', structure.white)
    ws.write(f'{letter_three}{number(12)}', '', structure.white)

    ws.write(f'{letter_one}{number(22)}', '', structure.white)
    ws.write(f'{letter_two}{number(22)}', '', structure.white)
    ws.write(f'{letter_three}{number(22)}', '', structure.white)

    ws.write(f'{letter_one}{number(27)}', '', structure.white)
    ws.write(f'{letter_two}{number(27)}', '', structure.white)
    ws.write(f'{letter_three}{number(27)}', '', structure.white)

    ws.write(f'{letter_one}{number(34)}', '', structure.white)
    ws.write(f'{letter_two}{number(34)}', '', structure.white)
    ws.write(f'{letter_three}{number(34)}', '', structure.white)

    ws.write(f'{letter_one}{number(41)}', '', structure.white)
    ws.write(f'{letter_two}{number(41)}', '', structure.white)
    ws.write(f'{letter_three}{number(41)}', '', structure.white)

    ws.write(f'{letter_one}{number(48)}', '', structure.white)
    ws.write(f'{letter_two}{number(48)}', '', structure.white)
    ws.write(f'{letter_three}{number(48)}', '', structure.white)

    ws.write(f'{letter_one}{number(55)}', '', structure.white)
    ws.write(f'{letter_two}{number(55)}', '', structure.white)
    ws.write(f'{letter_three}{number(55)}', '', structure.white)

    ws.write(f'{letter_one}{number(76)}', '', structure.white)
    ws.write(f'{letter_two}{number(76)}', '', structure.white)
    ws.write(f'{letter_three}{number(76)}', '', structure.white)


def set_graphs(wb, ws):
    bold1 = wb.add_format({'bold': 1})

    headings = ['Category', 'Values']
    data = [
        [f'Match: {sum(match)}', f'Mismatch: {sum(mismatch)}', f'Missing Info: {sum(missing)}'],
        [sum(match), sum(mismatch), sum(missing)],
    ]

    ws.write_row('M1', headings, bold1)
    ws.write_column('M2', data[0])
    ws.write_column('N2', data[1])

    chart2 = wb.add_chart({'type': 'pie'})

    chart2.add_series({
        'name': 'Future Pie Chart',
        'categories': '=Sheet1!$M$2:$M$4',
        'values': '=Sheet1!$N$2:$N$4',
        'points': [
            {'fill': {'color': '#00B050'}},
            {'fill': {'color': '#FF0000'}},
            {'fill': {'color': '#F5BD1F'}},
        ],
        'data_labels': {'percentage': True},
    })

    chart2.set_chartarea({'border': {'none': True}})

    chart2.set_title({'name': 'Future Pie Chart'})

    chart2.set_style(10)

    ws.insert_chart('G2', chart2, {'x_offset': 12, 'y_offset': 5})

    chart1 = wb.add_chart({'type': 'bar', 'subtype': 'percent_stacked'})

    chart1.add_series({
        'title_color: blue'
        'name': 'Future Bar Graph',
        'categories': '=Sheet1!$M$2:$M$4',
        'values': '=Sheet1!$N$2:$N$4',
        'points': [
            {'fill': {'color': '#00B050'}},
            {'fill': {'color': '#FF0000'}},
            {'fill': {'color': '#F5BD1F'}},
        ],
        'data_labels': {'percentage': True},
    })

    chart1.add_series({
        'name': '=Sheet1!$C$1',
        'categories': '=Sheet1!$A$2:$A$7',
        'values': '=Sheet1!$C$2:$C$7',
    })

    chart1.set_title({'name': 'Future Bar Chart'})
    chart1.set_x_axis({'name': 'Test number'})
    chart1.set_y_axis({'name': 'Sample length (mm)'})

    chart1.set_style(13)

    chart1.set_chartarea({'border': {'none': True}})

    ws.insert_chart('F2', chart1, {'x_offset': 25, 'y_offset': 10})

def fill_trr_column(wb, ws, trr):
    structure = Structure(wb)

    def start(num):
        start = initial + num
        return start

    compare_sources(wb, ws, trr, start(1), 'COMPONENT', 'COMPONENT', 'COMPONENT', structure.blue_middle)
    compare_sources(wb, ws, trr, start(2), 'SERVER', 'BI', 'OS', structure.blue_middle)
    compare_sources(wb, ws, trr, start(3), 'SERVER', 'BMC', 'BMC', structure.blue_middle)
    compare_sources(wb, ws, trr, start(4), 'SERVER', 'CPLD', 'CPLD', structure.blue_middle)
    compare_sources(wb, ws, trr, start(5), 'SERVER', 'SERVER', 'OS', structure.blue_middle)
    compare_sources(wb, ws, trr, start(6), 'CHIPSET', 'DRIVER', 'DRIVER', structure.blue_middle)
    compare_sources(wb, ws, trr, start(7), 'PARTITION', 'VOLUME', 'VOLUME', structure.blue_middle)
    compare_sources(wb, ws, trr, start(8), 'BOOT', 'DRIVE', 'DRIVE', structure.blue_middle)
    compare_sources(wb, ws, trr, start(9), 'MOTHERBOARD', 'PN', 'PN', structure.blue_middle)
    compare_sources(wb, ws, trr, start(10), 'PROCESSORS', 'PROCESSORS', 'PROCESSORS', structure.blue_middle)
    compare_sources(wb, ws, trr, start(11), 'SERVER', 'TPM', 'TPM', structure.blue_middle)
    # line break at 20
    compare_sources(wb, ws, trr, start(13), 'FPGA', 'RELEASE', 'PACKAGE', structure.blue_middle)
    compare_sources(wb, ws, trr, start(14), 'FPGA', 'BOARD', 'PN', structure.blue_middle)
    compare_sources(wb, ws, trr, start(15), 'FPGA', 'ACTIVE', 'IMAGE', structure.blue_middle)
    compare_sources(wb, ws, trr, start(16), 'FPGA', 'INACTIVE', 'IMAGE', structure.blue_middle)
    compare_sources(wb, ws, trr, start(17), 'HYPERBLASTER', 'HYPERBLASTER', 'HYPERBLASTER', structure.blue_middle)
    compare_sources(wb, ws, trr, start(18), 'HIP', 'DRIVER', 'DRIVER', structure.blue_middle)
    compare_sources(wb, ws, trr, start(19), 'FILTER', 'DRIVER', 'DRIVER', structure.blue_middle)
    compare_sources(wb, ws, trr, start(20), 'PORT', 'DRIVER', 'DRIVER', structure.blue_middle)
    compare_sources(wb, ws, trr, start(21), 'BUS', 'DRIVER', 'DRIVER', structure.blue_middle)
    # line break at 30
    compare_sources(wb, ws, trr, start(23), 'SERVER', 'NIC', 'FIRMWARE', structure.blue_middle)
    compare_sources(wb, ws, trr, start(24), 'SERVER', 'NIC', 'PXE', structure.blue_middle)
    compare_sources(wb, ws, trr, start(25), 'SERVER', 'NIC', 'UEFI', structure.blue_middle)
    compare_sources(wb, ws, trr, start(26), 'SERVER', 'NIC', 'DRIVER', structure.blue_middle)
    # line break at 35
    compare_sources(wb, ws, trr, start(28), 'QCL', 'NVME', '1', structure.blue_middle)
    compare_sources(wb, ws, trr, start(29), 'QCL', 'NVME', '2', structure.blue_middle)
    compare_sources(wb, ws, trr, start(30), 'QCL', 'NVME', '3', structure.blue_middle)
    compare_sources(wb, ws, trr, start(31), 'QCL', 'NVME', '4', structure.blue_middle)
    compare_sources(wb, ws, trr, start(32), 'QCL', 'NVME', '5', structure.blue_middle)
    compare_sources(wb, ws, trr, start(33), 'QCL', 'NVME', '6', structure.blue_middle)
    # line break at 42
    compare_sources(wb, ws, trr, start(35), 'QCL', 'SSD', '1', structure.blue_middle)
    compare_sources(wb, ws, trr, start(36), 'QCL', 'SSD', '2', structure.blue_middle)
    compare_sources(wb, ws, trr, start(37), 'QCL', 'SSD', '3', structure.blue_middle)
    compare_sources(wb, ws, trr, start(38), 'QCL', 'SSD', '4', structure.blue_middle)
    compare_sources(wb, ws, trr, start(39), 'QCL', 'SSD', '5', structure.blue_middle)
    compare_sources(wb, ws, trr, start(40), 'QCL', 'SSD', '6', structure.blue_middle)
    # line break at 49
    compare_sources(wb, ws, trr, start(42), 'QCL', 'HDD', '1', structure.blue_middle)
    compare_sources(wb, ws, trr, start(43), 'QCL', 'HDD', '2', structure.blue_middle)
    compare_sources(wb, ws, trr, start(44), 'QCL', 'HDD', '3', structure.blue_middle)
    compare_sources(wb, ws, trr, start(45), 'QCL', 'HDD', '4', structure.blue_middle)
    compare_sources(wb, ws, trr, start(46), 'QCL', 'HDD', '5', structure.blue_middle)
    compare_sources(wb, ws, trr, start(47), 'QCL', 'HDD', '6', structure.blue_middle)
    # line break at 56
    compare_sources(wb, ws, trr, start(49), 'QCL', 'DIMM', '1', structure.blue_middle)
    compare_sources(wb, ws, trr, start(50), 'QCL', 'DIMM', '2', structure.blue_middle)
    compare_sources(wb, ws, trr, start(51), 'QCL', 'DIMM', '3', structure.blue_middle)
    compare_sources(wb, ws, trr, start(52), 'QCL', 'DIMM', '4', structure.blue_middle)
    compare_sources(wb, ws, trr, start(53), 'QCL', 'DIMM', '5', structure.blue_middle)
    compare_sources(wb, ws, trr, start(54), 'QCL', 'DIMM', '6', structure.blue_middle)
    # line break at 63
    compare_sources(wb, ws, trr, start(56), 'CHASSIS', 'PSU', 'FIRMWARE', structure.blue_middle)
    compare_sources(wb, ws, trr, start(57), 'CHASSIS', 'PSU', 'PN', structure.blue_middle)
    compare_sources(wb, ws, trr, start(58), 'RACK', 'MANAGER', 'FIRMWARE', structure.blue_middle)
    compare_sources(wb, ws, trr, start(59), 'RACK', 'MANAGER', 'PN', structure.blue_middle)
    compare_sources(wb, ws, trr, start(60), 'MANAGER', 'SWTICH', 'FIRMWARE', structure.blue_middle)
    compare_sources(wb, ws, trr, start(61), 'PMDU', 'PMDU', 'PMDU', structure.blue_middle)
    compare_sources(wb, ws, trr, start(62), 'SWITCH', 'PN', 'PN', structure.blue_middle)
    compare_sources(wb, ws, trr, start(63), 'REQUEST', 'TYPE', 'TYPE', structure.blue_middle)
    compare_sources(wb, ws, trr, start(64), 'TARGET', 'TARGET', 'CONFIGURATION', structure.blue_middle)
    compare_sources(wb, ws, trr, start(65), 'PART', 'NUMBER', 'NUMBER', structure.blue_middle)
    compare_sources(wb, ws, trr, start(66), 'SUPPLIER', 'SUPPLIER', 'SUPPLIER', structure.blue_middle)
    compare_sources(wb, ws, trr, start(67), 'DESCRIPTION', 'DESCRIPTION', 'DESCRIPTION', structure.blue_middle)
    compare_sources(wb, ws, trr, start(68), 'MIXED', 'REQUIRED', 'REQUIRED', structure.blue_middle)
    compare_sources(wb, ws, trr, start(69), 'DATA', 'SHEET', 'SHEET', structure.blue_middle)
    compare_sources(wb, ws, trr, start(70), 'TOOLKIT', 'TOOLKIT', 'TOOLKIT', structure.blue_middle)
    compare_sources(wb, ws, trr, start(71), 'REFERENCE', 'SPECIFICATIONS', 'SPECIFICATIONS', structure.blue_middle)
    compare_sources(wb, ws, trr, start(72), 'REFERENCE', 'TEST', 'PLANS', structure.blue_middle)
    compare_sources(wb, ws, trr, start(73), 'REFERENCE', 'TEST', 'DATA', structure.blue_middle)
    compare_sources(wb, ws, trr, start(74), 'REFERENCE', 'CONFIGURATION', 'CONFIGURATION', structure.blue_middle)
    compare_sources(wb, ws, trr, start(75), 'KNOWN', 'ISSUE', 'ISSUE', structure.blue_middle)

def fill_host_column(wb, ws, host):
    letter = 'G'
    structure = Structure(wb)
    ws.write(f'{letter}9', 'Console Server System', structure.teal_middle)

    def number(num):
        start = initial + num
        return start

    with open(f'{Path.info}{str(host)}.json') as f:
        system = loads(f.read())

    # compare_sources(wb, ws, trr, start(1), 'COMPONENT', 'COMPONENT', 'COMPONENT', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(2), 'SERVER', 'BI', 'OS', structure.blue_middle)

    bios = str(system['dmi']['bios']['version'])
    ws.write(f'{letter}{number(2)}', f'{bios}', structure.blue_middle)
    bmc = str(system['bmc']['mc']['firmware'])
    ws.write(f'{letter}{number(3)}', f'{bmc}', structure.blue_middle)

    os = str(system['platform']['version'])
    ws.write(f'{letter}{number(5)}', f'{os}', structure.blue_middle)

    tpm = str(system['tpm']['version'])
    ws.write(f'{letter}{number(11)}', f'{tpm}', structure.blue_middle)

    # compare_sources(wb, ws, trr, start(3), 'SERVER', 'BMC', 'BMC', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(4), 'SERVER', 'CPLD', 'CPLD', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(5), 'SERVER', 'SERVER', 'OS', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(6), 'CHIPSET', 'DRIVER', 'DRIVER', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(7), 'PARTITION', 'VOLUME', 'VOLUME', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(8), 'BOOT', 'DRIVE', 'DRIVE', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(9), 'MOTHERBOARD', 'PN', 'PN', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(10), 'PROCESSORS', 'PROCESSORS', 'PROCESSORS', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(11), 'SERVER', 'TPM', 'TPM', structure.blue_middle)
    # # line break at 20
    # compare_sources(wb, ws, trr, start(13), 'FPGA', 'RELEASE', 'PACKAGE', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(14), 'FPGA', 'BOARD', 'PN', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(15), 'FPGA', 'ACTIVE', 'IMAGE', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(16), 'FPGA', 'INACTIVE', 'IMAGE', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(17), 'HYPERBLASTER', 'HYPERBLASTER', 'HYPERBLASTER', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(18), 'HIP', 'DRIVER', 'DRIVER', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(19), 'FILTER', 'DRIVER', 'DRIVER', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(20), 'PORT', 'DRIVER', 'DRIVER', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(21), 'BUS', 'DRIVER', 'DRIVER', structure.blue_middle)
    # # line break at 30
    # compare_sources(wb, ws, trr, start(23), 'SERVER', 'NIC', 'FIRMWARE', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(24), 'SERVER', 'NIC', 'PXE', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(25), 'SERVER', 'NIC', 'UEFI', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(26), 'SERVER', 'NIC', 'DRIVER', structure.blue_middle)
    # # line break at 35
    # compare_sources(wb, ws, trr, start(28), 'QCL', 'NVME', '1', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(29), 'QCL', 'NVME', '2', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(30), 'QCL', 'NVME', '3', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(31), 'QCL', 'NVME', '4', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(32), 'QCL', 'NVME', '5', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(33), 'QCL', 'NVME', '6', structure.blue_middle)
    # # line break at 42
    # compare_sources(wb, ws, trr, start(35), 'QCL', 'SSD', '1', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(36), 'QCL', 'SSD', '2', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(37), 'QCL', 'SSD', '3', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(38), 'QCL', 'SSD', '4', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(39), 'QCL', 'SSD', '5', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(40), 'QCL', 'SSD', '6', structure.blue_middle)
    # # line break at 49
    # compare_sources(wb, ws, trr, start(42), 'QCL', 'HDD', '1', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(43), 'QCL', 'HDD', '2', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(44), 'QCL', 'HDD', '3', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(45), 'QCL', 'HDD', '4', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(46), 'QCL', 'HDD', '5', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(47), 'QCL', 'HDD', '6', structure.blue_middle)
    # # line break at 56
    # compare_sources(wb, ws, trr, start(49), 'QCL', 'DIMM', '1', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(50), 'QCL', 'DIMM', '2', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(51), 'QCL', 'DIMM', '3', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(52), 'QCL', 'DIMM', '4', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(53), 'QCL', 'DIMM', '5', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(54), 'QCL', 'DIMM', '6', structure.blue_middle)
    # # line break at 63
    # compare_sources(wb, ws, trr, start(56), 'CHASSIS', 'PSU', 'FIRMWARE', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(57), 'CHASSIS', 'PSU', 'PN', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(58), 'RACK', 'MANAGER', 'FIRMWARE', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(59), 'RACK', 'MANAGER', 'PN', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(60), 'MANAGER', 'SWTICH', 'FIRMWARE', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(61), 'PMDU', 'PMDU', 'PMDU', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(62), 'SWITCH', 'PN', 'PN', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(63), 'REQUEST', 'TYPE', 'TYPE', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(64), 'TARGET', 'TARGET', 'CONFIGURATION', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(65), 'PART', 'NUMBER', 'NUMBER', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(66), 'SUPPLIER', 'SUPPLIER', 'SUPPLIER', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(67), 'DESCRIPTION', 'DESCRIPTION', 'DESCRIPTION', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(68), 'MIXED', 'REQUIRED', 'REQUIRED', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(69), 'DATA', 'SHEET', 'SHEET', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(70), 'TOOLKIT', 'TOOLKIT', 'TOOLKIT', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(71), 'REFERENCE', 'SPECIFICATIONS', 'SPECIFICATIONS', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(72), 'REFERENCE', 'TEST', 'PLANS', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(73), 'REFERENCE', 'TEST', 'DATA', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(74), 'REFERENCE', 'CONFIGURATION', 'CONFIGURATION', structure.blue_middle)
    # compare_sources(wb, ws, trr, start(75), 'KNOWN', 'ISSUE', 'ISSUE', structure.blue_middle)