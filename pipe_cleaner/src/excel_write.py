from pipe_cleaner.src.credentials import Path
from time import strftime
from json import loads
import xlsxwriter

# Adjustment for Excel Structure
initial = 8

# Tally for comparing JSON files
match = []
mismatch = []
missing = []
software_tally = []
hardware_tally = []


def set_lines(pipe_num, wb, ws, host):

    structure = Structure(wb)

    with open(f'{Path.info}{str(host)}.json') as f:
        system = loads(f.read())

    location = system['location']

    ws.set_row(0, 12, structure.white)
    ws.set_row(1, 80, structure.white)
    ws.set_row(2, 25, structure.white)
    ws.set_row(3, 25, structure.white)
    ws.set_row(4, 25, structure.white)
    ws.set_row(5, 25, structure.white)
    ws.set_row(6, 25, structure.white)
    ws.set_row(7, 25, structure.white)
    ws.set_row(19, 10)
    ws.set_row(29, 10)
    ws.set_row(34, 10)
    ws.set_row(41, 10)
    ws.set_row(48, 10)
    ws.set_row(55, 10)
    ws.set_row(62, 10)
    ws.set_row(83, 30, structure.white)
    ws.set_row(84, 30, structure.white)
    ws.set_row(85, 30, structure.white)
    ws.set_row(86, 30, structure.white)
    ws.set_row(87, 30, structure.white)
    ws.set_row(88, 30, structure.white)
    ws.set_row(89, 30, structure.white)
    ws.set_row(90, 30, structure.white)
    ws.set_row(91, 30, structure.white)
    ws.set_row(92, 30, structure.white)
    ws.set_row(93, 30, structure.white)
    ws.set_row(94, 30, structure.white)
    ws.set_row(95, 30, structure.white)
    ws.set_row(96, 30, structure.white)
    ws.set_row(97, 30, structure.white)
    ws.set_row(98, 30, structure.white)
    ws.set_row(99, 30, structure.white)
    ws.set_row(100, 30, structure.white)
    ws.set_column('A:A', 1, structure.white)
    ws.set_column('B:B', 1, structure.white)
    ws.set_column('C:C', 1, structure.white)
    ws.set_column('D:D', 32)
    ws.set_column('E:E', 32)
    ws.set_column('F:F', 74)
    ws.set_column('G:G', 74)
    ws.set_column('H:H', 25, structure.white)
    ws.set_column('I:I', 25, structure.white)
    ws.set_column('J:J', 25, structure.white)
    ws.set_column('K:K', 25, structure.white)
    ws.set_column('L:L', 25, structure.white)

    ws.write('D3', 'VSE Pipe Cleaner', structure.blue_font)
    ws.write('D4', 'Kirkland Lab Site', structure.blue_font)
    ws.write('D5', structure.date_time, structure.blue_font)
    ws.write('D6', f'Pipe: {pipe_num}', structure.blue_font)
    ws.write('D6', f'Location: {location}', structure.blue_font)


def set_component_column(wb, ws):
    letter = 'D'

    def number(num):
        start = initial + num
        return start

    structure = Structure(wb)

    ws.insert_image('B2', 'pipe_cleaner/img/vse_logo.png')

    ws.write(f'{letter}{number(1)}', 'Component', structure.teal_left)
    ws.write(f'{letter}{number(2)}', 'Server BIOS', structure.blue_component)
    ws.write(f'{letter}{number(3)}', 'Server BMC', structure.blue_component)
    ws.write(f'{letter}{number(4)}', 'Server CPLD', structure.blue_component)
    ws.write(f'{letter}{number(5)}', 'Server OS', structure.blue_component)
    ws.write(f'{letter}{number(6)}', 'Server Chipset Driver', structure.blue_component)
    ws.write(f'{letter}{number(7)}', 'Server Partition - Volume', structure.blue_component)
    ws.write(f'{letter}{number(8)}', 'Server Boot Drive', structure.blue_component)
    ws.write(f'{letter}{number(9)}', 'Server Motherboard PN#', structure.blue_component)
    ws.write(f'{letter}{number(10)}', 'Server Processors', structure.blue_component)
    ws.write(f'{letter}{number(11)}', 'Server TPM', structure.blue_component)
    ws.write(f'{letter}{number(12)}', '', structure.white)

    ws.write(f'{letter}{number(13)}', 'Server FPGA Release Package', structure.blue_component)
    ws.write(f'{letter}{number(14)}', 'Server FPGA Board PN#', structure.blue_component)
    ws.write(f'{letter}{number(15)}', 'Server FPGA Active Image', structure.blue_component)
    ws.write(f'{letter}{number(16)}', 'Server FPGA Inactive Images', structure.blue_component)
    ws.write(f'{letter}{number(17)}', 'Server Hyperblaster DLL', structure.blue_component)
    ws.write(f'{letter}{number(18)}', 'Server FPGA HIP Driver', structure.blue_component)
    ws.write(f'{letter}{number(19)}', 'Server FPGA Filter Driver', structure.blue_component)
    ws.write(f'{letter}{number(20)}', 'Server FTDI Port Driver', structure.blue_component)
    ws.write(f'{letter}{number(21)}', 'Server FTDI Bus Driver', structure.blue_component)
    ws.write(f'{letter}{number(22)}', '', structure.white)

    ws.write(f'{letter}{number(23)}', 'Server NIC Firmware', structure.blue_component)
    ws.write(f'{letter}{number(24)}', 'Server NIC PXE', structure.blue_component)
    ws.write(f'{letter}{number(25)}', 'Server NIC UEFI', structure.blue_component)
    ws.write(f'{letter}{number(26)}', 'Server NIC Driver', structure.blue_component)
    ws.write(f'{letter}{number(27)}', '', structure.white)

    ws.write(f'{letter}{number(28)}', 'QCL - NVME #1', structure.blue_component)
    ws.write(f'{letter}{number(29)}', 'QCL - NVME #2', structure.blue_component)
    ws.write(f'{letter}{number(30)}', 'QCL - NVME #3', structure.blue_component)
    ws.write(f'{letter}{number(31)}', 'QCL - NVME #4', structure.blue_component)
    ws.write(f'{letter}{number(32)}', 'QCL - NVME #5', structure.blue_component)
    ws.write(f'{letter}{number(33)}', 'QCL - NVME #6', structure.blue_component)
    ws.write(f'{letter}{number(34)}', '', structure.white)

    ws.write(f'{letter}{number(35)}', 'QCL - SATA SSD #1', structure.blue_component)
    ws.write(f'{letter}{number(36)}', 'QCL - SATA SSD #2', structure.blue_component)
    ws.write(f'{letter}{number(37)}', 'QCL - SATA SSD #3', structure.blue_component)
    ws.write(f'{letter}{number(38)}', 'QCL - SATA SSD #4', structure.blue_component)
    ws.write(f'{letter}{number(39)}', 'QCL - SATA SSD #5', structure.blue_component)
    ws.write(f'{letter}{number(40)}', 'QCL - SATA SSD #6', structure.blue_component)
    ws.write(f'{letter}{number(41)}', '', structure.white)

    ws.write(f'{letter}{number(42)}', 'QCL - SATA HDD #1', structure.blue_component)
    ws.write(f'{letter}{number(43)}', 'QCL - SATA HDD #2', structure.blue_component)
    ws.write(f'{letter}{number(44)}', 'QCL - SATA HDD #3', structure.blue_component)
    ws.write(f'{letter}{number(45)}', 'QCL - SATA HDD #4', structure.blue_component)
    ws.write(f'{letter}{number(46)}', 'QCL - SATA HDD #5', structure.blue_component)
    ws.write(f'{letter}{number(47)}', 'QCL - SATA HDD #6', structure.blue_component)
    ws.write(f'{letter}{number(48)}', '', structure.white)

    ws.write(f'{letter}{number(49)}', 'QCL - DIMM #1', structure.blue_component)
    ws.write(f'{letter}{number(50)}', 'QCL - DIMM #2', structure.blue_component)
    ws.write(f'{letter}{number(51)}', 'QCL - DIMM #3', structure.blue_component)
    ws.write(f'{letter}{number(52)}', 'QCL - DIMM #4', structure.blue_component)
    ws.write(f'{letter}{number(53)}', 'QCL - DIMM #5', structure.blue_component)
    ws.write(f'{letter}{number(54)}', 'QCL - DIMM #6', structure.blue_component)
    ws.write(f'{letter}{number(55)}', '', structure.white)

    ws.write(f'{letter}{number(56)}', 'Chassis PSU Firmware', structure.blue_component)
    ws.write(f'{letter}{number(57)}', 'Chassis PSU PN#', structure.blue_component)
    ws.write(f'{letter}{number(58)}', 'Rack Manager Firmware', structure.blue_component)
    ws.write(f'{letter}{number(59)}', 'Rack Manager PN#', structure.blue_component)
    ws.write(f'{letter}{number(60)}', 'Manager Switch Firmware', structure.blue_component)
    ws.write(f'{letter}{number(61)}', 'PMDU', structure.blue_component)
    ws.write(f'{letter}{number(62)}', 'Switch PN#', structure.blue_component)
    ws.write(f'{letter}{number(63)}', 'Request Type', structure.blue_component)
    ws.write(f'{letter}{number(64)}', 'Target Configuration', structure.blue_component)
    ws.write(f'{letter}{number(65)}', 'Part Number', structure.blue_component)
    ws.write(f'{letter}{number(66)}', 'Supplier', structure.blue_component)
    ws.write(f'{letter}{number(67)}', 'Description', structure.blue_component)
    ws.write(f'{letter}{number(68)}', 'Mixed Required', structure.blue_component)
    ws.write(f'{letter}{number(69)}', 'Datasheet', structure.blue_component)
    ws.write(f'{letter}{number(70)}', 'Toolkit', structure.blue_component)
    ws.write(f'{letter}{number(71)}', 'Reference Specifications', structure.blue_component)
    ws.write(f'{letter}{number(72)}', 'Reference Test Plans', structure.blue_component)
    ws.write(f'{letter}{number(73)}', 'Reference Test Data', structure.blue_component)
    ws.write(f'{letter}{number(74)}', 'Reference Configuration', structure.blue_component)
    ws.write(f'{letter}{number(75)}', 'Known Issue', structure.blue_component)


def set_white_lines(wb, ws):
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


def compare_sources(wb, ws, trr_id, num, term1, term2, term3, cell_type):

    trr_string = ''
    sku_string = ''

    structure = Structure(wb)
    test_run_request = TestRunRequest(trr_id)

    for item in test_run_request.ado:
        item_upper = str(item).upper()
        if f'{str(term1)}' in item_upper and f'{str(term2)}' in item_upper and f'{str(term3)}' in item_upper:
            cell = test_run_request.ado[str(item)]
            trr_string = cell
            ws.write(f'F{num}', f'{cell}', cell_type)
    # for item in Storage.host:
    #     item_upper = str(item).upper()
    #     if f'{str(term1)}' in item_upper and f'{str(term2)}' in item_upper and f'{str(term3)}' in item_upper:
    #         cell = Storage.host[str(item)]
    #         sku_string = cell
    #         ws.write(f'G{num}', f'{cell}', cell_type)
    if trr_string in sku_string or sku_string in trr_string:
        ws.write(f'E{num}', 'MATCH', structure.good_cell)
        match.append(1)
        for item in Items.software:
            if trr_string in item:
                software_tally.append(1)
        for item in Items.hardware:
            if sku_string in item:
                hardware_tally.append(1)
    else:
        ws.write(f'E{num}', 'MISMATCH', structure.bad_cell)
        mismatch.append(1)
    if not trr_string or 'empty' in trr_string:
        ws.write(f'E{num}', 'MISSING INFO', structure.neutral_cell)
        ws.write(f'F{num}', '', structure.missing_cell)
        missing.append(1)
    if not sku_string or 'empty' in sku_string:
        ws.write(f'E{num}', 'MISSING INFO', structure.neutral_cell)
        ws.write(f'G{num}', '', structure.missing_cell)
        missing.append(1)


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


class ExcelPath:
    def __init__(self, num, host):
        self.pipe_num = num
        self.host = host

        with open(f'{Path.info}{str(host)}.json') as f:
            self.system = loads(f.read())

        self.name = self.system['machine_name']
        self.short_name = self.name[-3::]

        self.wb = xlsxwriter.Workbook(f'{self.excel_path}')
        self.ws = self.wb.add_worksheet()

class Structure:
    time = strftime('%I:%M %p')
    date = strftime('%m/%d/%Y')
    date_time = strftime('%m/%d/%Y - %I:%M %p')

    def __init__(self, wb):
        self.wb = wb

        self.white = wb.add_format({'border': 2})
        self.white.set_border_color('white')

        self.blue_component = wb.add_format({'border': 2})
        self.blue_component.set_bg_color('#1f497d')
        self.blue_component.set_border_color('white')

        self.blue_component.set_bold()
        self.blue_component.set_font_color('white')

        self.blue_middle = wb.add_format({'border': 2})
        self.blue_middle.set_bg_color('#1f497d')
        self.blue_middle.set_align('center')
        self.blue_middle.set_border_color('white')
        self.blue_middle.set_bold()
        self.blue_middle.set_font_color('white')

        self.teal_left = wb.add_format({'border': 2})
        self.teal_left.set_bg_color('#00B0F0')
        self.teal_left.set_align('left')
        self.teal_left.set_border_color('white')
        self.teal_left.set_bold()
        self.teal_left.set_font_color('white')
        self.teal_left.set_font_size('16')

        self.teal_middle = wb.add_format({'border': 2})
        self.teal_middle.set_bg_color('#00B0F0')
        self.teal_middle.set_align('center')
        self.teal_middle.set_border_color('white')
        self.teal_middle.set_bold()
        self.teal_middle.set_font_color('white')
        self.teal_middle.set_font_size('16')

        self.grey_area = wb.add_format()
        self.grey_area.set_bg_color('gray')

        self.bold = wb.add_format({'bold': True})
        self.bold.set_font_size('13')

        self.bold_middle = wb.add_format()
        self.bold_middle.set_bold()
        self.bold_middle.set_align('center')
        self.bold_middle.set_font('13')

        self.blue_font = wb.add_format({'border': 2})
        self.blue_font.set_bold()
        self.blue_font.set_font_size('14')
        self.blue_font.set_font_color('#1f497d')
        self.blue_font.set_border_color('white')

        self.middle = wb.add_format()
        self.middle.set_align('center')

        self.good_cell = wb.add_format({'border': 2})
        self.good_cell.set_align('center')
        self.good_cell.set_bold()
        self.good_cell.set_bg_color('#00B050')
        self.good_cell.set_font_color('white')
        self.good_cell.set_border_color('white')

        self.bad_cell = wb.add_format({'border': 2})
        self.bad_cell.set_align('center')
        self.bad_cell.set_bold()
        self.bad_cell.set_bg_color('FF0000')
        self.bad_cell.set_font_color('white')
        self.bad_cell.set_border_color('white')

        self.neutral_cell = wb.add_format({'border': 2})
        self.neutral_cell.set_align('center')
        self.neutral_cell.set_bold()
        self.neutral_cell.set_bg_color('F5BD1F')
        self.neutral_cell.set_font_color('white')
        self.neutral_cell.set_border_color('white')

        self.missing_cell = wb.add_format({'border': 2})
        self.missing_cell.set_align('center')
        self.missing_cell.set_bold()
        # missing_cell.set_bg_color('#1f497d')
        self.missing_cell.set_fg_color('white')
        self.missing_cell.set_bg_color('#1f497d')
        self.missing_cell.set_font_color('#1f497d')
        self.missing_cell.set_border_color('white')
        self.missing_cell.set_pattern(7)


class TestRunRequest:
    def __init__(self, trr_id):
        self.trr_id = trr_id
        with open(f'{Path.info}{str(self.trr_id)}/final.json') as f:
            self.ado = loads(f.read())


class Items:
    software = ['Server BIOS',
                'Server BMC',
                'Server CPLD',
                'Server OS',
                'Server Chipset Driver',
                'Server Boot Drive',
                'Server Motherboard PN#',
                'Server Processors',
                'Server TPM',
                'Server FPGA Release Package',
                'Server FPGA Board PN#',
                'Server FPGA Active Image',
                'Server FPGA Inactive Images',
                'Server Hyperblaster DLL',
                'Server FPGA HIP Driver',
                'Server FTDI Port Driver',
                'Server FTDI Bus Driver',
                'Server NIC Firmware',
                'Server NIC PXE',
                'Server NIC UEFI',
                'Server NIC Driver']

    hardware = ['NVMe',
                'DIMM',
                'SSD',
                'HDD']


def main(num, trr_id, host_id):
    # setup excel path
    excel = ExcelPath(num, host_id)
    pipe_num = excel.pipe_num
    workbook = excel.wb
    worksheet = excel.ws
    
    # create excel structure
    set_lines(pipe_num, workbook, worksheet, host_id)
    set_component_column(workbook, worksheet)
    set_white_lines(workbook, worksheet)
    # set_graphs(workbook, worksheet)
    
    # two sources
    fill_trr_column(workbook, worksheet, trr_id)
    fill_host_column(workbook, worksheet, host_id)

    workbook.close()

    return ExcelPath