from pipe_cleaner.src.excel_properties import requested_configuration
from pipe_cleaner.src.excel_properties import Structure
from pipe_cleaner.src.credentials import Path
from pipe_cleaner.src.data_access import request_ado
import pipe_cleaner.extra.crd_scanner as crd_scanner
from json import loads

crd_bios = []
crd_bmc = []
crd_tpm = []
crd_cpld = []

match_tally = []
mismatch_tally = []
missing_tally = []

mismatch_summary = []
missing_summary = []

mismatch_microsoft = []


def create_graphs(wb, ws):
    bold = wb.add_format({'bold': 1})

    # Add the worksheet data that the charts will refer to.
    headings = ['Number', 'Tallies']
    data = [
        ['Match/Present', 'Mismatch', 'Missing'],
        [sum(match_tally), sum(mismatch_tally), sum(missing_tally)],
    ]

    ws.write_row('A1', headings, bold)
    ws.write_column('A2', data[0])
    ws.write_column('B2', data[1])

    # Create Side Bar
    chart1 = wb.add_chart({'type': 'bar'})
    # chart1 = wb.add_chart({'type': 'pie'})

    wb.define_name("CRD vs. TRR", '=Sheet2')
    # Configure the first series.
    chart1.add_series({
        'name':       '=CRD vs. TRR!$B$1',
        'categories': '=CRD vs. TRR!$A$2:$A$4',
        'values':     '=CRD vs. TRR!$B$2:$B$4',
        'points': [
            {'fill': {'color': '#00B050'}},
            {'fill': {'color': '#FF0000'}},
            {'fill': {'color': '#DCAA1B'}},
        ],
    })

    # Configure a second series. Note use of alternative syntax to define ranges.
    chart1.add_series({
        'name':       ['CRD vs. TRR', 0, 2],
        'categories': ['CRD vs. TRR', 1, 0, 3, 0],
        'values':     ['CRD vs. TRR', 1, 2, 3, 2],
    })

    # Add a chart title and some axis labels.
    chart1.set_title({'name': 'Graph of CRD   vs   TRR'})
    chart1.set_x_axis({'name': 'Tally of Status'})
    chart1.set_y_axis({'name': 'Status'})

    # Chart Style of Graph
    chart1.set_style(11)
    # chart1.set_style(10)
    chart1.set_legend({'none': True})

    # Size of Chart
    ws.insert_chart('E1', chart1, {'x_scale': 1.185, 'y_scale': 0.84})


def set_sheet_structure(pipe_number, host_ids, structure, ws, sheet_title, crd, unique_trrs):

    ws.set_landscape()

    ws.set_row(0, 12, structure.white)
    ws.set_row(1, 20, structure.white)
    ws.set_row(2, 16, structure.white)
    ws.set_row(3, 15, structure.white)
    ws.set_row(4, 15, structure.white)
    ws.set_row(5, 15, structure.white)
    ws.set_row(6, 15, structure.white)
    ws.set_row(7, 15, structure.white)
    ws.set_row(8, 15, structure.white)
    ws.set_row(9, 15, structure.white)
    ws.set_row(10, 15, structure.white)
    ws.set_row(11, 15, structure.white)

    start = 13
    while start < 500:
        ws.set_row(start, 16.5, structure.white)
        start +=1

    # ws.set_row(74, 50, structure.white)
    # ws.set_row(84, 30, structure.white)
    # ws.set_row(85, 30, structure.white)
    # ws.set_row(86, 30, structure.white)
    # ws.set_row(87, 30, structure.white)
    # ws.set_row(88, 30, structure.white)
    # ws.set_row(89, 30, structure.white)
    # ws.set_row(90, 30, structure.white)
    # ws.set_row(91, 30, structure.white)
    # ws.set_row(92, 30, structure.white)
    # ws.set_row(93, 30, structure.white)
    # ws.set_row(94, 30, structure.white)
    # ws.set_row(95, 30, structure.white)
    # ws.set_row(96, 30, structure.white)
    # ws.set_row(97, 30, structure.white)
    # ws.set_row(98, 30, structure.white)
    # ws.set_row(99, 30, structure.white)
    # ws.set_row(100, 30, structure.white)

    ws.set_column('A:A', 5.5, structure.white)
    ws.set_column('B:B', 13, structure.white)
    ws.set_column('C:C', 13, structure.white)
    ws.set_column('D:D', 24, structure.white)
    ws.set_column('E:E', 15, structure.white)
    ws.set_column('F:F', 65, structure.white)
    ws.set_column('G:G', 65, structure.white)
    ws.set_column('H:H', 3, structure.white)
    ws.set_column('I:I', 70, structure.white)
    ws.set_column('J:J', 70, structure.white)
    ws.set_column('K:K', 25, structure.white)
    ws.set_column('L:L', 25, structure.white)
    ws.set_column('M:M', 25, structure.white)
    ws.set_column('N:N', 25, structure.white)
    ws.set_column('O:O', 25, structure.white)
    ws.set_column('P:P', 25, structure.white)

    def create_grouping_1(unique_trrs):
        group_start = 13
        group_end = 43
        # for trr in unique_trrs:
        #     while group_start < group_end:
        while group_start <= group_end:
            ws.set_row(group_start, None, None, {'level': 0, 'hidden': False})
            group_start += 1

    def create_grouping_2(unique_trrs):
        group_start = 44
        group_end = 74
        # for trr in unique_trrs:
        #     while group_start < group_end:
        while group_start <= group_end:
            ws.set_row(group_start, None, None, {'level': 1, 'hidden': False})
            group_start += 1
            # group_end += 27

    if len(host_ids) > 1:
        total_systems = f'{len(host_ids)}'
    else:
        total_systems = f'{len(host_ids)}'

    ws.insert_image('A1', 'pipe_cleaner/img/vse_logo.png')

    ws.write('B5', f' Pipe Cleaner - {sheet_title}', structure.big_blue_font)
    ws.write('B6', f'       Kirkland Lab Site', structure.bold_italic_blue_font)
    ws.write('B7', f'       Pipe Number - {pipe_number}', structure.bold_italic_blue_font)
    ws.write('B8', f'       Total TRRs - {len(unique_trrs)}', structure.bold_italic_blue_font)
    ws.write('B10', f'       {Structure.date} - {Structure.time}', structure.italic_blue_font)

    ws.write('E2', f'', structure.big_blue_font)
    ws.write('F4', f'', structure.italic_blue_font)

    ws.write('B13', 'TRR ID', structure.teal_middle)
    ws.write('C13', 'Type', structure.teal_middle)
    ws.write('D13', 'Component', structure.teal_middle)
    ws.write('E13', 'Status', structure.teal_middle)
    ws.write('F13', 'Status Information', structure.teal_middle)
    ws.write('G13', 'General Notes', structure.teal_middle)
    ws.write('I13', 'Test Run Request', structure.teal_middle)
    ws.write('J13', 'Customer Requirements Document', structure.teal_middle)

    # Freeze Planes
    ws.freeze_panes(13, 4)

    # Groupings
    # create_grouping_1(unique_trrs)
    # create_grouping_2(unique_trrs)


def write_data(wb, ws, ids, crd, unique_trrs):
    structure = Structure(wb)

    def write_request_type_status(trr_id):  # Item 2
        component = 'Request Type'
        points_azure = []

        check_1 = 'STORAGE'
        check_2 = 'UTILITY'
        check_3 = '5.'
        check_4 = '6.'
        check_5 = '7.'
        check_6 = '.0'
        check_7 = '.1'
        check_8 = '.2'
        check_9 = '.3'
        check_10 = '.4'
        check_11 = '.5'
        check_12 = '.6'
        check_13 = 'XIO'
        check_14 = 'XSTORE'
        check_15 = 'COMPUTE'
        check_16 = 'SQL'
        check_17 = 'TE'
        check_18 = 'DB'
        check_19 = 'DW'
        check_20 = 'BALANCED'
        check_21 = 'WEB'
        check_22 = 'XDIRECT'
        check_23 = 'MBX'
        check_24 = 'GP'
        check_25 = 'LOW'
        check_26 = 'MID'
        check_27 = 'HIGH'
        check_28 = 'C2010'
        check_29 = 'C2020'
        check_30 = 'AMD'
        check_31 = 'OPTIMIZED'
        check_32 = 'AEP'
        check_33 = 'JBOF'
        check_34 = 'JBOD'
        check_35 = '4TIB'
        check_36 = '2TIB'
        check_37 = 'HM'
        check_38 = '1U'
        check_39 = 'ANALYTICS'
        check_40 = 'SSD'

        trr_azure = raw_target_configuration
        trr_upper = str(trr_azure).upper()
        try:
            crd_azure = crd_scanner.get_azure()[0]
        except IndexError:
            pass
        crd_upper = str(crd_azure).upper()

        def check_both(check):
            if check in trr_upper and check in crd_upper:
                add_total = + 1
                points_azure.append(add_total)

        check_both(check_1)
        check_both(check_2)
        check_both(check_3)
        check_both(check_4)
        check_both(check_5)
        check_both(check_6)
        check_both(check_7)
        check_both(check_8)
        check_both(check_9)
        check_both(check_10)
        check_both(check_11)
        check_both(check_12)
        check_both(check_13)
        check_both(check_14)
        check_both(check_15)
        check_both(check_16)
        check_both(check_17)
        check_both(check_18)
        check_both(check_19)
        check_both(check_20)
        check_both(check_21)
        check_both(check_22)
        check_both(check_23)
        check_both(check_24)
        check_both(check_25)
        check_both(check_26)
        check_both(check_27)
        check_both(check_28)
        check_both(check_29)
        check_both(check_30)
        check_both(check_31)
        check_both(check_32)
        check_both(check_33)
        check_both(check_34)
        check_both(check_35)
        check_both(check_36)
        check_both(check_37)
        check_both(check_38)
        check_both(check_39)
        check_both(check_40)

        if sum(points_azure) >= 3:
            ws.write(f'E{item_2}', f'MATCH', structure.good_cell)
            ws.write(f'F{item_2}', f'Target Configuration - TRR and CRD match', structure.alt_blue_middle)
            match_tally.append(1)
        else:
            ws.write(f'E{item_2}', f'MISMATCH', structure.bad_cell)
            ws.write(f'F{item_2}', f'Target Configuration - TRR and CRD do not match', structure.alt_blue_middle)
            summary = f'Mismatch = TRR {trr_id} - {extract_target(raw_target_configuration)} - {component}'
            mismatch_tally.append(1)
            mismatch_summary.append(summary)

    def write_bios_version_status(trr_id):  # Item 3
        component = 'BIOS Version'

        trr_bios = parsed_trr_bios
        trr_bios = str(trr_bios).split('.')[2]
        trr_upper = str(trr_bios).upper()
        crd_bios = crd_scanner.get_bios()[1]
        crd_bios = str(crd_bios).split('.')[2]
        crd_upper = str(crd_bios).upper()

        if trr_upper == '' or crd_upper == 'NONE':
            ws.write(f'E{item_3}', f'MISSING', structure.neutral_cell)
            ws.write(f'F{item_3}', f'TRR - {trr_bios}   vs   CRD - None', structure.alt_blue_middle)
            missing_tally.append(1)
        elif trr_upper == crd_upper:
            ws.write(f'E{item_3}', f'MATCH', structure.good_cell)
            ws.write(f'F{item_3}', f'TRR - {trr_bios}   vs   CRD - {crd_bios}', structure.blue_middle)
            match_tally.append(1)
        else:
            ws.write(f'E{item_3}', f'MISMATCH', structure.bad_cell)
            ws.write(f'F{item_3}', f'TRR - {trr_bios}   vs   CRD - {crd_bios}', structure.blue_middle)
            summary = f'Mismatch = TRR {trr_id} - {extract_target(raw_target_configuration)} - {component}'
            mismatch_tally.append(1)
            mismatch_summary.append(summary)
            mismatch_message = f'For {trr_id} - {extract_target(raw_target_configuration)}, we are missing {component}'
            mismatch_microsoft.append(mismatch_message)

    def write_bios_flavor_status(trr_id):  # Item 4  NEED TO BE FIXED
        component = 'BIOS Flavor'

        trr_bios = parsed_trr_bios
        trr_bios = str(trr_bios).split('.')[3]
        trr_upper = str(trr_bios).upper()
        crd_bios = crd_scanner.get_bios()[1]
        crd_bios = str(crd_bios).split('.')[3]
        crd_upper = str(crd_bios).upper()

        if trr_upper == '' or crd_upper == 'NONE':
            ws.write(f'E{item_4}', f'MISSING', structure.neutral_cell)
            ws.write(f'F{item_4}', f'TRR - {trr_bios}   vs   CRD (None)', structure.alt_blue_middle)
            missing_tally.append(1)
        elif trr_upper == crd_upper:
            ws.write(f'E{item_4}', f'MATCH', structure.good_cell)
            ws.write(f'F{item_4}', f'TRR - {trr_bios}   vs   CRD - {crd_bios}', structure.alt_blue_middle)
            match_tally.append(1)
        else:
            ws.write(f'E{item_4}', f'MISMATCH', structure.bad_cell)
            ws.write(f'F{item_4}', f'TRR - {trr_bios}   vs   CRD - {crd_bios}', structure.alt_blue_middle)
            summary = f'Mismatch = TRR {trr_id} - {extract_target(raw_target_configuration)} - {component}'
            mismatch_tally.append(1)
            mismatch_summary.append(summary)
            mismatch_message = f'For {trr_id} - {extract_target(raw_target_configuration)}, we are missing {component}'
            mismatch_microsoft.append(mismatch_message)

    def write_bmc_status(trr_id):  # Item 5
        component = 'BMC Version'
        trr_item = parsed_trr_bmc
        trr_upper = str(trr_item).upper()
        # crd_item = crd_scanner.get_bmc()[0][-3::]
        crd_item = crd_scanner.get_bmc()[0][10:-3]
        # crd_item = crd_item[:-3]
        crd_upper = str(crd_item).upper()

        if trr_upper == '' or crd_upper == 'NONE':
            ws.write(f'E{item_5}', f'MISSING', structure.neutral_cell)
            ws.write(f'F{item_5}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.blue_middle)
            missing_tally.append(1)
        elif trr_upper == crd_upper:
            ws.write(f'E{item_5}', f'MATCH', structure.good_cell)
            ws.write(f'F{item_5}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.blue_middle)
            match_tally.append(1)
        else:
            ws.write(f'E{item_5}', f'MISMATCH', structure.bad_cell)
            ws.write(f'F{item_5}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.blue_middle)
            summary = f'Mismatch = TRR {trr_id} - {extract_target(raw_target_configuration)} - {component}'
            mismatch_tally.append(1)
            mismatch_summary.append(summary)
            mismatch_message = f'For {trr_id} - {extract_target(raw_target_configuration)}, we are missing {component}'
            mismatch_microsoft.append(mismatch_message)

    def write_tpm_status(trr_id):  # Item 6
        component = 'TPM Version'
        trr_item = parsed_trr_tpm
        trr_upper = str(trr_item).upper()
        crd_item = crd_scanner.get_tpm()[0][:4:]
        crd_upper = str(crd_item).upper()

        if trr_upper == '' or crd_upper == 'NONE':
            ws.write(f'E{item_6}', f'MATCH', structure.good_cell)
            ws.write(f'F{item_6}', f'Missing for TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)
            # missing_tally.append(1)
        elif trr_upper == crd_upper:
            ws.write(f'E{item_6}', f'WAIVED', structure.alt_blue_middle)
            ws.write(f'F{item_6}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)
            # match_tally.append(1)
        else:
            ws.write(f'E{item_6}', f'WAIVED', structure.alt_blue_middle)
            ws.write(f'F{item_6}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)
            # summary = f'Mismatch = TRR {trr_id} - {extract_target(raw_target_configuration)} - {component}'
            # mismatch_tally.append(1)
            # mismatch_summary.append(summary)
            # mismatch_message = f'For {trr_id} - {extract_target(raw_target_configuration)}, we are missing {component}'
            # mismatch_microsoft.append(mismatch_message)

    def write_cpld_status(trr_id):  # Item 7 NEED TO FIX
        component = 'CPLD Version'
        trr_item = raw_cpld
        trr_upper = str(trr_item).upper()
        crd_upper = crd_scanner.get_cpld()[0][1:]
        crd_upper = crd_upper[:2]

        if trr_upper == '' or crd_upper == 'NONE':
            ws.write(f'E{item_7}', f'MISSING', structure.neutral_cell)
            ws.write(f'F{item_7}', f'TRR - {trr_item}   vs   CRD - {crd_upper}', structure.blue_middle)
            missing_tally.append(1)
        elif trr_upper == crd_upper:
            ws.write(f'E{item_7}', f'MATCH', structure.good_cell)
            ws.write(f'F{item_7}', f'TRR - {trr_item}   vs   CRD - {crd_upper}', structure.blue_middle)
            match_tally.append(1)
        else:
            ws.write(f'E{item_7}', f'MISMATCH', structure.bad_cell)
            ws.write(f'F{item_7}', f'TRR - {trr_item}   vs   CRD - {crd_upper}', structure.blue_middle)
            summary = f'Mismatch = TRR {trr_id} - {extract_target(raw_target_configuration)} - {component}'
            mismatch_tally.append(1)
            mismatch_summary.append(summary)
            mismatch_message = f'For {trr_id} - {extract_target(raw_target_configuration)}, we are missing {component}'
            mismatch_microsoft.append(mismatch_message)

    def write_chip_driver_status(id_trr):  # Item 8
        component = 'Chipset Driver'
        trr_item = raw_chipset
        trr_upper = str(trr_item).upper()
        try:
            crd_item = crd_scanner.get_chipset()[0]
        except IndexError:
            crd_item = 'NONE'
        crd_upper = str(crd_item).upper()

        if trr_upper == '' or crd_upper == 'NONE':
            ws.write(f'E{item_8}', f'MISSING', structure.neutral_cell)
            ws.write(f'F{item_8}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)
            missing_tally.append(1)
        elif trr_upper == crd_upper:
            ws.write(f'E{item_8}', f'MATCH', structure.good_cell)
            ws.write(f'F{item_8}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)
            match_tally.append(1)
        else:
            ws.write(f'E{item_8}', f'MISMATCH', structure.bad_cell)
            ws.write(f'F{item_8}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)
            summary = f'Mismatch = TRR {id_trr} - {extract_target(raw_target_configuration)} - {component}'
        mismatch_tally.append(1)
        # mismatch_summary.append(summary) # FIX IT
        mismatch_message = f'For {id_trr} - {extract_target(raw_target_configuration)}, we are missing {component}'
        mismatch_microsoft.append(mismatch_message)

    # Item 9 - Server Processor
    def write_processor_status(id_trr):
        component = 'Server Processor'
        trr_item = raw_processor
        trr_upper = str(trr_item).upper()

        def check_cpu(processor):
            if 'INTEL' in processor:
                return 'Intel'
            elif 'AMD' in processor:
                return 'AMD'

        if trr_upper == '' or trr_upper == 'NONE':
            ws.write(f'E{item_9}', f'MISSING', structure.neutral_cell)
            ws.write(f'F{item_9}', f'Not Present in TRR - Server Processor ({check_cpu(trr_upper)})',
                     structure.blue_middle)
            missing_tally.append(1)

            summary = f'Mismatch = TRR {id_trr} - {extract_target(raw_target_configuration)} - {component}'
            mismatch_summary.append(summary)
            mismatch_message = f'For {id_trr} - {extract_target(raw_target_configuration)}, we are missing {component}'
            mismatch_microsoft.append(mismatch_message)
        else:
            ws.write(f'E{item_9}', f'PRESENT', structure.good_cell)
            ws.write(f'F{item_9}', f'Present in TRR - Server Processor ({check_cpu(trr_upper)})',
                     structure.blue_middle)
            match_tally.append(1)

    # Item 10 - FPGA Release Package Status
    def write_fpga_release_package_status(trr_id):
        component = 'FPGA Release Version'
        trr_item = raw_fpga_release
        trr_upper = str(trr_item).upper()
        try:
            crd_item = crd_scanner.get_fpga_release()[0]
        except IndexError:
            crd_item = 'NONE'
        crd_upper = str(crd_item).upper()

        if trr_upper == '' or crd_upper == 'N':
            ws.write(f'E{item_10}', f'MISSING', structure.neutral_cell)
            ws.write(f'F{item_10}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)
            missing_tally.append(1)
        elif trr_upper == crd_upper:
            ws.write(f'E{item_10}', f'MATCH', structure.good_cell)
            ws.write(f'F{item_10}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)
            match_tally.append(1)
        else:
            ws.write(f'E{item_10}', f'MISMATCH', structure.bad_cell)
            ws.write(f'F{item_10}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)
            summary = f'Mismatch = TRR {trr_id} - {extract_target(raw_target_configuration)} - {component}'
            mismatch_tally.append(1)
            mismatch_summary.append(summary)
            mismatch_message = f'For {trr_id} - {extract_target(raw_target_configuration)}, we are missing {component}'
            mismatch_microsoft.append(mismatch_message)

    # Item 11 - FPGA Hyperblaster DLL
    def write_fpga_hyperblaster_dll(trr_id):
        component = 'FPGA Hypterblaster DLL'
        trr_item = raw_fpga_hyperblaster
        trr_upper = str(trr_item).upper()
        try:
            crd_item = crd_scanner.get_fpga_hyperblaster()[0]
        except IndexError:
            crd_item = 'NONE'
        crd_upper = str(crd_item).upper()

        if trr_upper == '' or crd_upper == 'N':
            ws.write(f'E{item_11}', f'MISSING', structure.neutral_cell)
            ws.write(f'F{item_11}', f'TRR - {trr_item}   vs   CRD - None', structure.blue_middle)
            missing_tally.append(1)
        elif trr_upper == crd_upper:
            ws.write(f'E{item_11}', f'MATCH', structure.good_cell)
            ws.write(f'F{item_11}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.blue_middle)
            match_tally.append(1)
        else:
            ws.write(f'E{item_11}', f'MISMATCH', structure.bad_cell)
            ws.write(f'F{item_11}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.blue_middle)
            summary = f'Mismatch = TRR {trr_id} - {extract_target(raw_target_configuration)} - {component}'
            mismatch_tally.append(1)
            mismatch_summary.append(summary)
            mismatch_message = f'For {trr_id} - {extract_target(raw_target_configuration)}, we are missing {component}'
            mismatch_microsoft.append(mismatch_message)

    # Item 12 - FPGA HIP
    def write_fpga_hip(trr_id):
        component = 'FPGA HIP'
        trr_item = raw_fpga_hip
        trr_upper = str(trr_item).upper()
        try:
            crd_item = crd_scanner.get_fpga_hip()[0]
        except IndexError:
            crd_item = 'NONE'
        crd_upper = str(crd_item).upper()

        if trr_upper == '' or crd_upper == 'N':
            ws.write(f'E{item_12}', f'MISSING', structure.neutral_cell)
            ws.write(f'F{item_12}', f'TRR - {trr_item}   vs   CRD - None', structure.alt_blue_middle)
            missing_tally.append(1)
        elif trr_upper == crd_upper:
            ws.write(f'E{item_12}', f'MATCH', structure.good_cell)
            ws.write(f'F{item_12}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)
            match_tally.append(1)
        else:
            ws.write(f'E{item_12}', f'MISMATCH', structure.bad_cell)
            ws.write(f'F{item_12}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)
            summary = f'Mismatch = TRR {trr_id} - {extract_target(raw_target_configuration)} - {component}'
            mismatch_tally.append(1)
            mismatch_summary.append(summary)
            mismatch_message = f'For {trr_id} - {extract_target(raw_target_configuration)}, we are missing {component}'
            mismatch_microsoft.append(mismatch_message)

    # Item 13 - FPGA Filter Status
    def write_fpga_filter_status(trr_id):
        component = 'FPGA Filter Status'
        trr_item = raw_fpga_filter
        trr_upper = str(trr_item).upper()
        try:
            crd_item = crd_scanner.get_fpga_filter()[0]
        except IndexError:
            crd_item = 'NONE'
        crd_upper = str(crd_item).upper()

        if trr_upper == '' or crd_upper == 'N':
            ws.write(f'E{item_13}', f'MISSING', structure.neutral_cell)
            ws.write(f'F{item_13}', f'TRR - {trr_item}   vs   CRD - None', structure.blue_middle)
            missing_tally.append(1)
        elif trr_upper == crd_upper:
            ws.write(f'E{item_13}', f'MATCH', structure.good_cell)
            ws.write(f'F{item_13}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.blue_middle)
            match_tally.append(1)
        else:
            ws.write(f'E{item_13}', f'MISMATCH', structure.bad_cell)
            ws.write(f'F{item_13}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.blue_middle)
            summary = f'Mismatch = TRR {trr_id} - {extract_target(raw_target_configuration)} - {component}'
            mismatch_tally.append(1)
            mismatch_summary.append(summary)
            mismatch_message = f'For {trr_id} - {extract_target(raw_target_configuration)}, we are missing {component}'
            mismatch_microsoft.append(mismatch_message)

    # Item 14 - FTDI port
    def write_ftdi_port(trr_id):
        component = 'FTDI Port'
        trr_item = raw_ftdi_port
        trr_upper = str(trr_item).upper()
        try:
            crd_item = crd_scanner.get_ftdi_port()[0]
        except IndexError:
            crd_item = 'NONE'
        crd_upper = str(crd_item).upper()

        if trr_upper == '' or crd_upper == 'N':
            ws.write(f'E{item_14}', f'MISSING', structure.neutral_cell)
            ws.write(f'F{item_14}', f'TRR - {trr_item}   vs   CRD - None', structure.alt_blue_middle)
            missing_tally.append(1)
        elif trr_upper == crd_upper:
            ws.write(f'E{item_14}', f'MATCH', structure.good_cell)
            ws.write(f'F{item_14}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)
            match_tally.append(1)
        else:
            ws.write(f'E{item_14}', f'MISMATCH', structure.bad_cell)
            ws.write(f'F{item_14}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)
            summary = f'Mismatch = TRR {trr_id} - {extract_target(raw_target_configuration)} - {component}'
            mismatch_tally.append(1)
            mismatch_summary.append(summary)
            mismatch_message = f'For {trr_id} - {extract_target(raw_target_configuration)}, we are missing {component}'
            mismatch_microsoft.append(mismatch_message)

    # Item 15 - FTDI Bus
    def write_ftdi_bus(trr_id):
        component = 'FTDI Filter'
        trr_item = raw_ftdi_port
        trr_upper = str(trr_item).upper()
        try:
            crd_item = crd_scanner.get_ftdi_bus()[0]
        except IndexError:
            crd_item = 'NONE'
        crd_upper = str(crd_item).upper()
        print(crd_upper)

        if trr_upper == '' or crd_upper == 'N':
            ws.write(f'E{item_15}', f'MISSING', structure.neutral_cell)
            ws.write(f'F{item_15}', f'TRR - {trr_item}   vs   CRD - None', structure.blue_middle)
            missing_tally.append(1)
        elif trr_upper == crd_upper:
            ws.write(f'E{item_15}', f'MATCH', structure.good_cell)
            ws.write(f'F{item_15}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.blue_middle)
            match_tally.append(1)
        else:
            ws.write(f'E{item_15}', f'MISMATCH', structure.bad_cell)
            ws.write(f'F{item_15}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.blue_middle)
            summary = f'Mismatch = TRR {trr_id} - {extract_target(raw_target_configuration)} - {component}'
            mismatch_tally.append(1)
            mismatch_summary.append(summary)
            mismatch_message = f'For {trr_id} - {extract_target(raw_target_configuration)}, we are missing {component}'
            mismatch_microsoft.append(mismatch_message)

    def write_nic_firmware_status(trr_id):  # Item 11
        component = 'NIC Firmware'
        trr_item = raw_nic_firmware
        trr_upper = str(trr_item).upper()
        try:
            crd_item = crd_scanner.get_nic()[0]
        except IndexError:
            crd_item = 'NONE'
        crd_upper = str(crd_item).upper()

        if trr_upper == '' or crd_upper == 'NONE':
            ws.write(f'E{item_16}', f'MISSING', structure.neutral_cell)
            ws.write(f'F{item_16}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)
            mismatch_tally.append(1)
        elif trr_upper == crd_upper:
            ws.write(f'E{item_16}', f'MATCH', structure.good_cell)
            ws.write(f'F{item_16}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)
            match_tally.append(1)
        else:
            ws.write(f'E{item_16}', f'MISMATCH', structure.bad_cell)
            ws.write(f'F{item_16}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)
            summary = f'Mismatch = TRR {trr_id} - {extract_target(raw_target_configuration)} - {component}'
            mismatch_tally.append(1)
            mismatch_summary.append(summary)
            mismatch_message = f'For {trr_id} - {extract_target(raw_target_configuration)}, we are missing {component}'
            mismatch_microsoft.append(mismatch_message)

    def write_nic_pxe_status(trr_id):  # Item 12
        component = 'NIC PXE'
        trr_item = raw_nic_pxe
        trr_upper = str(trr_item).upper()
        try:
            crd_item = crd_scanner.get_nic_pxe()[0]
        except IndexError:
            crd_item = 'NONE'
        crd_upper = str(crd_item).upper()

        if trr_upper == '' or crd_upper == 'NONE':
            ws.write(f'E{item_17}', f'MISSING', structure.neutral_cell)
            ws.write(f'F{item_17}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.blue_middle)
            mismatch_tally.append(1)
        elif trr_upper == crd_upper:
            ws.write(f'E{item_17}', f'MATCH', structure.good_cell)
            ws.write(f'F{item_17}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.blue_middle)
            match_tally.append(1)
        else:
            ws.write(f'E{item_17}', f'MISMATCH', structure.bad_cell)
            ws.write(f'F{item_17}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.blue_middle)
            summary = f'Mismatch = TRR {trr_id} - {extract_target(raw_target_configuration)} - {component}'
            mismatch_tally.append(1)
            mismatch_summary.append(summary)
            mismatch_message = f'For {trr_id} - {extract_target(raw_target_configuration)}, we are missing {component}'
            mismatch_microsoft.append(mismatch_message)

    def write_nic_uefi_status(trr_id):  # Item 13 NEED TO FIX
        component = 'NIC UEFI'
        trr_item = raw_nic_uefi
        trr_upper = str(trr_item).upper()
        try:
            crd_item = crd_scanner.get_nic_pxe()[0]
        except IndexError:
            crd_item = 'NONE'
        crd_upper = str(crd_item).upper()

        if trr_upper == '' or crd_upper == 'NONE':
            ws.write(f'E{item_18}', f'MISSING', structure.neutral_cell)
            ws.write(f'F{item_18}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)
            mismatch_tally.append(1)
        elif trr_upper == crd_upper:
            ws.write(f'E{item_18}', f'MATCH', structure.good_cell)
            ws.write(f'F{item_18}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)
            match_tally.append(1)
        else:
            ws.write(f'E{item_18}', f'MISMATCH', structure.bad_cell)
            ws.write(f'F{item_18}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)
            summary = f'Mismatch = TRR {trr_id} - {extract_target(raw_target_configuration)} - {component}'
            mismatch_tally.append(1)
            mismatch_summary.append(summary)
            mismatch_message = f'For {trr_id} - {extract_target(raw_target_configuration)}, we are missing {component}'
            mismatch_microsoft.append(mismatch_message)

    def write_nic_driver_status(trr_id):  # Item 14 NEED TO FIX
        component = 'NIC Driver'
        trr_item = raw_nic_driver
        trr_upper = str(trr_item).upper()
        try:
            crd_item = crd_scanner.get_nic_pxe()[0]
        except IndexError:
            crd_item = 'NONE'
        crd_upper = str(crd_item).upper()

        if trr_upper == '' or crd_upper == 'NONE':
            ws.write(f'E{item_19}', f'MISSING', structure.neutral_cell)
            ws.write(f'F{item_19}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.blue_middle)
            mismatch_tally.append(1)
        elif trr_upper == crd_upper:
            ws.write(f'E{item_19}', f'MATCH', structure.good_cell)
            ws.write(f'F{item_19}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.blue_middle)
            match_tally.append(1)
        else:
            ws.write(f'E{item_19}', f'MISMATCH', structure.bad_cell)
            ws.write(f'F{item_19}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.blue_middle)
            summary = f'Mismatch = TRR {trr_id} - {extract_target(raw_target_configuration)} - {component}'
            mismatch_tally.append(1)
            mismatch_summary.append(summary)
            mismatch_message = f'For {trr_id} - {extract_target(raw_target_configuration)}, we are missing {component}'
            mismatch_microsoft.append(mismatch_message)

    def write_boot_drive_status(trr_id):  # Item 15
        component = 'Boot Drive'

        trr_item = raw_trr_boot_driver
        trr_upper = str(trr_item).upper()

        if trr_upper == '' or trr_upper == 'None':
            ws.write(f'E{item_20}', f'MISSING', structure.neutral_cell)
            ws.write(f'F{item_20}', f'MISSING for TRR - {trr_item})', structure.alt_blue_middle)
            mismatch_tally.append(1)

            summary = f'Mismatch = TRR {trr_id} - {extract_target(raw_target_configuration)} - {component}'
            mismatch_summary.append(summary)
            mismatch_message = f'For {trr_id} - {extract_target(raw_target_configuration)}, we are missing {component}'
            mismatch_microsoft.append(mismatch_message)
        else:
            ws.write(f'E{item_20}', f'PRESENT', structure.good_cell)
            ws.write(f'F{item_20}', f'Present in TRR - Boot Drive ({trr_item})', structure.alt_blue_middle)
            match_tally.append(1)


    def write_nvme_part_number_status(trr_id):  # Item 16 NEED TO FIX
        component = 'NVMe Part Number'
        ws.write(f'E{item_21}', f'WAIVED', structure.blue_middle)
        ws.write(f'F{item_21}', f'NVMe Part Number to be worked on', structure.blue_middle)
        # trr_item = raw_nvme_trr
        # trr_upper = str(trr_item).upper()
        # crd_item = crd_scanner.get_nvme_pn()[0]
        # crd_upper = str(crd_item).upper()
        #
        # if trr_upper == '' or crd_upper == 'NONE':
        #     ws.write(f'E{item_21}', f'MISSING', structure.neutral_cell)
        #     ws.write(f'F{item_21}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)
        #
        #     missing_tally.append(1)
        # elif trr_upper == crd_upper:
        #     ws.write(f'E{item_21}', f'MATCH', structure.good_cell)
        #     ws.write(f'F{item_21}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)
        #
        #     match_tally.append(1)
        # else:
        #     ws.write(f'E{item_21}', f'MISMATCH', structure.bad_cell)
        #     ws.write(f'F{item_21}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)
        #
        #     mismatch_tally.append(1)

    def write_nvme_version_status(trr_id):  # Item 17 NEED TO FIX
        component = 'NVMe Version'
        ws.write(f'E{item_22}', f'WAIVED', structure.alt_blue_middle)
        ws.write(f'F{item_22}', f'NVMe Version to be worked on', structure.alt_blue_middle)
        # trr_item = raw_nvme_trr
        # trr_upper = str(trr_item).upper()
        # crd_item = crd_scanner.get_nvme_pn()[0]
        # crd_upper = str(crd_item).upper()
        #
        # if trr_upper == '' or crd_upper == 'NONE':
        #     ws.write(f'E{item_22}', f'MISSING', structure.neutral_cell)
        #     ws.write(f'F{item_22}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.blue_middle)
        # elif trr_upper == crd_upper:
        #     ws.write(f'E{item_22}', f'MATCH', structure.good_cell)
        #     ws.write(f'F{item_22}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.blue_middle)
        # else:
        #     ws.write(f'E{item_22}', f'MISMATCH', structure.bad_cell)
        #     ws.write(f'F{item_22}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.blue_middle)

    def write_hdd_part_number_status(trr_id):  # Item 18 NEED TO FIX
        component = 'HDD Part Number'
        ws.write(f'E{item_23}', f'WAIVED', structure.blue_middle)
        ws.write(f'F{item_23}', f'HDD Part Number to be worked on', structure.blue_middle)
        # trr_item = raw_hdd_trr
        # trr_upper = str(trr_item).upper()
        # crd_item = crd_scanner.get_nvme_pn()[0]
        # crd_upper = str(crd_item).upper()
        #
        # if trr_upper == '' or crd_upper == 'NONE':
        #     ws.write(f'E{item_23}', f'MISSING', structure.neutral_cell)
        #     ws.write(f'F{item_23}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.blue_middle)
        # elif trr_upper == crd_upper:
        #     ws.write(f'E{item_23}', f'MATCH', structure.good_cell)
        #     ws.write(f'F{item_23}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.blue_middle)
        # else:
        #     ws.write(f'E{item_23}', f'MISMATCH', structure.bad_cell)
        #     ws.write(f'F{item_23}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.blue_middle)

    def write_hdd_version_status(trr_id):  # Item 19 NEED TO FIX
        component = 'HDD Version'
        ws.write(f'E{item_24}', f'WAIVED', structure.alt_blue_middle)
        ws.write(f'F{item_24}', f'HDD Version to be worked on', structure.alt_blue_middle)
        # trr_item = raw_hdd_trr
        # trr_upper = str(trr_item).upper()
        # crd_item = crd_scanner.get_nvme_pn()[0]
        # crd_upper = str(crd_item).upper()
        #
        # if trr_upper == '' or crd_upper == 'NONE':
        #     ws.write(f'E{item_24}', f'MISSING', structure.neutral_cell)
        #     ws.write(f'F{item_24}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)
        # elif trr_upper == crd_upper:
        #     ws.write(f'E{item_24}', f'MATCH', structure.good_cell)
        #     ws.write(f'F{item_24}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)
        # else:
        #     ws.write(f'E{item_24}', f'MISMATCH', structure.bad_cell)
        #     ws.write(f'F{item_24}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)

    def write_dimm_part_number_status(trr_id):  # Item 20 NEED TO FIX
        component = 'DIMM Part Number'
        ws.write(f'E{item_25}', f'WAIVED', structure.blue_middle)
        ws.write(f'F{item_25}', f'DIMM Part Number to be worked on', structure.blue_middle)
        # trr_item = raw_dimm_trr
        # trr_upper = str(trr_item).upper()
        # crd_item = crd_scanner.get_nvme_pn()[0]
        # crd_upper = str(crd_item).upper()
        #
        # if trr_upper == '' or crd_upper == 'NONE':
        #     ws.write(f'E{item_25}', f'MISSING', structure.neutral_cell)
        #     ws.write(f'F{item_25}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.blue_middle)
        # elif trr_upper == crd_upper:
        #     ws.write(f'E{item_25}', f'MATCH', structure.good_cell)
        #     ws.write(f'F{item_25}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.blue_middle)
        # else:
        #     ws.write(f'E{item_25}', f'MISMATCH', structure.bad_cell)
        #     ws.write(f'F{item_25}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.blue_middle)

    def write_dimm_version_status(trr_id):  # Item 21 NEED TO FIX
        component = 'DIMM Version'
        ws.write(f'E{item_26}', f'WAIVED', structure.alt_blue_middle)
        ws.write(f'F{item_26}', f'DIMM Version to be worked on', structure.alt_blue_middle)
        # trr_item = raw_dimm_trr
        # trr_upper = str(trr_item).upper()
        # crd_item = crd_scanner.get_nvme_pn()[0]
        # crd_upper = str(crd_item).upper()
        #
        # if trr_upper == '' or crd_upper == 'NONE':
        #     ws.write(f'E{item_26}', f'MISSING', structure.neutral_cell)
        #     ws.write(f'F{item_26}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)
        # elif trr_upper == crd_upper:
        #     ws.write(f'E{item_26}', f'MATCH', structure.good_cell)
        #     ws.write(f'F{item_26}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)
        # else:
        #     ws.write(f'E{item_26}', f'MISMATCH', structure.bad_cell)
        #     ws.write(f'F{item_26}', f'TRR - {trr_bmc}   vs   CRD - {crd_item}', structure.alt_blue_middle)

    def write_psu_part_number_status(trr_id):  # Item 22 NEED TO FIX
        component = 'PSU Part Number'
        ws.write(f'E{item_27}', f'WAIVED', structure.blue_middle)
        ws.write(f'F{item_27}', f'PSU Part Number to be worked on', structure.blue_middle)
        # trr_item = raw_psu_part_number
        # trr_upper = str(trr_item).upper()
        # crd_item = crd_scanner.get_psu()[0][:28:]
        # crd_upper = str(crd_item).upper()
        #
        # if trr_upper == '' or crd_upper == 'NONE':
        #     ws.write(f'E{item_27}', f'MISSING', structure.neutral_cell)
        #     ws.write(f'F{item_27}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)
        # elif trr_upper == crd_upper:
        #     ws.write(f'E{item_27}', f'MATCH', structure.good_cell)
        #     ws.write(f'F{item_27}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)
        # else:
        #     ws.write(f'E{item_27}', f'MISMATCH', structure.bad_cell)
        #     ws.write(f'F{item_27}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)

    def write_psu_version_status(trr_id):  # Item 23 NEED TO FIX
        component = 'PSU Version'
        ws.write(f'E{item_28}', f'WAIVED', structure.alt_blue_middle)
        ws.write(f'F{item_28}', f'PSU Version to be worked on', structure.alt_blue_middle)
        # trr_item = raw_psu_version
        # trr_upper = str(trr_item).upper()
        # crd_item = crd_scanner.get_nvme_pn()[0]
        # crd_item = crd_item[28:][:8:]
        # crd_upper = str(crd_item).upper()
        # print(crd_item)
        #
        # if trr_upper == '' or crd_upper == 'NONE':
        #     ws.write(f'E{item_28}', f'MISSING', structure.neutral_cell)
        #     ws.write(f'F{item_28}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.blue_middle)
        # elif trr_upper == crd_upper:
        #     ws.write(f'E{item_28}', f'MATCH', structure.good_cell)
        #     ws.write(f'F{item_28}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.blue_middle)
        # else:
        #     ws.write(f'E{item_28}', f'MISMATCH', structure.bad_cell)
        #     ws.write(f'F{item_28}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.blue_middle)

    def write_manager_switch_status(trr_id):  # Item 24 NEED TO FIX
        component = 'Manager Switch'
        ws.write(f'E{item_29}', f'WAIVED', structure.blue_middle)
        ws.write(f'F{item_29}', f'Manager Switch to be worked on', structure.blue_middle)
        # trr_item = raw_bios_trr
        # trr_upper = str(trr_item).upper()
        # crd_item = crd_scanner.get_nvme_pn()[0]
        # crd_upper = str(crd_item).upper()
        #
        # if trr_upper == '' or crd_upper == 'NONE':
        #     ws.write(f'E{item_29}', f'MISSING', structure.neutral_cell)
        #     ws.write(f'F{item_29}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)
        # elif trr_upper == crd_upper:
        #     ws.write(f'E{item_29}', f'MATCH', structure.good_cell)
        #     ws.write(f'F{item_29}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)
        # else:
        #     ws.write(f'E{item_29}', f'MISMATCH', structure.bad_cell)
        #     ws.write(f'F{item_29}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)

    def write_jbof_status(trr_id):  # Item 25 NEED TO FIX
        component = 'JBOF BMC Version'
        trr_item = raw_trr_jbof_bmc

        trr_upper = str(trr_item).upper()
        crd_item = crd_scanner.get_nvme_pn()[0]
        crd_upper = str(crd_item).upper()

        if trr_upper == '' or trr_upper == 'NONE':
            ws.write(f'E{item_30}', f'NOT PRESENT', structure.neutral_cell)
            ws.write(f'F{item_30}', f'TRR - {trr_item}', structure.alt_blue_middle)
            missing_tally.append(1)
        elif trr_upper == crd_upper:
            ws.write(f'E{item_30}', f'MATCH', structure.good_cell)
            ws.write(f'F{item_30}', f'TRR - {trr_item}', structure.alt_blue_middle)
            match_tally.append(1)
        else:
            ws.write(f'E{item_30}', f'NOT PRESENT', structure.bad_cell)
            ws.write(f'F{item_30}', f'TRR - {trr_item}', structure.alt_blue_middle)
            mismatch_tally.append(1)

    def extract_target(raw_target):
        target = str(raw_target).replace(']', '')
        target = target.split('[')[2]
        if 'XIOServer' in target:
            return 'XIO Server'
        elif 'XIOStorage' in target:
            return 'XIO Storage'
        elif 'StorageServer' in target:
            return 'Storage Server'
        else:
            return target

    def extract_request(request_type):
        request_upper = str(request_type).upper()
        if 'SSD' in request_upper:
            return 'SSD'
        elif 'HDD' in request_upper:
            return 'HDD'
        elif 'NVME' in request_upper:
            return 'NVMe'
        elif 'DIMM' in request_upper:
            return 'DIMM'
        elif 'MEMORY' in request_upper:
            return 'Memory'

    def check_request_type():
        if raw_request_type != '':
            ws.write(f'E{item_1}', f'PRESENT', structure.good_cell)
            match_tally.append(1)
        else:
            ws.write(f'E{item_1}', f'NOT PRESENT', structure.bad_cell)
            mismatch_tally.append(1)


    start = 0
    previous = 0

    while start < len(unique_trrs):
        total = start + previous

        # Gets unique trr id from input file then requests one of the unique TRR IDs
        trr_id = unique_trrs[start]
        request_ado(trr_id)

        with open(f'{Path.info}{str(trr_id)}/final.json') as file:
            trr = loads(file.read())

        raw_target_configuration = requested_configuration(trr, 'TARGET', 'CONFIGURATION', 'CONFIGURATION')
        raw_bios = requested_configuration(trr, 'SERVER', 'BI', 'OS')
        raw_bmc = requested_configuration(trr, 'SERVER', 'BMC', 'BMC')
        raw_tpm = requested_configuration(trr, 'SERVER', 'TPM', 'TPM')
        raw_cpld = requested_configuration(trr, 'SERVER', 'SERVER CPLD', 'CPLD')
        raw_chipset = requested_configuration(trr, 'CHIPSET', 'CHIPSET', 'DRIVER')
        raw_processor = requested_configuration(trr, 'PROCESSORS', 'PROCESSORS', 'PROCESSORS')
        raw_fpga_release = requested_configuration(trr, 'FPGA', 'RELEASE', 'PACKAGE')
        raw_fpga_hyperblaster = requested_configuration(trr, 'SERVER', 'HYPERBLASTER', 'DRIVER')
        raw_fpga_hip = requested_configuration(trr, 'SERVER', 'FPGA', 'HIP')
        raw_fpga_filter = requested_configuration(trr, 'SERVER', 'FPGA', 'FILTER')
        raw_ftdi_port = requested_configuration(trr, 'SERVER', 'FTDI', 'PORT')
        raw_ftdi_bus = requested_configuration(trr, 'SERVER', 'FTDI', 'BUS')
        raw_nic_firmware = requested_configuration(trr, 'NIC', 'FIRMWARE', 'FIRMWARE')
        raw_nic_pxe = requested_configuration(trr, 'NIC', 'PXE', 'PXE')
        raw_nic_uefi = requested_configuration(trr, 'NIC', 'UEFI', 'UEFI')
        raw_nic_driver = requested_configuration(trr, 'NIC', 'DRIVER', 'DRIVER')
        raw_nvme = requested_configuration(trr, 'QCL', 'NVME', 'NVME')
        raw_hdd = requested_configuration(trr, 'QCL', 'HDD', 'HDD')
        raw_request_type = requested_configuration(trr, 'REQUEST', 'TYPE', 'TYPE')
        raw_dimm = requested_configuration(trr, 'DIMM', 'DIMM', '1')
        raw_psu_pn = requested_configuration(trr, 'PSU', 'PSU', 'PN')
        raw_psu_firmware = requested_configuration(trr, 'PSU', 'PSU', 'FIRMWARE')
        raw_trr_boot_driver = requested_configuration(trr, 'BOOT', 'BOOT', 'DRIVE')
        raw_trr_psu_version = requested_configuration(trr, 'PSU', 'PSU', 'FIRMWARE')
        raw_trr_jbof_bmc = requested_configuration(trr, 'JBOF', 'JFBOF', 'JBOF')

        parsed_trr_bios = raw_bios
        parsed_trr_bmc = raw_bmc.split('.')[2][-3::]
        parsed_trr_tpm = raw_tpm.replace('V', '')[:4:]

        item_1 = Structure.initial + total + 2  # Request Type
        item_2 = Structure.initial + total + 3  # Target Configuration
        item_3 = Structure.initial + total + 4  # BIOS Version
        item_4 = Structure.initial + total + 5  # BIOS Flavor
        item_5 = Structure.initial + total + 6  # BMC
        item_6 = Structure.initial + total + 7  # TPM
        item_7 = Structure.initial + total + 8  # CPLD
        item_8 = Structure.initial + total + 9  # Chipset Driver
        item_9 = Structure.initial + total + 10  # Server Processor
        item_10 = Structure.initial + total + 11  # FPGA Release Package
        item_11 = Structure.initial + total + 12  # FPGA Hyperblaster DLL
        item_12 = Structure.initial + total + 13  # FPGA HIP Driver
        item_13 = Structure.initial + total + 14  # FPGA Filter Driver
        item_14 = Structure.initial + total + 15  # FTDI Port Driver
        item_15 = Structure.initial + total + 16  # FTDI Bus Driver
        item_16 = Structure.initial + total + 17  # NIC Firmware
        item_17 = Structure.initial + total + 18  # NIC PXE
        item_18 = Structure.initial + total + 19  # NIC UEFI
        item_19 = Structure.initial + total + 20  # NIC Driver
        item_20 = Structure.initial + total + 21  # Boot Drive
        item_21 = Structure.initial + total + 22  # NVMe Part Number
        item_22 = Structure.initial + total + 23  # NVMe Version
        item_23 = Structure.initial + total + 24  # HDD Part Number
        item_24 = Structure.initial + total + 25  # HDD Version
        item_25 = Structure.initial + total + 26  # DIMM Part Number
        item_26 = Structure.initial + total + 27  # DIMM Version
        item_27 = Structure.initial + total + 28  # PSU Part Number
        item_28 = Structure.initial + total + 29  # PSU Firmware
        item_29 = Structure.initial + total + 30  # Manager Switch Firmware
        item_30 = Structure.initial + total + 31  # JBOF - BMC

        # TRR ID Column
        ws.merge_range(f'B{item_1}:B{item_30}', f'{int(trr_id)}', structure.merge_format)

        # Type Column
        ws.write(f'C{item_1}', f'Request', structure.blue_middle)
        ws.write(f'C{item_30}', f'JBOF/F2010', structure.blue_middle)
        ws.merge_range(f'C{item_2}:C{item_30}', f'{extract_target(raw_target_configuration)}',
                       structure.merge_format)

        ws.write(f'D{item_1}', f'{extract_request(raw_request_type)} Test', structure.blue_middle)
        ws.write(f'D{item_2}', f'Target Configuration', structure.alt_blue_middle)
        ws.write(f'D{item_3}', f'BIOS Version', structure.blue_middle)
        ws.write(f'D{item_4}', f'BIOS Flavor', structure.alt_blue_middle)
        ws.write(f'D{item_5}', f'BMC Version', structure.blue_middle)
        ws.write(f'D{item_6}', f'TPM Version', structure.alt_blue_middle)
        ws.write(f'D{item_7}', f'CPLD Version', structure.blue_middle)
        ws.write(f'D{item_8}', f'Chipset Driver', structure.alt_blue_middle)
        ws.write(f'D{item_9}', f'Server Processor', structure.blue_middle)
        ws.write(f'D{item_10}', f'FPGA Release Package', structure.alt_blue_middle)
        ws.write(f'D{item_11}', f'Hyperblaster DLL', structure.blue_middle)
        ws.write(f'D{item_12}', f'FPGA HIP Driver', structure.alt_blue_middle)
        ws.write(f'D{item_13}', f'FPGA Filter Driver', structure.blue_middle)
        ws.write(f'D{item_14}', f'FTDI Port Driver', structure.alt_blue_middle)
        ws.write(f'D{item_15}', f'FTDI Bus Driver', structure.blue_middle)
        ws.write(f'D{item_16}', f'NIC Firmware', structure.alt_blue_middle)
        ws.write(f'D{item_17}', f'NIC PXE', structure.blue_middle)
        ws.write(f'D{item_18}', f'NIC UEFI', structure.alt_blue_middle)
        ws.write(f'D{item_19}', f'NIC Driver', structure.blue_middle)
        ws.write(f'D{item_20}', f'Boot Drive', structure.alt_blue_middle)
        ws.write(f'D{item_21}', f'NVMe Part Number', structure.blue_middle)
        ws.write(f'D{item_22}', f'NVMe Version', structure.alt_blue_middle)
        ws.write(f'D{item_23}', f'HDD Part Number', structure.blue_middle)
        ws.write(f'D{item_24}', f'HDD Version', structure.alt_blue_middle)
        ws.write(f'D{item_25}', f'DIMM Part Number', structure.blue_middle)
        ws.write(f'D{item_26}', f'DIMM Version', structure.alt_blue_middle)
        ws.write(f'D{item_27}', f'PSU Part Number', structure.blue_middle)
        ws.write(f'D{item_28}', f'PSU Firmware', structure.alt_blue_middle)
        ws.write(f'D{item_29}', f'Manager Switch Firmware', structure.blue_middle)
        ws.write(f'D{item_30}', f'JBOF BMC Version', structure.alt_blue_middle)

        # Status Information Column
        check_request_type()
        ws.write(f'E{item_6}', f'WAIVED', structure.alt_blue_middle)

        ws.write(f'F{item_1}', f'Checks if Request Type is present in TRR', structure.blue_middle)
        ws.write(f'F{item_2}', f'', structure.missing_cell)  # NEED TO FIX
        ws.write(f'F{item_3}', f'', structure.missing_cell)
        ws.write(f'F{item_4}', f'', structure.missing_cell)
        ws.write(f'F{item_5}', f'', structure.missing_cell)
        ws.write(f'F{item_6}', f'Do not update firmware, might brick motherboard', structure.alt_blue_middle)
        ws.write(f'F{item_7}', f'', structure.missing_cell)
        ws.write(f'F{item_8}', f'', structure.missing_cell)
        ws.write(f'F{item_9}', f'', structure.missing_cell)
        ws.write(f'F{item_10}', f'', structure.missing_cell)
        ws.write(f'F{item_11}', f'', structure.missing_cell)
        ws.write(f'F{item_12}', f'', structure.missing_cell)
        ws.write(f'F{item_13}', f'', structure.missing_cell)
        ws.write(f'F{item_14}', f'', structure.missing_cell)
        ws.write(f'F{item_15}', f'', structure.missing_cell)
        ws.write(f'F{item_16}', f'', structure.missing_cell)
        ws.write(f'F{item_17}', f'', structure.missing_cell)
        ws.write(f'F{item_18}', f'', structure.missing_cell)
        ws.write(f'F{item_19}', f'', structure.missing_cell)
        ws.write(f'F{item_20}', f'', structure.missing_cell)
        ws.write(f'F{item_21}', f'', structure.missing_cell)
        ws.write(f'F{item_22}', f'', structure.missing_cell)
        ws.write(f'F{item_23}', f'', structure.missing_cell)
        ws.write(f'F{item_24}', f'', structure.missing_cell)
        ws.write(f'F{item_25}', f'', structure.missing_cell)
        ws.write(f'F{item_26}', f'', structure.missing_cell)
        ws.write(f'F{item_27}', f'', structure.missing_cell)
        ws.write(f'F{item_28}', f'', structure.missing_cell)
        ws.write(f'F{item_29}', f'', structure.missing_cell)
        ws.write(f'F{item_30}', f'Checks for JBOF/F2010 is in TRR', structure.blue_middle)

        # General Notes
        ws.write(f'G{item_1}', f'Request Types only show up in TRRs, not CRDs', structure.blue_middle)
        ws.write(f'G{item_2}', f'', structure.missing_cell)
        ws.write(f'G{item_3}', f'', structure.missing_cell)
        ws.write(f'G{item_4}', f'', structure.missing_cell)
        ws.write(f'G{item_5}', f'Use BMC 4.60 or higher for Gen 6', structure.blue_middle)
        ws.write(f'G{item_6}', f'Do not update, might brick motherboard', structure.alt_blue_middle)
        ws.write(f'G{item_7}', f'', structure.missing_cell)
        ws.write(f'G{item_8}', f'', structure.missing_cell)
        ws.write(f'G{item_9}', f'Only Available in TRR, Comes from SKUDOC', structure.blue_middle)
        ws.write(f'G{item_10}', f'', structure.missing_cell)
        ws.write(f'G{item_11}', f'', structure.missing_cell)
        ws.write(f'G{item_12}', f'', structure.missing_cell)
        ws.write(f'G{item_13}', f'', structure.missing_cell)
        ws.write(f'G{item_14}', f'', structure.missing_cell)
        ws.write(f'G{item_15}', f'', structure.missing_cell)
        ws.write(f'G{item_16}', f'', structure.missing_cell)
        ws.write(f'G{item_17}', f'', structure.missing_cell)
        ws.write(f'G{item_18}', f'', structure.missing_cell)
        ws.write(f'G{item_19}', f'', structure.missing_cell)
        ws.write(f'G{item_20}', f'Only Available in TRR, Comes from TRR Only', structure.alt_blue_middle)
        ws.write(f'G{item_21}', f'', structure.missing_cell)
        ws.write(f'G{item_22}', f'', structure.missing_cell)
        ws.write(f'G{item_23}', f'', structure.missing_cell)
        ws.write(f'G{item_24}', f'', structure.missing_cell)
        ws.write(f'G{item_25}', f'', structure.missing_cell)
        ws.write(f'G{item_26}', f'', structure.missing_cell)
        ws.write(f'G{item_27}', f'', structure.missing_cell)
        ws.write(f'G{item_28}', f'', structure.missing_cell)
        ws.write(f'G{item_29}', f'', structure.missing_cell)
        ws.write(f'G{item_30}', f'Need to check if JBOF or F2010 are in TRRs', structure.alt_blue_middle)

        ws.write_comment(f'G{item_5}', f'Make sure to use BMC 4.60 or higher for all Intel-Based Gen 6 WCS, '
                                       f'including xStore, xDirect and XIO Storage - MSFT, 8/3/2020', {'height': 200})
        ws.write_comment(f'G{item_6}', f'DO NOT attempt to update the TPM firmware. This is very likely to brick the '
                                       f'motherboard and should not be attempted without specific instructions.'
                                       f' - Eric Johnson, 5/14/2020', {'height': 200})

        def return_raw_trr(item, raw):
            def odd_or_even(position):
                if position % 2 == 0:
                    return 'EVEN'
                else:
                    return 'ODD'

            if raw == '' or raw == None:
                ws.write(f'I{item}', f'', structure.missing_cell)
            elif 'EVEN' == odd_or_even(item):
                ws.write(f'I{item}', f'{raw}', structure.blue_middle)
            else:
                ws.write(f'I{item}', f'{raw}', structure.alt_blue_middle)

        # Test Run Request Column
        return_raw_trr(item_1, raw_request_type)
        return_raw_trr(item_2, raw_target_configuration)
        return_raw_trr(item_3, raw_bios)
        return_raw_trr(item_4, raw_bios)
        return_raw_trr(item_5, raw_bmc)
        return_raw_trr(item_6, raw_tpm)
        return_raw_trr(item_7, raw_cpld)
        return_raw_trr(item_8, raw_chipset)
        return_raw_trr(item_9, raw_processor)
        return_raw_trr(item_10, raw_fpga_release)
        return_raw_trr(item_11, raw_fpga_hyperblaster)
        return_raw_trr(item_12, raw_fpga_hip)
        return_raw_trr(item_13, raw_fpga_filter)
        return_raw_trr(item_14, raw_ftdi_port)
        return_raw_trr(item_15, raw_ftdi_bus)
        return_raw_trr(item_16, raw_nic_firmware)
        return_raw_trr(item_17, raw_nic_pxe)
        return_raw_trr(item_18, raw_nic_uefi)
        return_raw_trr(item_19, raw_nic_driver)
        return_raw_trr(item_20, raw_trr_boot_driver)
        return_raw_trr(item_21, raw_nvme)
        return_raw_trr(item_22, raw_hdd)
        return_raw_trr(item_23, raw_hdd)
        return_raw_trr(item_24, raw_dimm)
        return_raw_trr(item_25, raw_dimm)
        return_raw_trr(item_26, raw_dimm)
        return_raw_trr(item_27, raw_psu_pn)
        return_raw_trr(item_28, raw_psu_firmware)
        return_raw_trr(item_29, raw_psu_firmware)
        return_raw_trr(item_30, 'JBOF')

        # Customer Requirements Document Column
        ws.write(f'J{item_1}', f'', structure.missing_cell)
        ws.write(f'J{item_2}', f'{crd_scanner.get_azure()[0]}', structure.alt_blue_middle)
        ws.write(f'J{item_3}', f'', structure.blue_middle)
        ws.write(f'J{item_4}', f'', structure.alt_blue_middle)
        ws.write(f'J{item_5}', f'', structure.blue_middle)
        ws.write(f'J{item_6}', f'', structure.alt_blue_middle)
        ws.write(f'J{item_7}', f'', structure.blue_middle)
        ws.write(f'J{item_8}', f'', structure.alt_blue_middle)
        ws.write(f'J{item_9}', f'', structure.blue_middle)
        ws.write(f'J{item_10}', f'', structure.alt_blue_middle)
        ws.write(f'J{item_11}', f'', structure.blue_middle)
        ws.write(f'J{item_12}', f'', structure.alt_blue_middle)
        ws.write(f'J{item_13}', f'', structure.blue_middle)
        ws.write(f'J{item_14}', f'', structure.alt_blue_middle)
        ws.write(f'J{item_15}', f'', structure.missing_cell)
        ws.write(f'J{item_16}', f'', structure.alt_blue_middle)
        ws.write(f'J{item_17}', f'', structure.blue_middle)
        ws.write(f'J{item_18}', f'', structure.alt_blue_middle)
        ws.write(f'J{item_19}', f'', structure.blue_middle)
        ws.write(f'J{item_20}', f'', structure.alt_blue_middle)
        ws.write(f'J{item_21}', f'', structure.blue_middle)
        ws.write(f'J{item_22}', f'', structure.alt_blue_middle)
        ws.write(f'J{item_23}', f'', structure.blue_middle)
        ws.write(f'J{item_24}', f'', structure.alt_blue_middle)
        ws.write(f'J{item_25}', f'', structure.blue_middle)
        ws.write(f'J{item_26}', f'', structure.alt_blue_middle)
        ws.write(f'J{item_27}', f'', structure.blue_middle)
        ws.write(f'J{item_28}', f'', structure.alt_blue_middle)
        ws.write(f'J{item_29}', f'', structure.blue_middle)
        ws.write(f'J{item_30}', f'', structure.alt_blue_middle)

        try:
            ws.write(f'J{item_3}', f'{crd_scanner.get_bios()[1]}', structure.blue_middle)
            ws.write(f'J{item_4}', f'{crd_scanner.get_bios()[1]}', structure.alt_blue_middle)
            ws.write(f'J{item_5}', f'{crd_scanner.get_bmc()[0]}', structure.blue_middle)
            ws.write(f'J{item_6}', f'{crd_scanner.get_tpm()[0]}', structure.alt_blue_middle)
            ws.write(f'J{item_7}', f'{crd_scanner.get_cpld()[0]}', structure.blue_middle)
            ws.write(f'J{item_8}', f'{crd_scanner.get_chipset()[0]}', structure.alt_blue_middle)
            ws.write(f'J{item_9}', f'', structure.missing_cell)  # Fix Processor
            ws.write(f'J{item_10}', f'{crd_scanner.get_fpga_release()[0]}', structure.alt_blue_middle)
            ws.write(f'J{item_11}', f'{crd_scanner.get_fpga_hyperblaster()[0]}', structure.blue_middle)
            ws.write(f'J{item_12}', f'{crd_scanner.get_fpga_hip()[0]}', structure.alt_blue_middle)
            ws.write(f'J{item_13}', f'{crd_scanner.get_fpga_filter()[0]}', structure.blue_middle)
            ws.write(f'J{item_14}', f'{crd_scanner.ftdi_port_list[0]}', structure.alt_blue_middle)
            ws.write(f'J{item_15}', f'{crd_scanner.ftdi_bus_list[0]}', structure.blue_middle)
            ws.write(f'J{item_16}', f'{crd_scanner.get_nic()[0]}', structure.alt_blue_middle)
            ws.write(f'J{item_17}', f'{crd_scanner.get_nic_pxe()[0]}', structure.blue_middle)
            ws.write(f'J{item_18}', f'', structure.missing_cell)
            ws.write(f'J{item_19}', f'', structure.missing_cell)
            ws.write(f'J{item_20}', f'', structure.missing_cell)
            ws.write(f'J{item_21}', f'{crd_scanner.get_nvme_pn()[0]}', structure.blue_middle)
            ws.write(f'J{item_22}', f'{crd_scanner.get_nvme_pn()[0]}', structure.alt_blue_middle)
            ws.write(f'J{item_23}', f'{crd_scanner.get_hdd_pn()[0]}', structure.blue_middle)
            ws.write(f'J{item_24}', f'{crd_scanner.get_hdd_pn()[0]}', structure.alt_blue_middle)
            ws.write(f'J{item_25}', f'', structure.missing_cell)
            ws.write(f'J{item_26}', f'', structure.missing_cell)
            ws.write(f'J{item_27}', f'', structure.missing_cell)  # Need to Fix PSU Part Number
            ws.write(f'J{item_28}', f'', structure.missing_cell)  # Need to Fix PSU Firmware
            ws.write(f'J{item_29}', f'', structure.missing_cell)  # Need to Fix Manager Switch Version
            ws.write(f'J{item_30}', f'', structure.missing_cell)  # Need to Fix BMC Version
        except IndexError:
            pass

        # Status Column
        write_request_type_status(trr_id)  # Item 2
        write_bios_version_status(trr_id)  # Item 3
        write_bios_flavor_status(trr_id)  # Item 4
        write_bmc_status(trr_id)  # Item 5
        write_cpld_status(trr_id)  # Item 6
        write_tpm_status(trr_id)  # Item 6
        write_chip_driver_status(trr_id)  # Item 6
        write_processor_status(trr_id)  # Item 7
        write_fpga_release_package_status(trr_id)  # Item 8
        write_fpga_hyperblaster_dll(trr_id)  # Item 8
        write_fpga_hip(trr_id)  # Item 8
        write_fpga_filter_status(trr_id)  # Item 8
        write_ftdi_port(trr_id)  # Item 8
        write_ftdi_bus(trr_id)  # Item 8
        write_nic_firmware_status(trr_id)  # Item 9
        write_nic_pxe_status(trr_id)  # Item 10
        write_nic_uefi_status(trr_id)  # Item 11
        write_nic_driver_status(trr_id)  # Item 12
        write_boot_drive_status(trr_id)  # Item 13
        write_nvme_part_number_status(trr_id)  # Item 14
        write_nvme_version_status(trr_id)  # Item 15
        write_hdd_part_number_status(trr_id)  # Item 16
        write_hdd_version_status(trr_id)  # Item 17
        write_dimm_part_number_status(trr_id)  # Item 18
        write_dimm_version_status(trr_id)  # Item 19
        write_psu_part_number_status(trr_id)  # Item 22
        write_psu_version_status(trr_id)  # Item 23
        write_manager_switch_status(trr_id)  # Item 24
        write_jbof_status(trr_id)  # Item 25

        start += 1
        previous += 30


def create_summary(structure, wb, ws):
    sheet_name = 'CRD vs. TRR'
    try:
        ws.write_url('G3', f"internal:'{sheet_name}'!D16", structure.white_thin_back, f'{mismatch_summary[0]}')
        ws.write_url('G4', f"internal:'{sheet_name}'!D17", structure.white_thin_back, f'{mismatch_summary[1]}')
        ws.write_url('G5', f"internal:'{sheet_name}'!D18", structure.white_thin_back, f'{mismatch_summary[2]}')
        ws.write_url('G6', f"internal:'{sheet_name}'!D19", structure.white_thin_back, f'{mismatch_summary[3]}')
        ws.write_url('G7', f"internal:'{sheet_name}'!D20", structure.white_thin_back, f'{mismatch_summary[4]}')
        ws.write_url('G8', f"internal:'{sheet_name}'!D21", structure.white_thin_back, f'{mismatch_summary[5]}')
        ws.write_url('G9', f"internal:'{sheet_name}'!D22", structure.white_thin_back, f'{mismatch_summary[6]}')
        ws.write_url('G10', f"internal:'{sheet_name}'!D23", structure.white_thin_back, f'{mismatch_summary[7]}')
        ws.write_url('G11', f"internal:'{sheet_name}'!D24", structure.white_thin_back, f'{mismatch_summary[8]}')
        ws.write_url('G12', f"internal:'{sheet_name}'!D25", structure.white_thin_back, f'{mismatch_summary[9]}')
    except IndexError:
        pass
    # ws.write('G5', f'{mismatch_summary[1]}', white_back)
    # ws.write('G6', f'{mismatch_summary[2]}', white_back)
    # ws.write('G7', f'{mismatch_summary[3]}', white_back)
    # ws.write('G8', f'{mismatch_summary[4]}', white_back)
    # ws.write('G9', f'{mismatch_summary[5]}', white_back)
    # ws.write('G10', f'{mismatch_summary[6]}', white_back)
    # ws.write('G11', f'{mismatch_summary[7]}', white_back)
    # ws.write('G12', f'{mismatch_summary[8]}', white_back)
    ws.merge_range('G1:G2', 'Summary of CRD vs. TRR', structure.teal_middle)
    ws.write('G3', f'Match = {sum(match_tally)} | Mismatch = {sum(mismatch_tally)} | '
                   f'Missing = {sum(missing_tally)}', structure.blue_middle_big)


def to_microsoft(structure, ws):
    ws.write('I3', f'{mismatch_microsoft[0]}', structure.white_thin_back)
    ws.write('I4', f'{mismatch_microsoft[1]}', structure.white_thin_back)
    ws.write('I5', f'{mismatch_microsoft[2]}', structure.white_thin_back)
    ws.write('I6', f'{mismatch_microsoft[3]}', structure.white_thin_back)
    ws.write('I7', f'{mismatch_microsoft[4]}', structure.white_thin_back)
    ws.merge_range('I1:I2', 'To Microsoft for TRR   vs   CRD', structure.teal_middle)


def set_groupings(ws):

    def create_grouping_1():
        group_start = 13
        group_end = 43
        # for trr in unique_trrs:
        #     while group_start < group_end:
        while group_start <= group_end:
            ws.set_row(group_start, None, None, {'level': 1, 'hidden': False})
            group_start += 1
        ws.outline_settings(symbols_below=False)

    def create_grouping_2():
        group_start = 44
        group_end = 74
        # for trr in unique_trrs:
        #     while group_start < group_end:
        while group_start <= group_end:
            ws.set_row(group_start, None, None, {'level': 1, 'hidden': False})
            group_start += 1
            # group_end += 27
            ws.outline_settings(symbols_below=False)

    create_grouping_1()
    create_grouping_2()


def create_sheet_5(pipe_number, host_ids, sheet_title, write_book, crd, unique_trrs):
    write_sheet = write_book.add_worksheet(sheet_title)
    structure = Structure(write_book)
    crd_scanner.create_csv()

    set_sheet_structure(pipe_number, host_ids, structure, write_sheet, sheet_title, crd, unique_trrs)
    print(f'  - Writing {sheet_title} data...')
    write_data(write_book, write_sheet, host_ids, crd, unique_trrs)

    create_graphs(write_book, write_sheet)
    create_summary(structure, write_book, write_sheet)
    to_microsoft(structure, write_sheet)
    # set_groupings(write_sheet)

    print('    * Created Sheet 2')
