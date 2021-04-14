from pipe_cleaner.src.excel_properties import requested_configuration
from pipe_cleaner.src.excel_properties import Structure
from pipe_cleaner.src.credentials import Path
from pipe_cleaner.src.data_access import request_ado
import pipe_cleaner.extra.crd_scanner as crd_scanner
from json import loads
import xlsxwriter.workbook

crd_bios: list = []
crd_bmc: list = []
crd_tpm: list = []
crd_cpld: list = []

match_tally: list = []
mismatch_tally: list = []
missing_tally: list = []

mismatch_summary: list = []
missing_summary: list = []

mismatch_microsoft: list = []

trr_components: dict = {}
crd_components: dict = {}


def create_graphs(workbook: object, worksheet: object, sheet_name: str):

    bold = workbook.add_format({'bold': 1})

    # Add the worksheet data that the charts will refer to.
    headings = ['Number', 'Tallies']
    data = [
        ['Match/Present', 'Mismatch', 'Missing'],
        [sum(match_tally), sum(mismatch_tally), sum(missing_tally)],
    ]

    worksheet.write_row('A1', headings, bold)
    worksheet.write_column('A2', data[0])
    worksheet.write_column('B2', data[1])

    chart_1 = workbook.add_chart({'type': 'bar'})
    # chart_1 = wb.add_chart({'type': 'pie'})

    # workbook.define_name(f'{sheet_name}', '=Sheet2')

    # Configure the first series.
    chart_1.add_series({
        'name': "='" + sheet_name + "'!$B$1",
        'categories': "='" + sheet_name + "'!$A$2:$A$4",
        'values': "='" + sheet_name + "'!$B$2:$B$4",
        'points': [
            {'fill': {'color': '#00B050'}},
            {'fill': {'color': '#FF0000'}},
            {'fill': {'color': '#DCAA1B'}},
        ],
    })

    # Configure a second series. Note use of alternative syntax to define ranges.
    chart_1.add_series({
        'name': [f"{sheet_name}", 0, 2],
        'categories': [f"{sheet_name}", 1, 0, 3, 0],
        'values': [f"{sheet_name}", 1, 2, 3, 2],
    })

    # Add a chart title and some axis labels.
    chart_1.set_title({'name': 'Status of CRD vs TRR'})
    # chart_1.set_x_axis({'name': 'Tally of Status'})
    # chart_1.set_y_axis({'name': 'Status'})

    # Chart Style of Graph
    chart_1.set_style(11)
    chart_1.set_legend({'none': True})

    # Size of Chart
    worksheet.insert_chart('E1', chart_1, {'x_scale': 1.185, 'y_scale': 0.84})


def set_sheet_structure(pipe_number, host_ids, structure, worksheet, sheet_title, unique_trrs):
    worksheet.set_landscape()

    worksheet.set_row(0, 12, structure.white)
    worksheet.set_row(1, 20, structure.white)
    worksheet.set_row(2, 16, structure.white)
    worksheet.set_row(3, 15, structure.white)
    worksheet.set_row(4, 15, structure.white)
    worksheet.set_row(5, 15, structure.white)
    worksheet.set_row(6, 15, structure.white)
    worksheet.set_row(7, 15, structure.white)
    worksheet.set_row(8, 15, structure.white)
    worksheet.set_row(9, 15, structure.white)
    worksheet.set_row(10, 15, structure.white)
    worksheet.set_row(11, 15, structure.white)

    start = 13
    while start < 500:
        worksheet.set_row(start, 16.5, structure.white)
        start += 1

    worksheet.set_column('A:A', 5.5, structure.white)
    worksheet.set_column('B:B', 13, structure.white)
    worksheet.set_column('C:C', 13, structure.white)
    worksheet.set_column('D:D', 24, structure.white)
    worksheet.set_column('E:E', 22, structure.white)
    worksheet.set_column('F:F', 65, structure.white)
    worksheet.set_column('G:G', 65, structure.white)
    worksheet.set_column('H:H', 3, structure.white)
    worksheet.set_column('I:I', 70, structure.white)
    worksheet.set_column('J:J', 70, structure.white)
    worksheet.set_column('K:K', 25, structure.white)
    worksheet.set_column('L:L', 25, structure.white)
    worksheet.set_column('M:M', 25, structure.white)
    worksheet.set_column('N:N', 25, structure.white)
    worksheet.set_column('O:O', 25, structure.white)
    worksheet.set_column('P:P', 25, structure.white)

    def create_grouping_1():
        group_start = 13
        group_end = 43
        # for trr in unique_trrs:
        #     while group_start < group_end:
        while group_start <= group_end:
            worksheet.set_row(group_start, None, None, {'level': 0, 'hidden': False})
            group_start += 1

    def create_grouping_2():
        group_start = 44
        group_end = 74
        # for trr in unique_trrs:
        #     while group_start < group_end:
        while group_start <= group_end:
            worksheet.set_row(group_start, None, None, {'level': 1, 'hidden': False})
            group_start += 1
            # group_end += 27

    if len(host_ids) > 1:
        f'{len(host_ids)}'
    else:
        f'{len(host_ids)}'

    worksheet.insert_image('A1', 'pipe_cleaner/img/vse_logo.png')
    worksheet.write('B5', f' Pipe Cleaner - {sheet_title}', structure.big_blue_font)
    worksheet.write('B6', f'       Kirkland Lab Site', structure.bold_italic_blue_font)
    worksheet.write('B7', f'       Pipe Name - {pipe_number}', structure.bold_italic_blue_font)
    worksheet.write('B8', f'       Total TRRs - {len(unique_trrs)}', structure.bold_italic_blue_font)
    worksheet.write('B10', f'       {Structure.date} - {Structure.time}', structure.italic_blue_font)

    worksheet.write('E2', f'', structure.big_blue_font)
    worksheet.write('F4', f'', structure.italic_blue_font)

    worksheet.write('B13', 'TRR ID', structure.teal_middle)
    worksheet.write('C13', 'Type', structure.teal_middle)
    worksheet.write('D13', 'Status', structure.teal_middle)
    worksheet.write('E13', 'Specific', structure.teal_middle)
    worksheet.write('F13', 'Additional Information', structure.teal_middle)
    worksheet.write('G13', 'General Notes', structure.teal_middle)
    worksheet.write('I13', 'Test Run Request', structure.teal_middle)
    worksheet.write('J13', 'Customer Requirements Document', structure.teal_middle)

    # Freeze Planes
    worksheet.freeze_panes(13, 4)

    # Groupings
    # create_grouping_1(unique_trrs)
    # create_grouping_2(unique_trrs)


def write_data(structure, worksheet, unique_requests):
    def write_request_type_status(trr_id):  # Item 2
        component = 'Request Type'

        check_key_terms: list = ['STORAGE',
                                 'UTILITY',
                                 '5.',
                                 '6.',
                                 '7.',
                                 '.0',
                                 '.1',
                                 '.2',
                                 '.3',
                                 '.4',
                                 '.5',
                                 '.8',
                                 'XIO',
                                 'XSTORE',
                                 'COMPUTE',
                                 'SQL',
                                 'TE',
                                 'DB',
                                 'DW',
                                 'BALANCED',
                                 'WEB',
                                 'XDIRECT',
                                 'MBX',
                                 'GP',
                                 'LOW',
                                 'MID',
                                 'HIGH',
                                 'C2010',
                                 'AMD',
                                 'OPTIMIZED',
                                 'AEP',
                                 'JBOF',
                                 'JBOD',
                                 '4TIB',
                                 '2TIB',
                                 'HM',
                                 '1U',
                                 'ANALYTICS',
                                 'SSD']

        trr_azure = target_configuration_raw
        trr_upper = str(trr_azure).upper()
        try:
            crd_azure = crd_scanner.get_azure()[0]
        except IndexError:
            pass
        crd_upper = str(crd_azure).upper()

        def check_trr_crd(terms: list) -> int:
            """
            Checks if
            :param terms:
            :return:
            """
            points: list = []
            initial: int = 0

            while initial < len(terms):
                if terms[initial] in trr_upper and terms[initial] in crd_upper:
                    add_total = + 1
                    points.append(add_total)
                initial += 1

            return sum(points)

        sum_points = check_trr_crd(check_key_terms)

        if sum_points >= 3:
            worksheet.write(f'E{item_02}', f'MATCH', structure.good_cell)
            worksheet.write(f'F{item_02}', f'Target Configuration - TRR and CRD match', structure.alt_blue_middle)
            match_tally.append(1)
        else:
            worksheet.write(f'E{item_02}', f'MISMATCH', structure.bad_cell)
            worksheet.write(f'F{item_02}', f'Target Configuration - TRR and CRD do not match',
                            structure.alt_blue_middle)
            summary = f'Mismatch = TRR {trr_id} - {parse_target_configuration(target_configuration_raw)} - {component}'
            mismatch_tally.append(1)
            mismatch_summary.append(summary)

    def contrast_row_colors(item_num: int) -> structure:
        """
        Determines if item number is odd or even for row color within excel file. Increases readability with
        contrasting row colors. Returns lighter blue if item number is odd. If not, returns dark blue color.
        :param item_num: Item number
        :return: alt_blue_middle or blue_middle
        """
        if item_num % 2 == 1:
            return structure.alt_blue_middle
        else:
            return structure.blue_middle

    def contrast_rich_normal(item_num: int) -> structure:
        """
        Determines if item number is odd or even for row color within excel file. Increases readability with
        contrasting row colors. Returns lighter blue if item number is odd. If not, returns dark blue color.
        :param item_num: Item number
        :return: alt_blue_middle or blue_middle
        """
        if item_num % 2 == 1:
            return structure.alt_blue_rich_normal
        else:
            return structure.blue_middle

    def contrast_rich_bold(item_num: int) -> structure:
        """
        Determines if item number is odd or even for row color within excel file. Increases readability with
        contrasting row colors. Returns lighter blue if item number is odd. If not, returns dark blue color.
        :param item_num: Item number
        :return: alt_blue_middle or blue_middle
        """
        if item_num % 2 == 1:
            return structure.alt_blue_rich_bold
        else:
            return structure.blue_middle

    def check_component(component: str) -> str:
        """Checks whether Server Processor is AMD or Intel for a shorter response.
        :param component:
        :return:
        """
        upper_component = component.upper()
        if 'INTEL' in upper_component:
            return 'Intel'
        elif 'AMD' in upper_component:
            return 'AMD'
        else:
            return 'None'

    def write_excel(item_num: int, request_num: int, trr_part: str, crd_part: str, component: str):
        color_of_row = contrast_row_colors(item_num)
        request_num = str(request_num)
        trr_upper = str(trr_part).upper()
        crd_upper = str(crd_part).upper()

        if trr_upper == '' or crd_upper == 'NONE':
            worksheet.write(f'D{item_num}', f'{component}', structure.neutral_cell)
            worksheet.write(f'E{item_num}', f'TRR - {trr_part}\n'
                                            f'CRD - None', color_of_row)
            missing_tally.append(1)
        elif trr_upper == crd_upper:
            worksheet.write(f'D{item_num}', f'{component}', structure.good_cell)
            worksheet.write(f'E{item_num}', f'TRR - {trr_part}\n'
                                            f'CRD - {crd_part}', color_of_row)
            match_tally.append(1)
        else:
            worksheet.write(f'D{item_num}', f'{component}', structure.bad_cell)
            worksheet.write(f'E{item_num}', f'TRR - {trr_part}\n'
                                            f'CRD - {crd_part}', color_of_row)
            summary = f'Mismatch = TRR {request_num} - {parse_target_configuration(target_configuration_raw)} - ' \
                      f'{component}'
            mismatch_summary.append(summary)
            mismatch_message = f'For {request_num} - {parse_target_configuration(target_configuration_raw)}, ' \
                               f'we are missing {component}'
            mismatch_microsoft.append(mismatch_message)
            mismatch_tally.append(1)

    def create_components_dictionary():
        trr_components.clear()

        trr_components['BIOS Version'] = f'{trr_bios_parsed.split(".")[2]}'
        crd_components['BIOS Version'] = f'{crd_scanner.get_bios()[1].split(".")[2]}'

        trr_components['BIOS Flavor'] = f'{trr_bios_parsed.split(".")[3]}'
        crd_components['BIOS Flavor'] = f'{crd_scanner.get_bios()[1].split(".")[3]}'

        trr_components['BMC Version'] = f'{trr_bmc_parsed}'
        crd_components['BMC Version'] = f'{crd_scanner.get_bmc()[0][10:-3]}'

        trr_components['TPM Version'] = f'{trr_tpm_parsed}'
        crd_components['TPM Version'] = f'{crd_scanner.get_tpm()[0][:4:]}'

        trr_components['CPLD Version'] = f'{cpld_raw}'
        crd_components['CPLD Version'] = f'{crd_scanner.get_cpld()[0].replace("V", "")[:2:]}'

        trr_components['Chipset Driver'] = f'{chipset_raw}'
        crd_components['Chipset Driver'] = f'{crd_scanner.get_chipset()[0]}'

        trr_components['Server Processor'] = f'{processor_raw}'
        crd_components['Server Processor'] = f'{processor_raw}'

        trr_components['FPGA Release Version'] = f'{fpga_release_raw}'
        crd_components['FPGA Release Version'] = f'{crd_scanner.get_fpga_release()[0]}'

        trr_components['FPGA Hyperblaster DLL'] = f'{fpga_hyperblaster_raw}'
        crd_components['FPGA Hyperblaster DLL'] = f'{crd_scanner.get_fpga_hyperblaster()[0]}'

        trr_components['FPGA HIP'] = f'{fpga_hip_raw}'
        crd_components['FPGA HIP'] = f'{crd_scanner.get_fpga_hip()[0]}'

        trr_components['FPGA HIP'] = f'{fpga_filter_raw}'
        crd_components['FPGA HIP'] = f'{crd_scanner.get_fpga_filter()[0]}'

        trr_components['FTDI Port'] = f'{raw_ftdi_port}'
        crd_components['FTDI Port'] = f'{crd_scanner.get_ftdi_port()[0]}'

        trr_components['FTDI Filter'] = f'{raw_ftdi_port}'
        crd_components['FTDI Filter'] = f'{crd_scanner.get_ftdi_bus()[0]}'

        trr_components['NIC Firmware'] = f'{nic_firmware_raw}'
        crd_components['NIC Firmware'] = f'{crd_scanner.get_nic()[0]}'

        trr_components['NIC PXE'] = f'{nic_pxe_raw}'
        crd_components['NIC PXE'] = f'{crd_scanner.get_nic_pxe()[0]}'

        trr_components['NIC UEFI'] = f'{nic_uefi_raw}'
        crd_components['NIC UEFI'] = f'{crd_scanner.get_nic_uefi()[0]}'

        trr_components['NIC Driver'] = f'{nic_driver_raw}'
        crd_components['NIC Driver'] = f'{crd_scanner.get_nic_pxe()[0]}'  # TODO NIC Driver

        trr_components['Boot Drive'] = f'{boot_driver_raw}'
        crd_components['Boot Drive'] = f'{crd_scanner.get_nic_uefi()[0]}'  # TODO Present

    def write_status(request_number: int, item_num: int, component: str):
        trr_item = trr_components.get(component)
        crd_item = crd_components.get(component)

        write_excel(item_num, request_number, trr_item, crd_item, component)

    def check_bios_bmc(raw_str: str, item_num: str):
        """ Check if BIOS or BMC. If not then check other components.
        :param item_num:
        :param raw_str:
        :return:
        """
        if '.BS.' in raw_str:
            worksheet.write_rich_string(f'F{item_num}',
                                        contrast_rich_normal(int(item_num)), f'{raw_str[9::]}',
                                        structure.bold_text, f'.{raw_str.split(".")[2]}.')
            # contrast_rich_normal(int(item_num)), f'{raw_str[-4::]}')
        # elif '.BC.' in raw_str:
        #     return 'BMC'

    def additional_information_stack(item_num: int, string_raw: str, trr_raw: str, crd_raw: str) -> None:
        """
        Writes excel for additional information column. Simply prints out trr and crd stacked on top of each other if
        there is no string raw
        :param item_num: abstract number for aligning rows up
        :param string_raw: just to print directly
        :param trr_raw: takes component's raw information from TRR
        :param crd_raw: takes component's raw information from CRD
        :return: None
        """
        item_num_str = str(item_num)
        # if not string_raw:
        #     worksheet.write(f'F{item_num_str}', f'{trr_raw}\n'
        #                                         f'{crd_raw}', contrast_row_colors(item_num))
        # else:
        #     worksheet.write(f'F{item_num_str}', f'{string_raw}', contrast_row_colors(item_num))

        check_bios_bmc(trr_raw, item_num_str)

    def write_nvme_part_number_status(trr_id):  # Item 16 NEED TO FIX
        component = 'NVMe Part Number'
        worksheet.write(f'E{item_21}', f'WAIVED', structure.blue_middle)
        worksheet.write(f'F{item_21}', f'NVMe Part Number to be worked on', structure.blue_middle)
        # trr_item = nvme_raw_trr
        # trr_upper = str(trr_item).upper()
        # crd_item = crd_scanner.get_nvme_pn()[0]
        # crd_upper = str(crd_item).upper()
        #
        # if trr_upper == '' or crd_upper == 'NONE':
        #     worksheet.write(f'E{item_21}', f'MISSING', structure.neutral_cell)
        #     worksheet.write(f'F{item_21}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)
        #
        #     missing_tally.append(1)
        # elif trr_upper == crd_upper:
        #     worksheet.write(f'E{item_21}', f'MATCH', structure.good_cell)
        #     worksheet.write(f'F{item_21}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)
        #
        #     match_tally.append(1)
        # else:
        #     worksheet.write(f'E{item_21}', f'MISMATCH', structure.bad_cell)
        #     worksheet.write(f'F{item_21}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)
        #
        #     mismatch_tally.append(1)

    def write_nvme_version_status(trr_id):  # Item 17 NEED TO FIX
        component = 'NVMe Version'
        worksheet.write(f'E{item_22}', f'WAIVED', structure.alt_blue_middle)
        worksheet.write(f'F{item_22}', f'NVMe Version to be worked on', structure.alt_blue_middle)
        # trr_item = nvme_raw_trr
        # trr_upper = str(trr_item).upper()
        # crd_item = crd_scanner.get_nvme_pn()[0]
        # crd_upper = str(crd_item).upper()
        #
        # if trr_upper == '' or crd_upper == 'NONE':
        #     worksheet.write(f'E{item_22}', f'MISSING', structure.neutral_cell)
        #     worksheet.write(f'F{item_22}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.blue_middle)
        # elif trr_upper == crd_upper:
        #     worksheet.write(f'E{item_22}', f'MATCH', structure.good_cell)
        #     worksheet.write(f'F{item_22}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.blue_middle)
        # else:
        #     worksheet.write(f'E{item_22}', f'MISMATCH', structure.bad_cell)
        #     worksheet.write(f'F{item_22}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.blue_middle)

    def write_hdd_part_number_status(trr_id):  # Item 18 NEED TO FIX
        component = 'HDD Part Number'
        worksheet.write(f'E{item_23}', f'WAIVED', structure.blue_middle)
        worksheet.write(f'F{item_23}', f'HDD Part Number to be worked on', structure.blue_middle)
        # trr_item = hdd_raw_trr
        # trr_upper = str(trr_item).upper()
        # crd_item = crd_scanner.get_nvme_pn()[0]
        # crd_upper = str(crd_item).upper()
        #
        # if trr_upper == '' or crd_upper == 'NONE':
        #     worksheet.write(f'E{item_23}', f'MISSING', structure.neutral_cell)
        #     worksheet.write(f'F{item_23}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.blue_middle)
        # elif trr_upper == crd_upper:
        #     worksheet.write(f'E{item_23}', f'MATCH', structure.good_cell)
        #     worksheet.write(f'F{item_23}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.blue_middle)
        # else:
        #     worksheet.write(f'E{item_23}', f'MISMATCH', structure.bad_cell)
        #     worksheet.write(f'F{item_23}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.blue_middle)

    def write_hdd_version_status(trr_id):  # Item 19 NEED TO FIX
        component = 'HDD Version'
        worksheet.write(f'E{item_24}', f'WAIVED', structure.alt_blue_middle)
        worksheet.write(f'F{item_24}', f'HDD Version to be worked on', structure.alt_blue_middle)
        # trr_item = hdd_raw_trr
        # trr_upper = str(trr_item).upper()
        # crd_item = crd_scanner.get_nvme_pn()[0]
        # crd_upper = str(crd_item).upper()
        #
        # if trr_upper == '' or crd_upper == 'NONE':
        #     worksheet.write(f'E{item_24}', f'MISSING', structure.neutral_cell)
        #     worksheet.write(f'F{item_24}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)
        # elif trr_upper == crd_upper:
        #     worksheet.write(f'E{item_24}', f'MATCH', structure.good_cell)
        #     worksheet.write(f'F{item_24}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)
        # else:
        #     worksheet.write(f'E{item_24}', f'MISMATCH', structure.bad_cell)
        #     worksheet.write(f'F{item_24}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)

    def write_dimm_part_number_status(trr_id):  # Item 20 NEED TO FIX
        component = 'DIMM Part Number'
        worksheet.write(f'E{item_25}', f'WAIVED', structure.blue_middle)
        worksheet.write(f'F{item_25}', f'DIMM Part Number to be worked on', structure.blue_middle)
        # trr_item = dimm_raw_trr
        # trr_upper = str(trr_item).upper()
        # crd_item = crd_scanner.get_nvme_pn()[0]
        # crd_upper = str(crd_item).upper()
        #
        # if trr_upper == '' or crd_upper == 'NONE':
        #     worksheet.write(f'E{item_25}', f'MISSING', structure.neutral_cell)
        #     worksheet.write(f'F{item_25}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.blue_middle)
        # elif trr_upper == crd_upper:
        #     worksheet.write(f'E{item_25}', f'MATCH', structure.good_cell)
        #     worksheet.write(f'F{item_25}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.blue_middle)
        # else:
        #     worksheet.write(f'E{item_25}', f'MISMATCH', structure.bad_cell)
        #     worksheet.write(f'F{item_25}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.blue_middle)

    def write_dimm_version_status(trr_id):  # Item 21 NEED TO FIX
        component = 'DIMM Version'
        worksheet.write(f'E{item_26}', f'WAIVED', structure.alt_blue_middle)
        worksheet.write(f'F{item_26}', f'DIMM Version to be worked on', structure.alt_blue_middle)
        # trr_item = dimm_raw_trr
        # trr_upper = str(trr_item).upper()
        # crd_item = crd_scanner.get_nvme_pn()[0]
        # crd_upper = str(crd_item).upper()
        #
        # if trr_upper == '' or crd_upper == 'NONE':
        #     worksheet.write(f'E{item_26}', f'MISSING', structure.neutral_cell)
        #     worksheet.write(f'F{item_26}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)
        # elif trr_upper == crd_upper:
        #     worksheet.write(f'E{item_26}', f'MATCH', structure.good_cell)
        #     worksheet.write(f'F{item_26}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)
        # else:
        #     worksheet.write(f'E{item_26}', f'MISMATCH', structure.bad_cell)
        #     worksheet.write(f'F{item_26}', f'TRR - {trr_bmc}   vs   CRD - {crd_item}', structure.alt_blue_middle)

    def write_psu_part_number_status(trr_id):  # Item 22 NEED TO FIX
        component = 'PSU Part Number'
        worksheet.write(f'E{item_27}', f'WAIVED', structure.blue_middle)
        worksheet.write(f'F{item_27}', f'PSU Part Number to be worked on', structure.blue_middle)
        # trr_item = raw_psu_part_number
        # trr_upper = str(trr_item).upper()
        # crd_item = crd_scanner.get_psu()[0][:28:]
        # crd_upper = str(crd_item).upper()
        #
        # if trr_upper == '' or crd_upper == 'NONE':
        #     worksheet.write(f'E{item_27}', f'MISSING', structure.neutral_cell)
        #     worksheet.write(f'F{item_27}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)
        # elif trr_upper == crd_upper:
        #     worksheet.write(f'E{item_27}', f'MATCH', structure.good_cell)
        #     worksheet.write(f'F{item_27}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)
        # else:
        #     worksheet.write(f'E{item_27}', f'MISMATCH', structure.bad_cell)
        #     worksheet.write(f'F{item_27}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)

    def write_psu_version_status(trr_id):  # Item 23 NEED TO FIX
        component = 'PSU Version'
        worksheet.write(f'E{item_28}', f'WAIVED', structure.alt_blue_middle)
        worksheet.write(f'F{item_28}', f'PSU Version to be worked on', structure.alt_blue_middle)
        # trr_item = raw_psu_version
        # trr_upper = str(trr_item).upper()
        # crd_item = crd_scanner.get_nvme_pn()[0]
        # crd_item = crd_item[28:][:8:]
        # crd_upper = str(crd_item).upper()
        #
        # if trr_upper == '' or crd_upper == 'NONE':
        #     worksheet.write(f'E{item_28}', f'MISSING', structure.neutral_cell)
        #     worksheet.write(f'F{item_28}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.blue_middle)
        # elif trr_upper == crd_upper:
        #     worksheet.write(f'E{item_28}', f'MATCH', structure.good_cell)
        #     worksheet.write(f'F{item_28}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.blue_middle)
        # else:
        #     worksheet.write(f'E{item_28}', f'MISMATCH', structure.bad_cell)
        #     worksheet.write(f'F{item_28}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.blue_middle)

    def write_manager_switch_status(trr_id):  # Item 24 NEED TO FIX
        component = 'Manager Switch'
        worksheet.write(f'E{item_29}', f'WAIVED', structure.blue_middle)
        worksheet.write(f'F{item_29}', f'Manager Switch to be worked on', structure.blue_middle)
        # trr_item = bios_raw_trr
        # trr_upper = str(trr_item).upper()
        # crd_item = crd_scanner.get_nvme_pn()[0]
        # crd_upper = str(crd_item).upper()
        #
        # if trr_upper == '' or crd_upper == 'NONE':
        #     worksheet.write(f'E{item_29}', f'MISSING', structure.neutral_cell)
        #     worksheet.write(f'F{item_29}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)
        # elif trr_upper == crd_upper:
        #     worksheet.write(f'E{item_29}', f'MATCH', structure.good_cell)
        #     worksheet.write(f'F{item_29}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)
        # else:
        #     worksheet.write(f'E{item_29}', f'MISMATCH', structure.bad_cell)
        #     worksheet.write(f'F{item_29}', f'TRR - {trr_item}   vs   CRD - {crd_item}', structure.alt_blue_middle)

    def write_jbof_status(trr_id):  # Item 25 NEED TO FIX
        component = 'JBOF BMC Version'
        trr_item = jbof_bmc_raw

        trr_upper = str(trr_item).upper()
        crd_item = crd_scanner.get_nvme_pn()[0]
        crd_upper = str(crd_item).upper()

        if trr_upper == '' or trr_upper == 'NONE':
            worksheet.write(f'E{item_30}', f'NOT PRESENT', structure.neutral_cell)
            worksheet.write(f'F{item_30}', f'TRR - {trr_item}', structure.alt_blue_middle)
            missing_tally.append(1)
        elif trr_upper == crd_upper:
            worksheet.write(f'E{item_30}', f'MATCH', structure.good_cell)
            worksheet.write(f'F{item_30}', f'TRR - {trr_item}', structure.alt_blue_middle)
            match_tally.append(1)
        else:
            worksheet.write(f'E{item_30}', f'NOT PRESENT', structure.bad_cell)
            worksheet.write(f'F{item_30}', f'TRR - {trr_item}', structure.alt_blue_middle)
            mismatch_tally.append(1)

    def parse_target_configuration(raw_target: str) -> str:
        """
        Parse Target Configuration from TRR within ADO ie. XIO Server, Storage Server
        :param raw_target: Raw target configuration grabbed from TRR table within ADO
        :return: Parsed Target Configuration
        """
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

    def extract_request(request_type: str) -> str:
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
        if request_type_raw != '':
            worksheet.write(f'D{item_01}', f'{extract_request(request_type_raw)} Test', structure.good_cell)
            match_tally.append(1)
        else:
            worksheet.write(f'D{item_01}', f'MISSING', structure.neutral_cell)
            mismatch_tally.append(1)

    start: int = 0
    previous: int = 0

    while start < len(unique_requests):
        total: int = start + previous

        # Gets unique trr id from input file then requests one of the unique TRR IDs
        trr_id: int = unique_requests[start]
        request_ado(trr_id)

        with open(f'{Path.info}{str(trr_id)}/final.json') as file:
            trr = loads(file.read())

        target_configuration_raw = requested_configuration(trr, 'TARGET', 'CONFIGURATION', 'CONFIGURATION')
        bios_raw = requested_configuration(trr, 'SERVER', 'BI', 'OS')
        bmc_raw = requested_configuration(trr, 'SERVER', 'BMC', 'BMC')
        tpm_raw = requested_configuration(trr, 'SERVER', 'TPM', 'TPM')
        cpld_raw = requested_configuration(trr, 'SERVER', 'SERVER CPLD', 'CPLD')
        chipset_raw = requested_configuration(trr, 'CHIPSET', 'CHIPSET', 'DRIVER')
        processor_raw = requested_configuration(trr, 'PROCESSORS', 'PROCESSORS', 'PROCESSORS')
        fpga_release_raw = requested_configuration(trr, 'FPGA', 'RELEASE', 'PACKAGE')
        fpga_hyperblaster_raw = requested_configuration(trr, 'SERVER', 'HYPERBLASTER', 'DRIVER')
        fpga_hip_raw = requested_configuration(trr, 'SERVER', 'FPGA', 'HIP')
        fpga_filter_raw = requested_configuration(trr, 'SERVER', 'FPGA', 'FILTER')
        raw_ftdi_port = requested_configuration(trr, 'SERVER', 'FTDI', 'PORT')
        ftdi_bus_raw = requested_configuration(trr, 'SERVER', 'FTDI', 'BUS')
        nic_firmware_raw = requested_configuration(trr, 'NIC', 'FIRMWARE', 'FIRMWARE')
        nic_pxe_raw = requested_configuration(trr, 'NIC', 'PXE', 'PXE')
        nic_uefi_raw = requested_configuration(trr, 'NIC', 'UEFI', 'UEFI')
        nic_driver_raw = requested_configuration(trr, 'NIC', 'DRIVER', 'DRIVER')
        nvme_raw = requested_configuration(trr, 'QCL', 'NVME', 'NVME')
        hdd_raw = requested_configuration(trr, 'QCL', 'HDD', 'HDD')
        request_type_raw = requested_configuration(trr, 'REQUEST', 'TYPE', 'TYPE')
        dimm_raw = requested_configuration(trr, 'DIMM', 'DIMM', '1')
        psu_pn_raw = requested_configuration(trr, 'PSU', 'PSU', 'PN')
        psu_firmware_raw = requested_configuration(trr, 'PSU', 'PSU', 'FIRMWARE')
        boot_driver_raw = requested_configuration(trr, 'BOOT', 'BOOT', 'DRIVE')
        psu_version_raw = requested_configuration(trr, 'PSU', 'PSU', 'FIRMWARE')
        jbof_bmc_raw = requested_configuration(trr, 'JBOF', 'JBOF', 'JBOF')

        trr_bios_parsed = bios_raw
        trr_bmc_parsed = bmc_raw.split('.')[2][-3::]
        trr_tpm_parsed = tpm_raw.replace('V', '')[:4:]

        item_01: int = Structure.initial + total + 2  # Request Type
        item_02: int = Structure.initial + total + 3  # Target Configuration
        item_03: int = Structure.initial + total + 4  # BIOS Version
        item_04: int = Structure.initial + total + 5  # BIOS Flavor
        item_05: int = Structure.initial + total + 6  # BMC
        item_06: int = Structure.initial + total + 7  # TPM
        item_07: int = Structure.initial + total + 8  # CPLD
        item_08: int = Structure.initial + total + 9  # Chipset Driver
        item_09: int = Structure.initial + total + 10  # Server Processor
        item_10: int = Structure.initial + total + 11  # FPGA Release Package
        item_11: int = Structure.initial + total + 12  # FPGA Hyperblaster DLL
        item_12: int = Structure.initial + total + 13  # FPGA HIP Driver
        item_13: int = Structure.initial + total + 14  # FPGA Filter Driver
        item_14: int = Structure.initial + total + 15  # FTDI Port Driver
        item_15: int = Structure.initial + total + 16  # FTDI Bus Driver
        item_16: int = Structure.initial + total + 17  # NIC Firmware
        item_17: int = Structure.initial + total + 18  # NIC PXE
        item_18: int = Structure.initial + total + 19  # NIC UEFI
        item_19: int = Structure.initial + total + 20  # NIC Driver
        item_20: int = Structure.initial + total + 21  # Boot Drive
        item_21: int = Structure.initial + total + 22  # NVMe Part Number
        item_22: int = Structure.initial + total + 23  # NVMe Version
        item_23: int = Structure.initial + total + 24  # HDD Part Number
        item_24: int = Structure.initial + total + 25  # HDD Version
        item_25: int = Structure.initial + total + 26  # DIMM Part Number
        item_26: int = Structure.initial + total + 27  # DIMM Version
        item_27: int = Structure.initial + total + 28  # PSU Part Number
        item_28: int = Structure.initial + total + 29  # PSU Firmware
        item_29: int = Structure.initial + total + 30  # Manager Switch Firmware
        item_30: int = Structure.initial + total + 31  # JBOF - BMC

        # TRR ID Column
        worksheet.merge_range(f'B{item_01}:B{item_30}', f'{int(trr_id)}', structure.merge_format)

        # Type Column
        worksheet.write(f'C{item_01}', f'Request', structure.blue_middle)
        worksheet.write(f'C{item_30}', f'JBOF/F2010', structure.blue_middle)
        worksheet.merge_range(f'C{item_02}:C{item_30}', f'{parse_target_configuration(target_configuration_raw)}',
                              structure.merge_format)

        worksheet.write(f'D{item_01}', f'{extract_request(request_type_raw)} Test', structure.blue_middle)
        worksheet.write(f'D{item_02}', f'Target Configuration', structure.alt_blue_middle)

        # Additional Information Column
        check_request_type()
        worksheet.write(f'E{item_06}', f'WAIVED', structure.alt_blue_middle)

        additional_information_stack(item_01, 'Checks if Request Type is present within TRR', '', '')
        worksheet.write(f'F{item_02}', f'', structure.missing_cell)  # TODO
        additional_information_stack(item_03, '', bios_raw, crd_scanner.get_bios()[1])
        additional_information_stack(item_04, '', bios_raw, crd_scanner.get_bios()[1])
        worksheet.write(f'F{item_05}', f'', structure.missing_cell)
        worksheet.write(f'F{item_06}', f'Do not update firmware, might brick motherboard', structure.alt_blue_middle)
        worksheet.write(f'F{item_07}', f'', structure.missing_cell)
        worksheet.write(f'F{item_08}', f'', structure.missing_cell)
        worksheet.write(f'F{item_09}', f'', structure.missing_cell)
        worksheet.write(f'F{item_10}', f'', structure.missing_cell)
        worksheet.write(f'F{item_11}', f'', structure.missing_cell)
        worksheet.write(f'F{item_12}', f'', structure.missing_cell)
        worksheet.write(f'F{item_13}', f'', structure.missing_cell)
        worksheet.write(f'F{item_14}', f'', structure.missing_cell)
        worksheet.write(f'F{item_15}', f'', structure.missing_cell)
        worksheet.write(f'F{item_16}', f'', structure.missing_cell)
        worksheet.write(f'F{item_17}', f'', structure.missing_cell)
        worksheet.write(f'F{item_18}', f'', structure.missing_cell)
        worksheet.write(f'F{item_19}', f'', structure.missing_cell)
        worksheet.write(f'F{item_20}', f'', structure.missing_cell)
        worksheet.write(f'F{item_21}', f'', structure.missing_cell)
        worksheet.write(f'F{item_22}', f'', structure.missing_cell)
        worksheet.write(f'F{item_23}', f'', structure.missing_cell)
        worksheet.write(f'F{item_24}', f'', structure.missing_cell)
        worksheet.write(f'F{item_25}', f'', structure.missing_cell)
        worksheet.write(f'F{item_26}', f'', structure.missing_cell)
        worksheet.write(f'F{item_27}', f'', structure.missing_cell)
        worksheet.write(f'F{item_28}', f'', structure.missing_cell)
        worksheet.write(f'F{item_29}', f'', structure.missing_cell)
        worksheet.write(f'F{item_30}', f'Checks for JBOF/F2010 is in TRR', structure.blue_middle)

        # General Notes
        worksheet.write(f'G{item_01}', f'Request Types only show up in TRRs, not CRDs', structure.blue_middle)
        worksheet.write(f'G{item_02}', f'', structure.missing_cell)
        worksheet.write(f'G{item_03}', f'', structure.missing_cell)
        worksheet.write(f'G{item_04}', f'', structure.missing_cell)
        worksheet.write(f'G{item_05}', f'Use BMC 4.60 or higher for Gen 6', structure.blue_middle)
        worksheet.write(f'G{item_06}', f'Do not update, might brick motherboard', structure.alt_blue_middle)
        worksheet.write(f'G{item_07}', f'', structure.missing_cell)
        worksheet.write(f'G{item_08}', f'', structure.missing_cell)
        worksheet.write(f'G{item_09}', f'Only Available in TRR, Comes from SKUDOC', structure.blue_middle)
        worksheet.write(f'G{item_10}', f'', structure.missing_cell)
        worksheet.write(f'G{item_11}', f'', structure.missing_cell)
        worksheet.write(f'G{item_12}', f'', structure.missing_cell)
        worksheet.write(f'G{item_13}', f'', structure.missing_cell)
        worksheet.write(f'G{item_14}', f'', structure.missing_cell)
        worksheet.write(f'G{item_15}', f'', structure.missing_cell)
        worksheet.write(f'G{item_16}', f'', structure.missing_cell)
        worksheet.write(f'G{item_17}', f'', structure.missing_cell)
        worksheet.write(f'G{item_18}', f'', structure.missing_cell)
        worksheet.write(f'G{item_19}', f'', structure.missing_cell)
        worksheet.write(f'G{item_20}', f'Only Available in TRR, Comes from TRR Only', structure.alt_blue_middle)
        worksheet.write(f'G{item_21}', f'', structure.missing_cell)
        worksheet.write(f'G{item_22}', f'', structure.missing_cell)
        worksheet.write(f'G{item_23}', f'', structure.missing_cell)
        worksheet.write(f'G{item_24}', f'', structure.missing_cell)
        worksheet.write(f'G{item_25}', f'', structure.missing_cell)
        worksheet.write(f'G{item_26}', f'', structure.missing_cell)
        worksheet.write(f'G{item_27}', f'', structure.missing_cell)
        worksheet.write(f'G{item_28}', f'', structure.missing_cell)
        worksheet.write(f'G{item_29}', f'', structure.missing_cell)
        worksheet.write(f'G{item_30}', f'Need to check if JBOF or F2010 are in TRRs', structure.alt_blue_middle)

        worksheet.write_comment(f'G{item_05}', f'Make sure to use BMC 4.60 or higher for all Intel-Based Gen 6 WCS, '
                                               f'including xStore, xDirect and XIO Storage - MSFT, 8/3/2020',
                                {'height': 200})
        worksheet.write_comment(f'G{item_06}',
                                f'DO NOT attempt to update the TPM firmware. This is very likely to brick the '
                                f'motherboard and should not be attempted without specific instructions.'
                                f' - Eric Johnson, 5/14/2020', {'height': 200})

        def return_trr_raw(item, raw):
            def odd_or_even(position):
                if position % 2 == 0:
                    return 'EVEN'
                else:
                    return 'ODD'

            if raw == '' or raw == None:
                worksheet.write(f'I{item}', f'', structure.missing_cell)
            elif 'EVEN' == odd_or_even(item):
                worksheet.write(f'I{item}', f'{raw}', structure.blue_middle)
            else:
                worksheet.write(f'I{item}', f'{raw}', structure.alt_blue_middle)

        # Test Run Request Column
        return_trr_raw(item_01, request_type_raw)
        return_trr_raw(item_02, target_configuration_raw)
        return_trr_raw(item_03, bios_raw)
        return_trr_raw(item_04, bios_raw)
        return_trr_raw(item_05, bmc_raw)
        return_trr_raw(item_06, tpm_raw)
        return_trr_raw(item_07, cpld_raw)
        return_trr_raw(item_08, chipset_raw)
        return_trr_raw(item_09, processor_raw)
        return_trr_raw(item_10, fpga_release_raw)
        return_trr_raw(item_11, fpga_hyperblaster_raw)
        return_trr_raw(item_12, fpga_hip_raw)
        return_trr_raw(item_13, fpga_filter_raw)
        return_trr_raw(item_14, raw_ftdi_port)
        return_trr_raw(item_15, ftdi_bus_raw)
        return_trr_raw(item_16, nic_firmware_raw)
        return_trr_raw(item_17, nic_pxe_raw)
        return_trr_raw(item_18, nic_uefi_raw)
        return_trr_raw(item_19, nic_driver_raw)
        return_trr_raw(item_20, boot_driver_raw)
        return_trr_raw(item_21, nvme_raw)
        return_trr_raw(item_22, hdd_raw)
        return_trr_raw(item_23, hdd_raw)
        return_trr_raw(item_24, dimm_raw)
        return_trr_raw(item_25, dimm_raw)
        return_trr_raw(item_26, dimm_raw)
        return_trr_raw(item_27, psu_pn_raw)
        return_trr_raw(item_28, psu_firmware_raw)
        return_trr_raw(item_29, psu_firmware_raw)
        return_trr_raw(item_30, 'JBOF')

        # Customer Requirements Document Column
        worksheet.write(f'J{item_01}', f'', structure.missing_cell)
        worksheet.write(f'J{item_02}', f'{crd_scanner.get_azure()[0]}', structure.alt_blue_middle)
        worksheet.write(f'J{item_03}', f'', structure.blue_middle)
        worksheet.write(f'J{item_04}', f'', structure.alt_blue_middle)
        worksheet.write(f'J{item_05}', f'', structure.blue_middle)
        worksheet.write(f'J{item_06}', f'', structure.alt_blue_middle)
        worksheet.write(f'J{item_07}', f'', structure.blue_middle)
        worksheet.write(f'J{item_08}', f'', structure.alt_blue_middle)
        worksheet.write(f'J{item_09}', f'', structure.blue_middle)
        worksheet.write(f'J{item_10}', f'', structure.alt_blue_middle)
        worksheet.write(f'J{item_11}', f'', structure.blue_middle)
        worksheet.write(f'J{item_12}', f'', structure.alt_blue_middle)
        worksheet.write(f'J{item_13}', f'', structure.blue_middle)
        worksheet.write(f'J{item_14}', f'', structure.alt_blue_middle)
        worksheet.write(f'J{item_15}', f'', structure.missing_cell)
        worksheet.write(f'J{item_16}', f'', structure.alt_blue_middle)
        worksheet.write(f'J{item_17}', f'', structure.blue_middle)
        worksheet.write(f'J{item_18}', f'', structure.alt_blue_middle)
        worksheet.write(f'J{item_19}', f'', structure.blue_middle)
        worksheet.write(f'J{item_20}', f'', structure.alt_blue_middle)
        worksheet.write(f'J{item_21}', f'', structure.blue_middle)
        worksheet.write(f'J{item_22}', f'', structure.alt_blue_middle)
        worksheet.write(f'J{item_23}', f'', structure.blue_middle)
        worksheet.write(f'J{item_24}', f'', structure.alt_blue_middle)
        worksheet.write(f'J{item_25}', f'', structure.blue_middle)
        worksheet.write(f'J{item_26}', f'', structure.alt_blue_middle)
        worksheet.write(f'J{item_27}', f'', structure.blue_middle)
        worksheet.write(f'J{item_28}', f'', structure.alt_blue_middle)
        worksheet.write(f'J{item_29}', f'', structure.blue_middle)
        worksheet.write(f'J{item_30}', f'', structure.alt_blue_middle)

        try:
            worksheet.write(f'J{item_03}', f'{crd_scanner.get_bios()[1]}', structure.blue_middle)
            worksheet.write(f'J{item_04}', f'{crd_scanner.get_bios()[1]}', structure.alt_blue_middle)
            worksheet.write(f'J{item_05}', f'{crd_scanner.get_bmc()[0]}', structure.blue_middle)
            worksheet.write(f'J{item_06}', f'{crd_scanner.get_tpm()[0]}', structure.alt_blue_middle)
            worksheet.write(f'J{item_07}', f'{crd_scanner.get_cpld()[0]}', structure.blue_middle)
            worksheet.write(f'J{item_08}', f'{crd_scanner.get_chipset()[0]}', structure.alt_blue_middle)
            worksheet.write(f'J{item_09}', f'', structure.missing_cell)  # Fix Processor
            worksheet.write(f'J{item_10}', f'{crd_scanner.get_fpga_release()[0]}', structure.alt_blue_middle)
            worksheet.write(f'J{item_11}', f'{crd_scanner.get_fpga_hyperblaster()[0]}', structure.blue_middle)
            worksheet.write(f'J{item_12}', f'{crd_scanner.get_fpga_hip()[0]}', structure.alt_blue_middle)
            worksheet.write(f'J{item_13}', f'{crd_scanner.get_fpga_filter()[0]}', structure.blue_middle)
            worksheet.write(f'J{item_14}', f'{crd_scanner.ftdi_port_list[0]}', structure.alt_blue_middle)
            worksheet.write(f'J{item_15}', f'{crd_scanner.ftdi_bus_list[0]}', structure.blue_middle)
            worksheet.write(f'J{item_16}', f'{crd_scanner.get_nic()[0]}', structure.alt_blue_middle)
            worksheet.write(f'J{item_17}', f'{crd_scanner.get_nic_pxe()[0]}', structure.blue_middle)
            worksheet.write(f'J{item_18}', f'', structure.missing_cell)
            worksheet.write(f'J{item_19}', f'', structure.missing_cell)
            worksheet.write(f'J{item_20}', f'', structure.missing_cell)
            worksheet.write(f'J{item_21}', f'{crd_scanner.get_nvme_pn()[0]}', structure.blue_middle)
            worksheet.write(f'J{item_22}', f'{crd_scanner.get_nvme_pn()[0]}', structure.alt_blue_middle)
            worksheet.write(f'J{item_23}', f'{crd_scanner.get_hdd_pn()[0]}', structure.blue_middle)
            worksheet.write(f'J{item_24}', f'{crd_scanner.get_hdd_pn()[0]}', structure.alt_blue_middle)
            worksheet.write(f'J{item_25}', f'', structure.missing_cell)
            worksheet.write(f'J{item_26}', f'', structure.missing_cell)
            worksheet.write(f'J{item_27}', f'', structure.missing_cell)  # Need to Fix PSU Part Number
            worksheet.write(f'J{item_28}', f'', structure.missing_cell)  # Need to Fix PSU Firmware
            worksheet.write(f'J{item_29}', f'', structure.missing_cell)  # Need to Fix Manager Switch Version
            worksheet.write(f'J{item_30}', f'', structure.missing_cell)  # Need to Fix BMC Version
        except IndexError:
            pass

        create_components_dictionary()

        # Status Column
        write_request_type_status(trr_id)
        write_status(trr_id, item_03, 'BIOS Version')
        write_status(trr_id, item_04, 'BIOS Flavor')
        write_status(trr_id, item_05, 'BMC Version')
        write_status(trr_id, item_06, 'TPM Version')
        write_status(trr_id, item_07, 'CPLD Version')
        write_status(trr_id, item_08, 'Chipset Driver')
        write_status(trr_id, item_09, 'Server Processor')
        write_status(trr_id, item_10, 'FPGA Release Version')
        write_status(trr_id, item_11, 'FPGA Hyperblaster DLL')
        write_status(trr_id, item_12, 'FPGA HIP')
        write_status(trr_id, item_13, 'FTDI Port')
        write_status(trr_id, item_14, 'FTDI Filter')
        write_status(trr_id, item_15, 'NIC Firmware')
        write_status(trr_id, item_16, 'NIC PXE')
        write_status(trr_id, item_17, 'NIC UEFI')
        write_status(trr_id, item_18, 'NIC Driver')
        write_status(trr_id, item_19, 'Boot Drive')
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


def create_summary(structure: object, worksheet: object, sheet_name: str):
    try:
        worksheet.write_url('G3', f"internal:'{sheet_name}'!D16", structure.white_thin_back, f'{mismatch_summary[0]}')
        worksheet.write_url('G4', f"internal:'{sheet_name}'!D17", structure.white_thin_back, f'{mismatch_summary[1]}')
        worksheet.write_url('G5', f"internal:'{sheet_name}'!D18", structure.white_thin_back, f'{mismatch_summary[2]}')
        worksheet.write_url('G6', f"internal:'{sheet_name}'!D19", structure.white_thin_back, f'{mismatch_summary[3]}')
        worksheet.write_url('G7', f"internal:'{sheet_name}'!D20", structure.white_thin_back, f'{mismatch_summary[4]}')
        worksheet.write_url('G8', f"internal:'{sheet_name}'!D21", structure.white_thin_back, f'{mismatch_summary[5]}')
        worksheet.write_url('G9', f"internal:'{sheet_name}'!D22", structure.white_thin_back, f'{mismatch_summary[6]}')
        worksheet.write_url('G10', f"internal:'{sheet_name}'!D23", structure.white_thin_back, f'{mismatch_summary[7]}')
        worksheet.write_url('G11', f"internal:'{sheet_name}'!D24", structure.white_thin_back, f'{mismatch_summary[8]}')
        worksheet.write_url('G12', f"internal:'{sheet_name}'!D25", structure.white_thin_back, f'{mismatch_summary[9]}')
    except IndexError:
        pass

    worksheet.merge_range('G1:G2', 'Summary of CRD vs. TRR', structure.teal_middle)
    worksheet.write('G3', f'Match = {sum(match_tally)} | Mismatch = {sum(mismatch_tally)} | '
                          f'Missing = {sum(missing_tally)}', structure.blue_middle_big)


def to_microsoft(structure, worksheet):
    worksheet.write('I3', f'Message to Dipak to reconfigure CRD', structure.blue_middle_big)
    worksheet.write('I4', f'{mismatch_microsoft[1]}', structure.white_thin_back)
    worksheet.write('I5', f'{mismatch_microsoft[2]}', structure.white_thin_back)
    worksheet.write('I6', f'{mismatch_microsoft[3]}', structure.white_thin_back)
    worksheet.write('I7', f'{mismatch_microsoft[4]}', structure.white_thin_back)
    worksheet.write('I8', f'{mismatch_microsoft[5]}', structure.white_thin_back)
    worksheet.write('I9', f'{mismatch_microsoft[6]}', structure.white_thin_back)
    worksheet.write('I10', f'{mismatch_microsoft[7]}', structure.white_thin_back)
    worksheet.write('I11', f'{mismatch_microsoft[8]}', structure.white_thin_back)
    worksheet.write('I12', f'{mismatch_microsoft[9]}', structure.white_thin_back)
    worksheet.merge_range('I1:I2', 'To Microsoft for TRR   vs   CRD', structure.teal_middle)


# def set_groupings(worksheet) -> any:
#     def create_grouping_1():
#         group_start = 13
#         group_end = 43
#         # for trr in unique_trrs:
#         #     while group_start < group_end:
#         while group_start <= group_end:
#             worksheet.set_row(group_start, None, None, {'level': 1, 'hidden': False})
#             group_start += 1
#         worksheet.outline_settings(symbols_below=False)
#
#     def create_grouping_2():
#         group_start = 44
#         group_end = 74
#         # for trr in unique_trrs:
#         #     while group_start < group_end:
#         while group_start <= group_end:
#             worksheet.set_row(group_start, None, None, {'level': 2, 'hidden': False})
#             group_start += 1
#             # group_end += 27
#             worksheet.outline_settings(symbols_below=False)
#
#     create_grouping_1()
#     create_grouping_2()


def create_sheet_2(name_to_id: dict, name_to_ticket: dict, unique_tickets: list, pipe_name: str, 
                   write_book: xlsxwriter.workbook.Workbook, sheet_title: str):
    """
    Create Sheet 2 dedicated towards comparing CRD vs TRR information.
    :param sheet_title:
    :param name_to_id:
    :param name_to_ticket: 
    :param unique_tickets: 
    :param pipe_name: 
    :param write_book: 
    :return: 
    """
    write_sheet = write_book.add_worksheet(sheet_title)
    structure = Structure(write_book)
    crd_scanner.create_csv()

    set_sheet_structure(pipe_name, name_to_id, structure, write_sheet, sheet_title, unique_tickets)
    print(f'  - Writing {sheet_title} data...')
    write_data(structure, write_sheet, unique_tickets)

    create_graphs(write_book, write_sheet, sheet_title)
    # create_summary(structure, write_sheet, sheet_name)
    to_microsoft(structure, write_sheet)
    # set_groupings(write_sheet)

    print(f'    * Finished {sheet_title} Sheet')
