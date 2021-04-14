import xlsxwriter
from xlsxwriter.exceptions import FileCreateError
from time import strftime

from pipe_cleaner.src.sheet_1 import setup_excel
from colorama import Fore, Style

import sys


def set_column_names(top_plane_height, worksheet, structure):
    """
    Set up Column Names in the Excel table for adding data later
    :param top_plane_height:
    :param worksheet:
    :param structure:
    :return:
    """
    name_to_number: dict = {}

    column_names: list = ['Pipe Name',
                          'Description',
                          'Checked Out To',
                          'Status',
                          'Setup',
                          'PM',
                          'TECH',
                          'ENG',
                          'Expected Start',
                          'Schedule']

    # Number part of the excel position
    num = str(top_plane_height)

    initial = 0
    while initial < len(column_names):
        little = chr(ord('b') + initial)
        let = str(little).upper()

        if let == 'B' or let == 'C':
            worksheet.write(f'{let}{num}', f'{column_names[initial]}', structure.teal_left)
        else:
            worksheet.write(f'{let}{num}', f'{column_names[initial]}', structure.teal_middle)

        # Create key for dictionary
        name = str(column_names[initial]).lower().replace(' ', '_')
        number = initial + 1

        name_to_number[name] = str(number)

        initial += 1

    return name_to_number


def set_layout(worksheet, structure):
    """
    Beginning of the Excel Structure
    :return:
    """
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

    worksheet.set_column('A:A', 5.5, structure.white)
    worksheet.set_column('B:B', 26, structure.white)
    worksheet.set_column('C:C', 40, structure.white)
    worksheet.set_column('D:D', 25, structure.white)
    worksheet.set_column('E:E', 24, structure.white)
    worksheet.set_column('F:F', 11, structure.white)
    worksheet.set_column('G:G', 11, structure.white)
    worksheet.set_column('H:H', 11, structure.white)
    worksheet.set_column('I:I', 11, structure.white)
    worksheet.set_column('J:J', 25, structure.white)
    worksheet.set_column('K:K', 18, structure.white)
    worksheet.set_column('L:L', 25, structure.white)
    worksheet.set_column('M:M', 25, structure.white)
    worksheet.set_column('N:N', 25, structure.white)
    worksheet.set_column('O:O', 25, structure.white)
    worksheet.set_column('P:P', 25, structure.white)


def set_sheet_structure(prepare_setup, sheet_title, site_location, total_systems):
    """
    Create dashboard structure
    :param prepare_setup:
    :param sheet_title:
    :param site_location:
    :param total_systems:
    :return:
    """
    worksheet = prepare_setup[0]
    structure = prepare_setup[1]

    time = strftime('%I:%M %p')
    date = strftime('%m/%d/%Y')

    # Set Top Plane of Excel Sheet
    top_plane_height = 13

    # Structure of the Excel Sheet
    set_layout(worksheet, structure)
    set_column_names(top_plane_height, worksheet, structure)

    # Freeze Planes
    worksheet.freeze_panes(top_plane_height, 3)

    while top_plane_height < 500:
        worksheet.set_row(top_plane_height, 16.5, structure.white)
        top_plane_height += 1

    # Top Left Plane
    worksheet.insert_image('A1', 'pipe_cleaner/img/vse_logo.png')
    worksheet.write('B5', f' Pipe Cleaner - {sheet_title}', structure.big_blue_font)
    worksheet.write('B6', f'       {site_location}', structure.bold_italic_blue_font)
    worksheet.write('B7', f'       Total Systems - {total_systems}', structure.bold_italic_blue_font)
    worksheet.write('B9', f'       {date} - {time}', structure.italic_blue_font)

    worksheet.write('D5', f'Certain:', structure.big_blue_font)
    worksheet.write('D6', f'Testing:', structure.big_blue_font)
    worksheet.write('D7', f'Not Included:', structure.big_blue_font)
    worksheet.write('D9', f'PM Review:', structure.big_blue_font)
    worksheet.write('D11', f'Setup Review:', structure.big_blue_font)

    worksheet.write('E5', f'', structure.bold_italic_blue_font)
    worksheet.write('E6', f'BIOS, BMC, CPLD, OS, SSD',
                    structure.bold_italic_blue_font)
    worksheet.write('E7', f'TPM', structure.bold_italic_blue_font)
    worksheet.write('E9', f'BIOS, BMC, CPLD, OS, SSD', structure.bold_italic_blue_font)
    worksheet.write('E11', f'Systems with Tickets out of Total Tickets', structure.bold_italic_blue_font)


# def create_graphs(workbook: object, worksheet: object, sheet_name: str):
#     bold = workbook.add_format({'bold': 1})
#
#     # Add the worksheet data that the charts will refer to.
#     headings = ['Number', 'Tallies']
#     data = [
#         ['Match/Present', 'Mismatch', 'Missing'],
#         [sum(match_tally), sum(mismatch_tally), sum(missing_tally)],
#     ]
#
#     worksheet.write_row('A1', headings, bold)
#     worksheet.write_column('A2', data[0])
#     worksheet.write_column('B2', data[1])
#
#     chart_1 = workbook.add_chart({'type': 'bar'})
#     # chart_1 = wb.add_chart({'type': 'pie'})
#
#     # workbook.define_name(f'{sheet_name}', '=Sheet2')
#
#     # Configure the first series.
#     chart_1.add_series({
#         'name': "='" + sheet_name + "'!$B$1",
#         'categories': "='" + sheet_name + "'!$A$2:$A$4",
#         'values': "='" + sheet_name + "'!$B$2:$B$4",
#         'points': [
#             {'fill': {'color': '#00B050'}},
#             {'fill': {'color': '#FF0000'}},
#             {'fill': {'color': '#DCAA1B'}},
#         ],
#     })
#
#     # Configure a second series. Note use of alternative syntax to define ranges.
#     chart_1.add_series({
#         'name': [f"{sheet_name}", 0, 2],
#         'categories': [f"{sheet_name}", 1, 0, 3, 0],
#         'values': [f"{sheet_name}", 1, 2, 3, 2],
#     })
#
#     # Add a chart title and some axis labels.
#     chart_1.set_title({'name': 'Status of TRR vs Console Server'})
#     # chart_1.set_x_axis({'name': 'Tally of Status'})
#     # chart_1.set_y_axis({'name': 'Status'})
#
#     # Chart Style of Graph
#     chart_1.set_style(11)
#     chart_1.set_legend({'none': True})
#
#     # Size of Chart
#     # worksheet.insert_chart('E1', chart_1, {'x_scale': 1.185, 'y_scale': 0.84})
#     worksheet.insert_chart('E1', chart_1, {'x_scale': 1.485, 'y_scale': 0.84})


def main_method(pipe_name: str, console_server_data: dict, ado_data: dict, all_issues: dict):
    """
    Create Pipe dashboard
    :param pipe_name: Host Group name
    :param console_server_data:
    :param ado_data:
    :param all_issues:
    :return:
    """
    workbook = xlsxwriter.Workbook(fr'pipes/{pipe_name}.xlsx')
    prepare_ticket_vs_system: list = setup_excel(workbook, 'Ticket vs System')
    prepare_project_manager_section: list = setup_excel(workbook, 'PM Section')
    prepare_engineer_section: list = setup_excel(workbook, 'Engineer Section')
    prepare_setup: list = setup_excel(workbook, 'Setup')
    prepare_issues: list = setup_excel(workbook, 'Issues')

    set_sheet_structure(prepare_ticket_vs_system, pipe_name, 'Kirkland', 'N/A')
    set_sheet_structure(prepare_project_manager_section, pipe_name, 'Kirkland', 'N/A')
    set_sheet_structure(prepare_engineer_section, pipe_name, 'Kirkland', 'N/A')
    set_sheet_structure(prepare_setup, pipe_name, 'Kirkland', 'N/A')
    set_sheet_structure(prepare_issues, pipe_name, 'Kirkland', 'N/A')

    try:
        workbook.close()
    except FileCreateError:
        print(f'\t{pipe_name}.xlsx {Fore.RED}already open{Style.RESET_ALL}. Please close and restart Pipe Cleaner.')
        input()
        sys.exit()
