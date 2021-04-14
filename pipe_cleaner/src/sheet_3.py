from pipe_cleaner.src.data_access import request_ado
from pipe_cleaner.src.credentials import Path
from xlrd import open_workbook
from time import strftime
from json import loads
import xlrd


total_components: list = []


class Structure:
    initial = 8

    time = strftime('%I:%M %p')
    date = strftime('%B %d, %Y')
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
        self.blue_component.set_align('vcenter')

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
        self.teal_left.set_font_size('18')

        self.teal_middle = wb.add_format({'border': 2})
        self.teal_middle.set_bg_color('#00B0F0')
        self.teal_middle.set_align('center')
        self.teal_middle.set_border_color('white')
        self.teal_middle.set_bold()
        self.teal_middle.set_font_color('white')
        self.teal_middle.set_font_size('18')
        self.teal_middle.set_align('vcenter')

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

        self.big_blue_font = wb.add_format({'border': 2})
        self.big_blue_font.set_bold()
        self.big_blue_font.set_font_size('22')
        self.big_blue_font.set_font_color('#1f497d')
        self.big_blue_font.set_border_color('white')

        self.italic_blue_font = wb.add_format({'border': 2})
        self.italic_blue_font.set_italic()
        self.italic_blue_font.set_font_size('14')
        self.italic_blue_font.set_font_color('#1f497d')
        self.italic_blue_font.set_border_color('white')

        self.bold_italic_blue_font = wb.add_format({'border': 2})
        self.bold_italic_blue_font.set_italic()
        self.bold_italic_blue_font.set_bold()
        self.bold_italic_blue_font.set_font_size('16')
        self.bold_italic_blue_font.set_font_color('#1f497d')
        self.bold_italic_blue_font.set_border_color('white')
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


alphabet = {'1': 'A',
            '2': 'B',
            '3': 'C',
            '4': 'D',
            '5': 'E',
            '6': 'F',
            '7': 'G',
            '8': 'H',
            '9': 'I',
            '10': 'J',
            '11': 'K',
            '12': 'L',
            '13': 'M',
            '14': 'N',
            '15': 'O',
            '16': 'P',
            '17': 'Q',
            '18': 'R',
            '19': 'S',
            '20': 'T',
            '21': 'U',
            '22': 'V',
            '23': 'W',
            '24': 'X',
            '25': 'Y',
            '26': 'Z'}


def check_empty_both(sys, trr):
    """
    Check if system configuration or TRR configuration is empty.
    Returns boolean on True/False.

    :param sys: Console Server system configuration
    :type sys: str
    :param trr: Test Run Request configuration
    :type sys: str
    :return: returns true if correct
    :rtype: bool
    """
    trr = str(trr)
    sys = str(sys)
    if trr == "" or sys == "":
        return True
    elif trr.isspace() or sys.isspace():
        return True
    else:
        return False


def set_sheet_structure(full_name, write_book, ws, sheet_title, uniques):
    structure = Structure(write_book)

    # Freeze Panes
    ws.freeze_panes(9, 0)

    ws.set_row(0, 12, structure.white)
    ws.set_row(1, 60, structure.white)
    ws.set_row(2, 25, structure.white)
    ws.set_row(3, 20, structure.white)
    ws.set_row(4, 20, structure.white)
    ws.set_row(5, 20, structure.white)
    ws.set_row(6, 20, structure.white)
    ws.set_row(7, 30, structure.white)

    ws.set_column('A:A', 1, structure.white)
    ws.set_column('B:B', 1, structure.white)
    ws.set_column('C:C', 1, structure.white)
    ws.set_column('D:D', 30, structure.white)
    ws.set_column('E:E', 70, structure.white)
    ws.set_column('F:F', 85, structure.white)
    ws.set_column('G:G', 3, structure.white)
    ws.set_column('H:H', 30, structure.white)
    ws.set_column('I:I', 82, structure.white)
    ws.set_column('J:J', 70, structure.white)
    ws.set_column('K:K', 3, structure.white)
    ws.set_column('L:L', 30, structure.white)
    ws.set_column('M:M', 70, structure.white)
    ws.set_column('N:N', 85, structure.white)
    ws.set_column('O:O', 3, structure.white)
    ws.set_column('P:P', 30, structure.white)
    ws.set_column('Q:Q', 70, structure.white)
    ws.set_column('R:R', 85, structure.white)
    ws.set_column('S:S', 3, structure.white)
    ws.set_column('T:T', 30, structure.white)
    ws.set_column('U:U', 70, structure.white)
    ws.set_column('V:V', 85, structure.white)
    ws.set_column('W:W', 3, structure.white)
    ws.set_column('X:X', 30, structure.white)
    ws.set_column('Y:Y', 70, structure.white)
    ws.set_column('Z:Z', 85, structure.white)
    ws.set_column('AA:AA', 3, structure.white)
    ws.set_column('AB:AB', 30, structure.white)
    ws.set_column('AC:AC', 70, structure.white)
    ws.set_column('AD:AD', 85, structure.white)
    ws.set_column('AE:AE', 3, structure.white)
    ws.set_column('AF:AF', 30, structure.white)
    ws.set_column('AG:AG', 70, structure.white)
    ws.set_column('AH:AH', 85, structure.white)

    if len(uniques) > 1:
        total = f'{len(uniques)}'
    else:
        total = f'{len(uniques)}'

    ws.write('D3', f' Pipe Cleaner - {sheet_title}', structure.big_blue_font)
    ws.write('D4', f'       Kirkland Lab Site', structure.bold_italic_blue_font)
    ws.write('D5', f'       Pipe Name: {full_name}', structure.bold_italic_blue_font)
    ws.write('D6', f'       Total Tickets (TRR): {total}', structure.bold_italic_blue_font)
    ws.write('D7', f'         {Structure.date} - {Structure.time}', structure.italic_blue_font)

    ws.write('F3', f'   Summary and Graphs for Overall Pipe Report', structure.big_blue_font)
    ws.write('F4', f'         Coming soon in this area...', structure.italic_blue_font)

    letter_start = 4
    count_start = 1
    while count_start <= len(uniques):
        letter = alphabet[str(letter_start)]
        ws.write(f'{letter}9', 'Component', structure.teal_left)
        letter_start += 4
        count_start += 1

    ws.write('D9', 'Component', structure.teal_left)

    letter_start_1 = 5
    unique_start_1 = 0
    count_start_1 = 1
    while count_start_1 <= len(uniques):
        letter = alphabet[str(letter_start_1)]
        trr_id = uniques[unique_start_1]
        ws.write(f'{letter}9', f'TRR {trr_id}', structure.teal_middle)
        letter_start_1 += 4
        unique_start_1 += 1
        count_start_1 += 1

    letter_start_2 = 6
    count_start_2 = 1
    while count_start_2 <= len(uniques):
        letter = alphabet[str(letter_start_2)]
        ws.write(f'{letter}9', f'PM Notes', structure.teal_middle)
        letter_start_2 += 4
        count_start_2 += 1

    ws.insert_image('B2', 'pipe_cleaner/img/vse_logo.png')


def cleans_xlrd_cell(cell: str) -> str:
    """
    WHAT: Removes unnecessary information when getting cell data from an excel sheet using the XLRD library
    WHY: Needs to clean xlrd cell to be presentable in the excel output
    :param cell: xlrd cell pulled from excel sheet
    :return: empty string, parsed cell, or original cell
    """
    if "empty:" in cell:
        return ''

    elif "text:" in cell:
        new_cell = cell.replace("text:", "").replace("'", "")
        return new_cell

    elif "number:" in cell:
        new_cell = cell.replace("number:", "").replace("'", "")
        return new_cell

    else:
        return cell


def get_components_from_toggle(excel_path: str, ticket_to_type: dict, workbook, worksheet) -> dict:
    """
    Get components from PM Section Toggle excel file within input folder based on ON/OFF/REQUIRED status.
    :return: list
    """
    component_to_status: dict = {}

    for ticket_index, ticket in enumerate(ticket_to_type, start=0):

        new_ticket: dict = {}
        components: list = []

        # Pipe might have different open_work
        request_type_sheet = open_workbook(excel_path).sheet_by_name(ticket_to_type.get(ticket))

        # Gets components and their status to zip together later into dict
        for part in range(6, request_type_sheet.nrows):

            # Converts to string type for later parsing
            key = str(request_type_sheet.cell(part, 1))
            value = str(request_type_sheet.cell(part, 2))

            # Cleans xlrd extra information ie. text and number
            cleaned_key = cleans_xlrd_cell(key)
            cleaned_value = cleans_xlrd_cell(value)

            # Gets list and dict for excel structure
            components.append(cleaned_key)
            new_ticket.update({cleaned_key: cleaned_value})

        component_to_status[ticket] = new_ticket
        component_to_status[f'{ticket}_list'] = components

        write_components(workbook, worksheet, ticket_index, components, component_to_status, ticket)

    return component_to_status


def get_information_from_toggle(excel_path: str, toggle_max_row, start_row):
    """
    Get components from PM Section Toggle excel file within input folder based on ON/OFF status.
    :return: list
    """
    workbook = xlrd.open_workbook(excel_path)
    sheet = workbook.sheet_by_index(0)

    information_toggle = []

    while start_row < toggle_max_row:
        status_column = sheet.cell_value(start_row, 2)
        information_column = sheet.cell_value(start_row, 3)

        upper_status_column = str(status_column).upper()

        if 'END' in status_column:
            break
        elif 'OFF' in upper_status_column:
            pass
        elif 'ON' in upper_status_column:
            information_toggle.append(information_column)
        elif '' in upper_status_column:
            information_toggle.append('None')
        start_row += 1

    return information_toggle


def check_gpu(component_toggle: list, target_configuration):
    for item in component_toggle:
        if 'GPU' in item:
            if 'GPU' not in target_configuration:
                component_toggle.pop(item)


def get_component_value(component_index, current_component_name, test_run_request_json):
    """
    Get component value from TRR table
    :param component_index:
    :param current_component_name:
    :param test_run_request_json:
    :return:
    """
    try:
        four_items = parse_configuration_names(component_index, current_component_name)
        item_1 = four_items[0]
        item_2 = four_items[1]
        item_3 = four_items[2]
        item_4 = four_items[3]

        component = requested_configuration(test_run_request_json, item_1, item_2, item_3, item_4)

        return component

    except TypeError:
        return 'None'
    except IndexError:
        return 'None'


def write_status_column(workbook, worksheet, unique_tickets, component_to_status: dict):
    """

    :param workbook:
    :param worksheet:
    :param unique_tickets:
    :param component_to_status:
    :return:
    """
    structure = Structure(workbook)
    letter_start = 5

    def write_data(number_to_letter, test_run_request_json, ticket_number):
        component_index = 0
        row_start = 10

        component_content = component_to_status[f"{ticket_number}_list"]

        while component_index < len(component_content):
            letter = alphabet[str(number_to_letter)]
            current_component_name = component_content[component_index]
            current_component_state = component_to_status[ticket_number].get(component_content[component_index])

            # Write blank rows first
            if current_component_state == 'OFF':
                component_index += 1

            elif current_component_name == '':
                worksheet.write(f'{letter}{row_start}', '', structure.white)
                component_index += 1
                row_start += 1

            elif current_component_state == 'END':
                break

            elif current_component_state == 'REQUIRED' or current_component_state == 'ON':

                component_value = get_component_value(component_index, current_component_name, test_run_request_json)

                # Required, but missing. Turns red
                if current_component_state == 'REQUIRED' and component_value == '':
                    worksheet.write(f'{letter}{row_start}', f'Missing Required Info', structure.bad_cell)
                    component_index += 1
                    row_start += 1

                # Required and filled. Turns blue
                elif current_component_state == 'REQUIRED':
                    worksheet.write(f'{letter}{row_start}', f'{component_value}', structure.blue_middle)
                    component_index += 1
                    row_start += 1

                # ON, missing and unfilled. Turns missing
                elif current_component_state == 'ON' and component_value == '':
                    worksheet.write(f'{letter}{row_start}', f'{component_value}', structure.missing_cell)
                    component_index += 1
                    row_start += 1

                elif current_component_state == 'ON' and component_value == 'N/A':
                    worksheet.write(f'{letter}{row_start}', f'', structure.missing_cell)
                    component_index += 1
                    row_start += 1

                elif current_component_state == 'ON' and component_value == 'None':
                    worksheet.write(f'{letter}{row_start}', f'', structure.missing_cell)
                    component_index += 1
                    row_start += 1

                elif current_component_state == 'ON' and component_value is None:
                    worksheet.write(f'{letter}{row_start}', f'', structure.missing_cell)
                    component_index += 1
                    row_start += 1

                # ON, missing and filled. Turns blue
                else:
                    worksheet.write(f'{letter}{row_start}', f'{component_value}', structure.blue_middle)
                    component_index += 1
                    row_start += 1


    def write_os_data(number_to_letter, test_run_request_json, ticket_number):
        unique_start = 0
        row_start = 31

        letter = alphabet[str(number_to_letter)]
        os_component = requested_configuration(test_run_request_json, 'SERVER', 'OS', 'SERVER OS', 'SERVER OS')
        # os_component = test_run_request_json['Server OS']

        if os_component == '' or os_component is None or os_component == 'N/A' or os_component == 'None' \
                or os_component == '--':
            worksheet.write(f'{letter}{row_start}', '', structure.missing_cell)
        else:
            worksheet.write(f'{letter}{row_start}', f'{os_component}', structure.blue_middle)

    for ticket in unique_tickets:

        # Creates local data to extract via files
        request_ado(ticket)

        with open(f'{Path.info}{str(ticket)}/final.json') as file:
            ticket_file = loads(file.read())

        write_data(letter_start, ticket_file, ticket)
        write_os_data(letter_start, ticket_file, ticket)

        letter_start += 4


def write_components(workbook, worksheet, ticket_index: int, components: list, component_to_status: dict, ticket):
    """
    Write components on Component Column of the PM view.
    :param component_to_status:
    :param workbook:
    :param worksheet:
    :param ticket_index:
    :param components:
    :return:
    """
    structure = Structure(workbook)

    total_index = 2
    component_index = 0
    letter_start = 4 + 4 * ticket_index
    total = len(components) + total_index

    while total_index < total:
        number = Structure.initial + total_index
        letter = alphabet[str(letter_start)]
        current_component = component_to_status[ticket].get(components[component_index])

        if current_component == '' or current_component is None:
            worksheet.write(f'{letter}{number}', '', structure.white)
            total_index += 1
            component_index += 1

        elif current_component == 'OFF':
            component_index += 1

        elif current_component == 'ON':
            worksheet.write(f'{letter}{number}', f'{components[component_index]}', structure.blue_component)
            total_index += 1
            component_index += 1

        elif current_component == 'REQUIRED':
            worksheet.write(f'{letter}{number}', f'{components[component_index]}', structure.blue_component)
            total_index += 1
            component_index += 1

        elif current_component == 'END':
            break


def set_notes_column(wb, ws, uniques, info_toggle):
    structure = Structure(wb)

    def write_data(unique_number):
        begin = 2
        list_count = 0
        letter_start = 6 + 4 * unique_number
        total = len(info_toggle) + begin

        while begin < total:
            num = Structure.initial + begin
            letter = alphabet[str(letter_start)]

            if info_toggle[list_count] == '':
                ws.write(f'{letter}{num}', '', structure.missing_cell)
            elif info_toggle[list_count] == 'None':
                ws.write(f'{letter}{num}', '', structure.white)
            else:
                ws.write(f'{letter}{num}', f'{info_toggle[list_count]}', structure.blue_middle)
            begin += 1
            list_count += 1

    for number, item in enumerate(uniques):
        write_data(number)


def parse_configuration_names(list_position, component_name):
    """
    Parses the first, second, last terms and last term bit for future gathering data from TRR.
    Returns a new list of configuration names.
    :return: three all caps terms
    """
    new_configuration_names = []

    first = str(component_name).split(' ')[0]
    try:
        second = str(component_name).split(' ')[1]
    except IndexError:
        second = first
        pass
    last = str(component_name).split(' ')[-1]
    last_bit = last

    together = f'{first.upper()} {second.upper()} {last.upper()} {last_bit.upper()}'
    new_configuration_names.append(together)

    first_upper = first.upper()
    second_upper = second.upper()
    last_upper = last.upper()
    last_bit_upper = last_bit.upper()

    items = [first_upper, second_upper, last_upper, last_bit_upper]

    return items



def requested_configuration(test_run_request_json, term_1, term_2, term_3, term_4):
    """
    Finds configuration with four key parts based on text segments within the configuration name.
    Returns the requested configuration for potential discrepancies between TRRs and Host configuration.

    :param test_run_request_json: created JSON file for TRR
    :type test_run_request_json: dict
    :param term_1: text segment within configuration name
    :type term_1: str
    :param term_2:text segment within configuration name
    :type term_2: str
    :param term_3: text segment within configuration name
    :type term_3: str
    :param term_4: first 2 characters text segment within configuration name
    :type term_4: str

    :return: requested configuration from ADO
    :rtype: str
    """
    upper_term_1 = str(term_1).upper()
    upper_term_2 = str(term_2).upper()
    upper_term_3 = str(term_3).upper()
    upper_term_4 = str(term_4).upper()

    for dependency in test_run_request_json:
        upper_item = str(dependency).upper()
        if upper_term_1 in upper_item \
                and upper_term_2 in upper_item \
                and upper_term_3 in upper_item \
                and upper_term_4 in upper_item:
            configuration = str(test_run_request_json[dependency])
            return configuration


def ticket_to_type(unique_tickets: list, ticket_to_ado: dict) -> dict:
    """
    Pair unique ticket to Request Type from ADO
    For Unique Toggle Settings
    :param unique_tickets:
    :param ticket_to_ado:
    :return:
    """
    ticket_type: dict = {}

    for ticket in unique_tickets:
        json_ado = ticket_to_ado.get(ticket)
        system_title = f"{json_ado['fields']['System.Title']}".upper()
        if 'DIMM' in system_title:
            ticket_type[ticket] = 'DIMM'
        elif 'NVME' in system_title:
            ticket_type[ticket] = 'NVME'
        elif 'SSD' in system_title:
            ticket_type[ticket] = 'SSD'
        elif 'HDD' in system_title:
            ticket_type[ticket] = 'HDD'

    return ticket_type


def create_sheet_3(full_name, sheet_title, write_book, unique_tickets, ticket_to_ado):
    toggle_max_row = 1_000
    start_row = 6

    excel_path = 'settings/all_toggles.xlsx'

    # Gather and Create Template for Sheet
    write_sheet = write_book.add_worksheet(sheet_title)
    ticket_type: dict = ticket_to_type(unique_tickets, ticket_to_ado)
    component_to_status: dict = get_components_from_toggle(excel_path, ticket_type, write_book, write_sheet)
    information_list = get_information_from_toggle(excel_path, toggle_max_row, start_row)

    print(f'\t\t- Writing {sheet_title} data...')
    set_sheet_structure(full_name, write_book, write_sheet, sheet_title, unique_tickets)
    write_status_column(write_book, write_sheet, unique_tickets, component_to_status)
    set_notes_column(write_book, write_sheet, unique_tickets, information_list)

    print(f'\t\t\t* Finished {sheet_title} Sheet')

    return component_to_status
