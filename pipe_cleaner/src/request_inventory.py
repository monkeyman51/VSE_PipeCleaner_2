"""
Responsible for handling request input for inventory transactions.  Includes handling part number, logistics, task name,
and other important information needed to be sent to the inventory team to handle material movement.

6/3/2021

"""
import sys
from os import system
from time import strftime
from datetime import datetime

import win32com.client as client
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from psutil import process_iter as task_manager, NoSuchProcess
from pipe_cleaner.src.log_database import access_database_document


def add_basic_documents_info(workbook, basic_info: dict, sheet_name: str) -> dict:
    """
    Handles basic documenting information like data, time, user, and site.
    """
    sheet_name: str = sheet_name.title()

    print(f'\t\t- Creating {sheet_name} excel')
    if 'Request' in sheet_name:
        return add_excel_data(sheet_name, basic_info, workbook['Request'])

    elif 'Update' in sheet_name:
        return add_excel_data(sheet_name, basic_info, workbook['Update'])


def get_template_dictionary(start_location: str, end_location: str) -> dict:
    """
    Get template for fill later.  Everything is optional until it is assigned as main.
    """
    normal_default: str = 'Copy From Task'

    return {'task_name': {'cell': '$C$13',
                          'value': 'optional',
                          'default': normal_default},

            'part_number': {'cell': '$C$17',
                            'value': 'optional',
                            'default': normal_default},

            'start_location': {'cell': '$C$20',
                               'value': start_location.title(),
                               'default': 'None'},

            'end_location': {'cell': '$C$23',
                             'value': end_location.title(),
                             'default': 'None'},

            'notes': {'cell': '$C$26',
                      'value': 'optional',
                      'default': 'None'},

            'pipe_number': {'cell': '$C$29',
                            'value': 'optional',
                            'default': normal_default},

            'machine_name': {'cell': '$C$32',
                             'value': 'optional',
                             'default': normal_default},

            'trr_number': {'cell': '$C$35',
                           'value': 'optional',
                           'default': normal_default},

            'purchase_order': {'cell': '$C$38',
                               'value': 'optional',
                               'default': 'Copy From Package'},

            'main_fields': {'cell': '$F$10',
                            'value': 0,
                            'default': 'None'}}


def get_template_pipe_to_cage(start_location: str, end_location: str) -> dict:
    """
    Get request template based on start and end locations.
    """
    template: dict = get_template_dictionary(start_location, end_location)

    template['task_name']['value']: str = 'main'
    template['part_number']['value']: str = 'main'
    template['pipe_number']['value']: str = 'main'
    template['machine_name']['value']: str = 'main'
    template['trr_number']['value']: str = 'main'

    return get_total_mains(template)


def get_total_mains(template: dict) -> dict:
    """
    Get total of main fields needed to be filled.
    """
    main_fields: int = 0
    for field_name in template:
        state: str = template[field_name]['value']

        if state == 'main':
            main_fields += 1

    template['main_fields']['value'] = main_fields

    return template


def get_template_pipe_to_shipment(start_location: str, end_location: str) -> dict:
    """
    Get request template based on start and end locations.
    """
    template: dict = get_template_dictionary(start_location, end_location)

    template['task_name']['value']: str = 'main'
    template['part_number']['value']: str = 'main'
    template['pipe_number']['value']: str = 'main'
    template['machine_name']['value']: str = 'main'
    template['purchase_order']['value']: str = 'main'

    return get_total_mains(template)


def get_locations_template(basic_info: dict) -> dict:
    """
    Get the excel template based off of start location and end location.
    """
    start: str = basic_info['start'].upper()
    end: str = basic_info['end'].upper()

    if start == 'PIPE' and end == 'CAGE':
        return get_template_pipe_to_cage(start, end)

    elif start == 'PIPE' and end == 'SHIPMENT':
        return get_template_pipe_to_shipment(start, end)


def add_excel_colors(template: dict, worksheet) -> None:
    """
    Add colors based on template given.
    """
    for field_name in template:
        value = str(template[field_name]['value'])
        cell: str = template[field_name]['cell']
        main_cell: str = cell.replace('$C', '$B')

        if cell == '$C$20' or cell == '$C$23':
            worksheet[main_cell].fill = PatternFill(start_color="A5A5A5", fill_type="solid")
            worksheet[cell].fill = PatternFill(start_color="DBDBDB", fill_type="solid")

        elif value == 'optional':
            worksheet[main_cell].fill = PatternFill(start_color="A5A5A5", fill_type="solid")
            worksheet[cell].fill = PatternFill(start_color="DBDBDB", fill_type="solid")


def add_field_data(template: dict, worksheet) -> None:
    """
    Add field based on template given.
    """
    valid_fields: list = []

    for field_name in template:
        value = str(template[field_name]['value'])
        cell: str = template[field_name]['cell']

        if cell == '$F$10':
            continue

        if cell == '$C$20' or cell == '$C$23':
            worksheet[cell].value = value

        elif value != 'optional':
            default = template[field_name]['default']

            worksheet[cell].value = str(default)
            worksheet[cell].font = Font(bold=True)
            validation: str = f'COUNTIFS({cell},"<>{default}",{cell},"<>")'
            valid_fields.append(validation)

    worksheet['D10'].value = f'={"+".join(valid_fields)}'


def add_excel_data(sheet_name: str, basic_info: dict, worksheet) -> dict:
    """
    Adds info to amount of inventory movement and location.
    """
    template: dict = get_locations_template(basic_info)

    main_fields = str(template['main_fields']['value'])
    add_excel_colors(template, worksheet)
    add_field_data(template, worksheet)

    worksheet['D7']: str = f'=IF($D$10={main_fields},"Save and Exit. Check Pipe Cleaner.","")'
    worksheet['F10']: str = main_fields

    add_default_values(basic_info, worksheet, sheet_name)

    return template


def add_default_values(basic_info: dict, worksheet, sheet_name: str) -> None:
    """
    Add default values to the excel form.
    """
    start_location: str = basic_info['start']
    end_location: str = basic_info['end']

    worksheet['D5']: str = f'{sheet_name.upper()} FORM - {start_location.upper()} TO {end_location.upper()}'
    worksheet['K2']: str = strftime('%m/%d/%Y')
    worksheet['K3']: str = strftime('%I:%M %p')
    worksheet['K4']: str = basic_info['name']
    worksheet['K5']: str = basic_info['location']
    worksheet['K7']: str = basic_info['version']
    worksheet['K10'] = int(basic_info['quantity'])


def get_data_from_excel() -> dict:
    """
    Gather essential information from the request form.  Also includes optional data as well.
    """
    # Part Number
    # Start Location
    # End Location
    # Task Name
    pass


def handle_new_file(basic_info: dict) -> load_workbook:
    """
    Delete file. Start new excel file.
    """
    template_source: str = basic_info['request_template']

    try:
        print(f'\n\tRequest Form...')
        return load_workbook(template_source)

    except FileNotFoundError:
        return load_workbook(template_source)

    except PermissionError:
        print(f'\n\tWARNING!!!')
        print(f'\tWARNING!!!')
        print(f'\tWARNING!!!')
        print(f'\n\tClose down request_form to request inventory movement.')
        input(f'\tPress enter to exit.')
        sys.exit()


def clean_default_name(default_user_name: str) -> str:
    """
    Remove unnecessary characters from name to be clean and presentable in the excel document.
    """
    return default_user_name.upper().replace('.', ' ').replace('-EXT', '').title()


def get_request_file_name(user_data: dict) -> str:
    """
    Return file name based on the given start_location and end_location.
    """
    start_location: str = user_data['start']
    end_location: str = user_data['end']

    start: str = start_location.lower().strip().replace(' ', '_')
    end: str = end_location.lower().strip().replace(' ', '_')

    return f'{start}_to_{end}.xlsx'


def get_backup_name() -> str:
    """

    """
    date: str = strftime('%m/%d/%Y').replace('/', '')
    time: str = strftime('%I:%M %p').replace(':', '').replace(' ', '')

    return f'request_{date}_{time}'


def consolidate_basic_info(user_data: dict) -> dict:
    """
    Request inventory amount and gather basic documents info.
    """
    user_data['request_file_name'] = f'logs/{get_backup_name()}'
    user_data['name'] = clean_default_name(user_data['name'])
    user_data['request_file'] = 'request_template.xlsx'
    user_data['request_template'] = 'settings/request_template.xlsx'
    user_data['location'] = str(user_data['location']).replace(' Lab Site', '')

    return user_data


def setup_excel_output(basic_info) -> dict:
    """
    Adds data to the excel output.
    """
    workbook: load_workbook = handle_new_file(basic_info)

    template: dict = add_basic_documents_info(workbook, basic_info, 'request')

    try:
        workbook.save(f'request_form.xlsx')
        return template

    except PermissionError:
        print_line_divider()
        print(f'\tProblem: request_form.xlsx already open.')
        print(f'\tSolution: Close request_form.xlsx and run Pipe Cleaner again.')
        input(f'\n\tPress enter to close:')
        sys.exit()


def print_line_divider() -> None:
    """
    Print terminal line divider.
    """
    print(f'\n\n\t{"-" * 60}\n\n')


def open_excel_file(file_name: str) -> None:
    """
    Automatically opens excel file.
    """
    system(f'start EXCEL.EXE {file_name}')


def close_message(file_name: str) -> None:
    """
    Close message when excel closed.
    """
    print(f'\t\t- {file_name} closed')


def is_excel_file_running(file_name: str) -> bool:
    """
    Check if excel file is running through the task manager.
    """
    for application in task_manager():
        if 'EXCEL.EXE' in application.name().upper():
            try:
                for excel_file in application.as_dict()['cmdline']:
                    if file_name in excel_file:
                        return True

            except ProcessLookupError:
                close_message(file_name)
                return False
            except NoSuchProcess:
                close_message(file_name)
                return False
    else:
        close_message(file_name)
        return False


def is_excel_file_closed(file_name: str, form_type: str) -> bool:
    """
    Pipe Cleaner continues to check excel file runtime until it terminates by the user.
    """
    print(f'\t\t- {file_name} opened')
    print(f'\t\t- Fill out {form_type} form...\n')

    while True:
        if not is_excel_file_running(file_name):
            return True


def is_field_correct(field_input: str) -> bool:
    """
    Checks if valid input for a given field or is empty.
    """
    try:
        field_input: str = field_input.upper()

        if not field_input:
            return False
        elif field_input == '':
            return False
        elif field_input == 'COPY TASK TITLE':
            return False
        elif field_input == 'REQUIRED':
            return False
        elif field_input == 'COPY FROM TASK':
            return False
        elif field_input == 'RACK / STORAGE / OFFSITE':
            return False
        elif field_input == 'OPTIONAL':
            return False
        elif field_input == 'INVENTORY SUPERVISOR':
            return False
        elif field_input == 'SCAN MATERIAL P/N':
            return False
        elif field_input == 'SCAN RACK / STORAGE / OFFSITE':
            return False
        else:
            return True

    except AttributeError:
        return False


def clean_field(task_name: str) -> str:
    """
    Cleans data.
    """
    return task_name.strip()


def respond_request_form_wrong(request_fields: dict) -> None:
    """
    Message printed if request form is filled incorrectly.
    """
    bad_message: str = 'Invalid or Missing'

    if not is_field_correct(request_fields['task_name']):
        print(f'\t\t\t- Task Name: {bad_message}')

    if not is_field_correct(request_fields['part_number']):
        print(f'\t\t\t- Part Number: {bad_message}')

    if not is_field_correct(request_fields['end_location']):
        print(f'\t\t\t- End Location: {bad_message}\n')


def respond_update_form_wrong(update_fields: dict) -> None:
    """
    Message printed if update form is filled incorrectly.
    """
    bad_message: str = 'Invalid or Missing'

    if not is_field_correct(update_fields['approved_by']):
        print(f'\t\t\t- Approved By: {bad_message}')

    if not is_field_correct(update_fields['task_name']):
        print(f'\t\t\t- Task Name: {bad_message}')

    if not is_field_correct(update_fields['part_number']):
        print(f'\t\t\t- Part Number: {bad_message}')

    if not is_field_correct(update_fields['start_location']):
        print(f'\t\t\t- Start Location: {bad_message}')

    if not is_field_correct(update_fields['end_location']):
        print(f'\t\t\t- End Location: {bad_message}')

    if not is_all_serial_numbers_scanned(update_fields):
        print(f'\t\t\t- Serial Numbers Scanned: Not Enough Scanned')

    print(f'')


def get_request_fields(worksheet: load_workbook) -> dict:
    """
    Get request fields information.
    """
    return {'task_name': worksheet['C13'].value,
            'part_number': worksheet['C17'].value,
            'start_location': worksheet['C20'].value,
            'end_location': worksheet['C23'].value,
            'notes': worksheet['C26'].value,
            'pipe_number': worksheet['C29'].value,
            'machine_name': worksheet['C32'].value,
            'trr_number': worksheet['C35'].value,
            'purchase_order': worksheet['C39'].value,
            'quantity': worksheet['K10'].value}


def get_update_fields(request_worksheet, update_worksheet) -> dict:
    """
    Get update fields information.
    """
    return {'approved_by': update_worksheet['C13'].value,
            'part_number': update_worksheet['C16'].value,
            'start_location': update_worksheet['C19'].value,
            'end_location': update_worksheet['C22'].value,
            'task_name': request_worksheet['C13'].value,
            'notes': clean_optional_field(request_worksheet['C26'].value),
            'pipe_number': clean_optional_field(request_worksheet['C29'].value),
            'machine_name': clean_optional_field(request_worksheet['C32'].value),
            'trr_number': clean_optional_field(request_worksheet['C35'].value),
            'purchase_order': clean_optional_field(request_worksheet['C38'].value),
            'current_quantity': get_serial_numbers(update_worksheet),
            'total_quantity': request_worksheet['K10'].value,
            'date': update_worksheet['K2'].value,
            'time': update_worksheet['K3'].value,
            'user': update_worksheet['K4'].value,
            'site': update_worksheet['K5'].value,
            'form': update_worksheet['K6'].value,
            'version': update_worksheet['K7'].value}


def clean_optional_field(field_input: str) -> str:
    """
    Get rid of optional or empty of optional fields.
    """
    upper_field_input = str(field_input).upper()

    if 'OPTIONAL' in upper_field_input:
        return 'None'
    elif upper_field_input == '':
        return 'None'
    else:
        return field_input


def get_serial_numbers(worksheet: load_workbook) -> list:
    """
    Get serial numbers from Update Form sheet.
    """
    print(f'\t\t- Collecting Serial Numbers')
    serial_numbers: list = []
    for number in range(12, 1012):
        current_value: str = worksheet[f'H{number}'].value

        if not current_value or current_value == '' or current_value == 'Scan Here':
            pass
        else:
            serial_numbers.append(current_value)

    return serial_numbers


def print_notification_receipt(request_fields: dict) -> None:
    """
    Print in terminal the receipt of the notification that will be sent to the inventory team from the request form.
    """
    print(f'\n-------------------------------------------------------------------')
    print(f'\n\tRequest Form:')
    print(f'\t\t- Quantity: {request_fields["quantity"]}\n')
    for index, field_name in enumerate(request_fields, start=0):

        if index == 4:
            print(f'')
        if 'Quantity' in field_name:
            continue

        field_input: str = request_fields[field_name]

        if is_field_correct(field_input):
            field_name: str = field_name.replace('_', ' ').title()
            print(f'\t\t- {field_name}: {field_input}')

        else:
            field_name: str = field_name.replace('_', ' ').title()
            print(f'\t\t- {field_name}: None')

    user_response: str = response_notification_receipt()
    process_notification_input(user_response, request_fields)


def response_notification_receipt() -> str:
    """
    Response notification receipt.
    """
    print_line_divider()
    print(f'\tNotify inventory team?')
    print(f'\t\tY  ->  Yes')
    print(f'\t\tN  ->  No')
    return input(f'\n\tResponse: ')


def add_email_body(request_fields: dict) -> str:
    """
    Add the inventory body of the email.
    """
    body_message: str = 'Please provide confirmation via email to person requesting inventory.  Press "reply" ' \
                        'found on this page to do so.  \nBe sure to coordinate both time and location for meeting ' \
                        'up with the inventory requester.  \n'
    body_message += f'\nInventory Request:'
    body_message += f'\n\t- Quantity: {request_fields["quantity"]}\n'

    for index, field in enumerate(request_fields, start=0):

        if 'quantity' in field:
            continue

        field_input: str = request_fields[field]
        field_name: str = field.replace('_', ' ').title()

        if index == 4:
            body_message += f'\n'

        if is_field_correct(field_input):
            body_message += f'\n\t- {field_name}: {field_input}'
        else:
            body_message += f'\n\t- {field_name}: None'

    return body_message


def email_inventory_request(request_fields: dict) -> None:
    """
    Sends email to people responsible dealing with inventory.
    """
    print(f'\t\t- Writing email to Inventory_Kirkland@veritasdcservices.com')
    # real_location: str = 'Inventory_Kirkland@veritasdcservices.com'
    person_location: str = 'joe.ton@VeritasDCservices.com'

    outlook = client.Dispatch('Outlook.Application')
    message = outlook.CreateItem(0)
    message.To = person_location
    message.Subject = request_fields['task_name']
    message.Body = add_email_body(request_fields)
    message.Send()

    print(f'\t\t- Email sent.')
    print(f'\n\tWait for a response from inventory team via email.  Prepare to meet with update form excel ready.')
    print(f'\tIf time lapsed, can restart Pipe Cleaner to bring up update form via pressing "U"')

    input(f'\n\tPress enter to open Update Form: ')
    print(f'\t\t- Getting Update Form...')


def process_notification_input(user_response: str, request_fields: dict) -> None:
    """
    Handle the response to notification on whether to send to inventory team.
    """
    user_response: str = user_response.upper()
    if user_response == 'Y':
        email_inventory_request(request_fields)
        start_update_form()

    elif 'YES' in user_response:
        email_inventory_request(request_fields)
        start_update_form()

    elif user_response == 'N':
        print(f'\n\tNotification Stopped.')
        input(f'\tPress enter to exit Pipe Cleaner.')
        sys.exit()

    elif 'NO' in user_response:
        print(f'\n\tNotification Stopped.')
        input(f'\tPress enter to exit Pipe Cleaner.')
        sys.exit()


def check_request_form(file_name: str, template: dict) -> None:
    """
    Check if inventory request form done correctly.
    """
    print(f'\t\t- Checking update form')

    worksheet = load_workbook(file_name)['Request']

    request_fields: dict = get_request_fields(worksheet)

    if is_essential_fields_filled(request_fields, template):
        print_notification_receipt(request_fields)

    else:
        print(f'\n-------------------------------------------------------------------')
        print(f'\n\t\t*** INVALID REQUEST FORM ***')
        respond_request_form_wrong(request_fields)
        print(f'-------------------------------------------------------------------')
        print(f'\t\t* Need to fix inventory request form...')
        input(f'\t\t- Press Enter to open request form: ')
        open_excel_file(file_name)
        start_request_stage(file_name, template)


def get_update_form_data(fields_data: dict) -> dict:
    """
    Get data from update form to log in the cloud database.
    """
    serial_number_logs: dict = {}

    serial_numbers_entered: list = fields_data['current_quantity']

    for serial_number in serial_numbers_entered:
        serial_number_logs[serial_number]: dict = {}
        current: dict = serial_number_logs[serial_number]

        current['approved_by'] = fields_data['approved_by']
        current['date'] = fields_data['date']
        current['end_location'] = fields_data['end_location']
        current['form'] = fields_data['form']
        current['machine_name'] = fields_data['machine_name']
        current['notes'] = fields_data['notes']
        current['part_number'] = fields_data['part_number']
        current['pipe_number'] = fields_data['pipe_number']
        current['pipe_number'] = fields_data['pipe_number']
        current['purchase_order'] = fields_data['purchase_order']
        current['site'] = fields_data['site']
        current['start_location'] = fields_data['start_location']
        current['task_name'] = fields_data['task_name']
        current['time'] = fields_data['time']
        current['total_quantity'] = fields_data['total_quantity']
        current['trr_number'] = fields_data['trr_number']
        current['user'] = fields_data['user']
        current['version'] = fields_data['version']

    return serial_number_logs


def check_update_form(file_name: str) -> None:
    """
    Check if inventory request form done correctly.
    """
    print(f'\t\t- Checking update form')

    fields_data: dict = get_update_fields(load_workbook(f'request_form.xlsx')['Request'],
                                          load_workbook(f'update_form.xlsx')['Update'])

    if is_update_fields_filled(fields_data):
        print_line_divider()
        print(f'\tLog inventory:')
        print(f'\t\tY  ->  Yes')
        print(f'\t\tN  ->  No')
        response: str = input(f'\n\tResponse:').upper()

        if 'Y' == response:
            update_serial_numbers(fields_data)
            print(f'\t')

        elif 'N' == response:
            print(f'\tNot updated to database.')
            input(f'\tPress enter to exit: ')
            sys.exit()

    else:
        print(f'\n-------------------------------------------------------------------')
        print(f'\n\t\t*** INVALID UPDATE FORM ***')
        respond_update_form_wrong(fields_data)
        print(f'-------------------------------------------------------------------')
        print(f'\n\t\t* Need to fix inventory update form...')
        input(f'\t\t- Press enter to open excel:')
        open_excel_file(file_name)
        check_update_form(file_name)


def update_serial_numbers(fields_data: dict) -> None:
    """
    Update serial numbers based off of update form's fields.
    """
    update_data: dict = get_update_form_data(fields_data)
    document: MongoClient = access_database_document('serial_numbers', 'all')

    for serial_number in update_data:

        db_serial_numbers: dict = document.find_one({'_id': serial_number})
        if not db_serial_numbers:
            current: dict = update_data[serial_number]

            document.insert_one({'_id': serial_number,
                                 'approved_by': current['approved_by'],
                                 'date': current['date'],
                                 'end_location': current['end_location'],
                                 'form': current['form'],
                                 'machine_name': current['machine_name'],
                                 'notes': current['notes'],
                                 'part_number': current['part_number'],
                                 'pipe_number': current['pipe_number'],
                                 'purchase_order': current['purchase_order'],
                                 'site': current['site'],
                                 'start_location': current['start_location'],
                                 'task_name': current['task_name'],
                                 'time': current['time'],
                                 'trr_number': current['trr_number'],
                                 'user': current['user'],
                                 'version': current['version'],
                                 'date_time': datetime.today().strftime('%Y-%m-%d-%H:%M:%S')})


def is_essential_fields_filled(request_fields: dict, template: dict) -> bool:
    """
    Checks to make sure the four essential fields are filled in for the request form.
    """
    main_fields: int = template['main_fields']['value']

    count: int = 0
    for field_name in template:

        try:
            field_input: str = request_fields[field_name]
            value: str = template[field_name]['value']

            if 'main' in value and field_input != 'Copy From Ticket':
                count += 1

        except KeyError:
            pass

    if main_fields == count:
        return True

    else:
        return False


def is_update_fields_filled(request_fields: dict) -> bool:
    """
    Checks to make sure the four essential fields are filled in for the request form.
    """
    is_quantity_match: bool = is_all_serial_numbers_scanned(request_fields)

    if is_field_correct(request_fields['approved_by']) and \
            is_field_correct(request_fields['part_number']) and \
            is_field_correct(request_fields['start_location']) and \
            is_field_correct(request_fields['end_location']) and \
            is_field_correct(request_fields['task_name']) and \
            is_quantity_match:
        return True

    else:
        return False


def is_all_serial_numbers_scanned(request_fields: dict) -> bool:
    """
    Are all the serial numbers scanned properly. Do they match the request with the total scanned?
    """
    total_quantity = int(request_fields['total_quantity'])
    current_quantity = int(len(request_fields['current_quantity']))

    if total_quantity == current_quantity:
        return True
    elif total_quantity != current_quantity:
        return False


def setup_update_form() -> None:
    """
    Setup up the fields to have consistent look.
    """
    request_worksheet = load_workbook('request_form.xlsx')['Request']

    update_workbook: load_workbook = load_workbook(f'settings/update_template.xlsx')
    update_worksheet = update_workbook['Update']

    update_worksheet['C13'] = 'Inventory Supervisor'
    update_worksheet['C16'] = 'Scan Material P/N'
    update_worksheet['C19'] = 'Scan Rack / Storage / Offsite'
    update_worksheet['C22'] = 'Scan Rack / Storage / Offsite'

    update_worksheet['D5'] = str(request_worksheet['D5'].value).replace('REQUEST', 'UPDATE')

    update_worksheet['C25'] = request_worksheet['C13'].value
    update_worksheet['C28'] = request_worksheet['C26'].value
    update_worksheet['C29'] = request_worksheet['C29'].value
    update_worksheet['C30'] = request_worksheet['C32'].value
    update_worksheet['C31'] = request_worksheet['C35'].value
    update_worksheet['C32'] = request_worksheet['C38'].value

    update_worksheet['K2'] = request_worksheet['K2'].value
    update_worksheet['K3'] = request_worksheet['K3'].value
    update_worksheet['K4'] = request_worksheet['K4'].value
    update_worksheet['K5'] = request_worksheet['K5'].value
    update_worksheet['K6'] = request_worksheet['K6'].value
    update_worksheet['K7'] = request_worksheet['K7'].value
    update_worksheet['M10'] = request_worksheet['K10'].value

    try:
        update_workbook.save(f'update_form.xlsx')
    except PermissionError:
        print(f'\tPermission Denied: update_template.xlsx is already open')
        print(f'\tPlease close file and restart Pipe Cleaner')
        input(f'\tPress enter to close:')
        sys.exit()


def start_request_stage(file_name: str, template: dict) -> None:
    """
    Start the request stage before notifying inventory team
    """
    if is_excel_file_closed(file_name, 'request'):
        check_request_form(file_name, template)


def start_update_form() -> None:
    """
    After sending email to inventory team. Have update form ready.
    """
    file_name: str = 'update_form.xlsx'

    setup_update_form()
    open_excel_file(file_name)

    if is_excel_file_closed(file_name, 'update'):
        check_update_form(file_name)


def main_method(user_data: dict) -> None:
    """
    Starting point for handling inventory requests sent to inventory team.
    """
    basic_info: dict = consolidate_basic_info(user_data)

    template: dict = setup_excel_output(basic_info)

    open_excel_file('request_form.xlsx')
    start_request_stage('request_form.xlsx', template)
