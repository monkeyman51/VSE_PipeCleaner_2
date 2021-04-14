from pipe_cleaner.src.naming_conventions.host_group_name import main_method as check_pipe_name
from pipe_cleaner.src.naming_conventions.description_name import main_method as check_description


def write_pipe_name(column_name, structure, worksheet, start, check_color, column_location,
                    pipe_name: str, current_host_group_id: str):
    """
    Current Pipe:
    """
    # Stores RIGHT/WRONG/INDEX to flag later
    all_responses: list = []

    naming_standard_results: dict = check_pipe_name(pipe_name)
    url_path: str = f'http://172.30.1.100/console/host_group_host_list.php?host_group_id={current_host_group_id}'

    # Unpacks results RIGHT, WRONG, or INDEX
    # Index 2 represents the results
    try:
        all_responses.append(naming_standard_results.get('length')[2])
        all_responses.append(naming_standard_results.get('pipe_in_name')[2])
        all_responses.append(naming_standard_results.get('pipe_number')[2])
        for character_result in naming_standard_results.get('check_paths'):
            all_responses.append(character_result[2])
    except IndexError:
        pass

    if 'WRONG' in all_responses:
        worksheet.write(f'{column_location}{start}', f'{pipe_name}', structure.left_neutral_cell)

    elif current_host_group_id != 'None':
        if check_color == 1:
            worksheet.write_url(f'{column_location}{start}', url_path, structure.blue_left, string=pipe_name)

        elif check_color == 0:
            worksheet.write_url(f'{column_location}{start}', url_path, structure.alt_blue_left, string=pipe_name)


def write_description(structure, worksheet, start, check_color, column_location,
                      description_name: str, current_host_group_id: str):
    """
    Checks description field of the Console Server in Host Group page.
    Warns user if description field is away from standard agreed on in VSE
    """
    description_results: list = check_description(description_name)
    url_path: str = f'http://172.30.1.100/console/host_group_host_list.php?host_group_id={current_host_group_id}'

    if 'WRONG' in description_results:
        pass
        worksheet.write(f'{column_location}{start}', f'{description_name}', structure.left_neutral_cell)

    elif current_host_group_id != 'None':
        if check_color == 1:
            worksheet.write_url(f'{column_location}{start}', url_path, structure.blue_left, string=description_name)
        else:
            worksheet.write_url(f'{column_location}{start}', url_path, structure.alt_blue_left, string=description_name)


def write_checked_out_to(column_name, structure, worksheet, start, check_color, column_location, username):
    """
    Add information to the
    :param column_name:
    :param structure:
    :param worksheet:
    :param start:
    :param check_color:
    :param column_location:
    :param username: username in the checked out to field
    :return:
    """
    clean_name: list = []

    # Accounts for NoneType
    try:
        if '.' in username:
            new_name = str(username).replace('.', ' ')
            capital_name = new_name.title()
            clean_name.append(capital_name)
    except TypeError:
        clean_name.append(username)

    if check_color == 1:
        if username == 'None' or username == '':
            worksheet.write(f'{column_location}{start}', f'', structure.missing_cell)
        else:
            worksheet.write(f'{column_location}{start}', f'{clean_name[0]}', structure.blue_middle)
    else:
        if username == 'None' or username == '':
            worksheet.write(f'{column_location}{start}', f'', structure.missing_cell)
        else:
            worksheet.write(f'{column_location}{start}', f'{clean_name[0]}', structure.alt_blue_middle)


def write_status(column_name, structure, worksheet, start, check_color, column_location, write_data):
    """

    :param column_name:
    :param structure:
    :param worksheet:
    :param start:
    :param check_color:
    :param column_location:
    :param write_data:
    :return:
    """
    if check_color == 1:
        if write_data is None or write_data == '':
            worksheet.write(f'{column_location}{start}', f'', structure.missing_cell)
        else:
            worksheet.write(f'{column_location}{start}', f'{write_data}', structure.blue_middle)
    else:
        if write_data is None or write_data == '':
            worksheet.write(f'{column_location}{start}', f'', structure.missing_cell)
        else:
            worksheet.write(f'{column_location}{start}', f'{write_data}', structure.alt_blue_middle)


def check_due_dates(due_date_data) -> str:
    check_for_one = list(set(due_date_data))

    if len(check_for_one) > 1:
        return 'More Than One Due Date'
    else:
        try:
            return check_for_one[0]
        except IndexError:
            return 'None'


def get_difference_dates(expected_start: str, actual_start: str) -> str:
    """
    Find difference between expected start and actual start from ADO
    This is for late, early, or on time for quals.
    :param expected_start:
    :param actual_start:
    :return: Either string None or int
    """
    if expected_start is None or actual_start is None:
        return 'None'

    elif expected_start != 'None' and actual_start != 'None':
        expected_day = int(expected_start.replace(',', '').split(' ')[1])
        actual_day = int(actual_start.replace(',', '').split(' ')[1])
        difference: int = expected_day - actual_day
        clean_difference = str(difference).replace('-', '')

        if difference == 0:
            return 'On Time'
        elif difference > 0 and difference == 1:
            return f'{clean_difference} Day Ahead'
        elif difference > 0:
            return f'{clean_difference} Days Ahead'
        elif difference < 0 and difference == -1:
            return f'{clean_difference} Day Late'
        elif difference < 0:
            return f'{clean_difference} Days Late'

    else:
        return 'None'


def write_due_date(column_name, structure, worksheet, start, check_color, column_location, due_dates: dict):
    """
    Checks for multiple due dates
    :param due_dates:
    :param column_name:
    :param structure:
    :param worksheet:
    :param start:
    :param check_color:
    :param column_location:
    :return:
    """
    # actual_qual_end_date: str = due_dates.get('actual_qual_end_date', 'None')
    actual_qual_start_date: str = due_dates.get('actual_qual_start_date', 'None')
    # expected_task_completion: str = due_dates.get('expected_task_completion', 'None')
    expected_task_start: str = due_dates.get('expected_task_start', 'None')
    base_ticket: str = due_dates.get('base_ticket', 'None')
    azure_url: str = f'https://azurecsi.visualstudio.com/CSI%20Commodity%20Qualification/_workitems/edit/{base_ticket}'

    expected_start: str = parsed_date(expected_task_start)
    actual_date: str = parsed_date(actual_qual_start_date)

    difference_in_date: str = get_difference_dates(expected_start, actual_date)

    if check_color == 1:
        if actual_date == '' or actual_date == 'None' or actual_date is None:
            worksheet.write(f'J{start}', f'', structure.missing_cell)
        else:
            worksheet.write(f'J{start}', f'{actual_date}', structure.blue_middle)

    elif check_color == 0:
        if actual_date == '' or actual_date == 'None' or actual_date is None:
            worksheet.write(f'J{start}', f'', structure.missing_cell)
        else:
            worksheet.write(f'J{start}', f'{actual_date}', structure.alt_blue_middle)

    if check_color == 1:

        if difference_in_date == 'None':
            worksheet.write_url(f'K{start}', azure_url, structure.missing_cell, string='')
            # worksheet.write(f'K{start}', f'', structure.missing_cell)
        else:
            if 'Late' in difference_in_date:
                worksheet.write_url(f'K{start}', azure_url, structure.bad_cell, string=difference_in_date)
                # worksheet.write(f'K{start}', f'{difference_in_date}', structure.bad_cell)
            else:
                worksheet.write_url(f'K{start}', azure_url, structure.good_cell, string=difference_in_date)
                # worksheet.write(f'K{start}', f'{difference_in_date}', structure.good_cell)

    elif check_color == 0:

        if difference_in_date == 'None':
            # worksheet.write(f'K{start}', f'', structure.missing_cell)
            worksheet.write_url(f'K{start}', azure_url, structure.missing_cell, string='')
        else:
            if 'Late' in difference_in_date:
                worksheet.write_url(f'K{start}', azure_url, structure.bad_cell, string=difference_in_date)
                # worksheet.write(f'K{start}', f'{difference_in_date}', structure.bad_cell)
            else:
                worksheet.write_url(f'K{start}', azure_url, structure.good_cell, string=difference_in_date)
                # worksheet.write(f'K{start}', f'{difference_in_date}', structure.good_cell)


def parsed_date(due_date: str) -> str:
    """
    Convert weird CSI due date into actual due date.
    :param due_date:
    :return:
    """

    try:
        actual_date = due_date[0:10]
        raw_year = actual_date.split('-')[0]
        raw_month = actual_date.split('-')[1]
        raw_day = actual_date.split('-')[2]

        actual_month = convert_month(raw_month)

        convert_date = f'{actual_month} {raw_day}, {raw_year}'
        return convert_date
    except TypeError:
        pass
    except IndexError:
        pass


def convert_month(month_in_number: str) -> str:
    """

    :param month_in_number:
    :return:
    """
    if month_in_number == '01':
        return 'January'
    elif month_in_number == '02':
        return 'February'
    elif month_in_number == '03':
        return 'March'
    elif month_in_number == '04':
        return 'April'
    elif month_in_number == '05':
        return 'May'
    elif month_in_number == '06':
        return 'June'
    elif month_in_number == '07':
        return 'July'
    elif month_in_number == '08':
        return 'August'
    elif month_in_number == '09':
        return 'September'
    elif month_in_number == '10':
        return 'October'
    elif month_in_number == '11':
        return 'November'
    elif month_in_number == '12':
        return 'December'


def convert_day(day_in_number: str) -> str:
    """

    :param day_in_number:
    :return:
    """
    if day_in_number == '01':
        return '1st'
    elif day_in_number == '02':
        return '2nd'
    elif day_in_number == '03':
        return '3rd'


def write_tech(column_name, structure, worksheet, start, check_color, column_location, tech_data: dict,
               current_pipe_name: str):
    """

    :param current_pipe_name:
    :param column_name:
    :param structure:
    :param worksheet:
    :param start:
    :param check_color:
    :param column_location:
    :param tech_data:
    :return:
    """
    # Unpack data
    total_tally: int = tech_data.get('total_tally')
    match_tally: int = tech_data.get('match_tally')
    file_path = f'{current_pipe_name}.xlsx'

    try:
        if total_tally == 'None' or total_tally is None or total_tally == 0 or \
                match_tally == 'None' or match_tally is None or match_tally == 0:
            # worksheet.write(f'{column_location}{start}', f'', structure.missing_cell)
            worksheet.write_url(f'{column_location}{start}', file_path, structure.missing_cell, string='')
        else:
            tally_division: str = str((match_tally / total_tally) * 100)

            if check_color == 1:
                if tally_division == '100.0':
                    worksheet.write_url(f'{column_location}{start}', file_path, structure.good_cell, string='100 %')
                    # worksheet.write(f'{column_location}{start}', f'100 %', structure.good_cell)
                else:
                    # worksheet.write(f'{column_location}{start}', f'{tally_division[0:2]} %', structure.blue_middle)
                    worksheet.write_url(f'{column_location}{start}', file_path, structure.blue_middle,
                                        string=f'{tally_division[0:2]} %')

            elif check_color == 0:
                if tally_division == '100.0':
                    # worksheet.write(f'{column_location}{start}', f'100 %', structure.good_cell)
                    worksheet.write_url(f'{column_location}{start}', file_path, structure.good_cell, string=f'100 %')
                else:
                    # worksheet.write(f'{column_location}{start}', f'{tally_division[0:2]} %',
                    # structure.alt_blue_middle)
                    worksheet.write_url(f'{column_location}{start}', file_path, structure.alt_blue_middle,
                                        string=f'{tally_division[0:2]} %')
    except TypeError:
        # worksheet.write(f'{column_location}{start}', f'', structure.missing_cell)
        worksheet.write_url(f'{column_location}{start}', file_path, structure.missing_cell, string=f'')


def write_pm(column_name, structure, worksheet, start, check_color, column_location, tech_data: dict,
             current_pipe_name):
    """

    :param current_pipe_name:
    :param column_name:
    :param structure:
    :param worksheet:
    :param start:
    :param check_color:
    :param column_location:
    :param tech_data:
    :return:
    """
    # Unpack data
    total_tally: int = tech_data.get('ticket_total_tally')
    match_tally: int = tech_data.get('ticket_match_tally')
    file_path = f'{current_pipe_name}.xlsx'

    try:
        if total_tally == 'None' or total_tally is None or total_tally == 0 or \
                match_tally == 'None' or match_tally is None or match_tally == 0:
            worksheet.write(f'{column_location}{start}', f'', structure.missing_cell)
            worksheet.write_url(f'{column_location}{start}', file_path, structure.missing_cell, string='')
        else:
            tally_division: str = str((match_tally / total_tally) * 100)

            if check_color == 1:
                if tally_division == '100.0':
                    # worksheet.write(f'{column_location}{start}', f'100 %',
                    #                 structure.good_cell)
                    worksheet.write_url(f'{column_location}{start}', file_path, structure.good_cell, string='100 %')
                else:
                    # worksheet.write(f'{column_location}{start}', f'{tally_division[0:2]} %',
                    #                 structure.blue_middle)
                    worksheet.write_url(f'{column_location}{start}', file_path, structure.blue_middle,
                                        string=f'{tally_division[0:2]} %')
            elif check_color == 0:
                if tally_division == '100.0':
                    # worksheet.write(f'{column_location}{start}', f'100 %',
                    #                 structure.good_cell)
                    worksheet.write_url(f'{column_location}{start}', file_path, structure.good_cell, string=f'100 %')
                else:
                    # worksheet.write(f'{column_location}{start}', f'{tally_division[0:2]} %',
                    #                 structure.alt_blue_middle)
                    worksheet.write_url(f'{column_location}{start}', file_path, structure.alt_blue_middle,
                                        string=f'{tally_division[0:2]} %')
    except TypeError:
        worksheet.write(f'{column_location}{start}', f'', structure.missing_cell)
        worksheet.write_url(f'{column_location}{start}', file_path, structure.missing_cell, string=f'')


def write_setup(column_name, structure, worksheet, start, check_color, column_location, compare_data: dict,
                current_pipe_name):
    """

    :param column_name:
    :param structure:
    :param worksheet:
    :param start:
    :param check_color:
    :param column_location:
    :param compare_data:
    :return:
    """
    # Unpack data
    total_systems: int = sum(compare_data.get('total_systems'))
    systems_with_ticket: int = sum(compare_data.get('systems_with_ticket'))
    text_output = f'{str(systems_with_ticket)}  /  {str(total_systems)}'
    calculate_output = str(systems_with_ticket / total_systems)

    file_path = f'{current_pipe_name}.xlsx'

    if check_color == 1:
        if calculate_output == '1.0':
            # worksheet.write(f'{column_location}{start}', text_output, structure.good_cell)
            worksheet.write_url(f'{column_location}{start}', file_path, structure.good_cell, string=text_output)
        else:
            # worksheet.write(f'{column_location}{start}', text_output, structure.bad_cell)
            worksheet.write_url(f'{column_location}{start}', file_path, structure.bad_cell, string=text_output)
            # worksheet.write(f'{column_location}{start}', text_output, structure.blue_middle)
    elif check_color == 0:
        if calculate_output == '1.0':
            # worksheet.write(f'{column_location}{start}', text_output, structure.good_cell)
            worksheet.write_url(f'{column_location}{start}', file_path, structure.good_cell, string=text_output)
        else:
            # worksheet.write(f'{column_location}{start}', text_output, structure.bad_cell)
            worksheet.write_url(f'{column_location}{start}', file_path, structure.bad_cell, string=text_output)
            # worksheet.write(f'{column_location}{start}', text_output, structure.alt_blue_middle)


def main_method(column_name, structure, worksheet, start, check_color, column_location, write_data,
                current_pipe_name, current_host_group_id):
    """
    Write data to the Main Dashboard
    :param current_host_group_id:
    :param column_name:
    :param structure:
    :param worksheet:
    :param start:
    :param check_color:
    :param column_location:
    :param write_data:
    :param current_pipe_name:
    :return:
    """
    if column_name == 'pipe_name':
        write_pipe_name(column_name, structure, worksheet, start, check_color, column_location,
                        write_data, current_host_group_id)

    elif column_name == 'description':
        write_description(structure, worksheet, start, check_color, column_location,
                          write_data, current_host_group_id)

    elif column_name == 'checked_out_to':
        write_checked_out_to(column_name, structure, worksheet, start, check_color, column_location, write_data)

    elif column_name == 'status':
        write_status(column_name, structure, worksheet, start, check_color, column_location, write_data)

    elif column_name == 'due_date_column':
        write_due_date(column_name, structure, worksheet, start, check_color, column_location, write_data)

    elif column_name == 'eng_column':
        file_path = f'{current_pipe_name}.xlsx'
        worksheet.write_url(f'{column_location}{start}', file_path, structure.missing_cell, string='')

    elif column_name == 'tech_column':
        write_tech(column_name, structure, worksheet, start, check_color, column_location,
                   write_data, current_pipe_name)

    elif column_name == 'pm_column':
        write_pm(column_name, structure, worksheet, start, check_color, column_location,
                 write_data, current_pipe_name)

    elif column_name == 'setup_column':
        write_setup(column_name, structure, worksheet, start, check_color, column_location, write_data,
                    current_pipe_name)
