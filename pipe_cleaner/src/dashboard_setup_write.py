# from pipe_cleaner.src.naming_conventions.host_group_name import main_method as check_pipe_name
# from pipe_cleaner.src.naming_conventions.description_name import main_method as check_description
#
#
# def write_pipe_name(column_name, structure, worksheet, start, check_color, column_location, pipe_name: str):
#     """
#     Current Pipe:
#     :param pipe_name:
#     :param column_name:
#     :param structure:
#     :param worksheet:
#     :param start:
#     :param check_color:
#     :param column_location:
#     :return:
#     """
#     # Stores RIGHT/WRONG/INDEX to flag later
#     all_responses: list = []
#
#     naming_standard_results: dict = check_pipe_name(pipe_name)
#
#     # Unpacks results RIGHT, WRONG, or INDEX
#     # Index 2 represents the results
#     try:
#         all_responses.append(naming_standard_results.get('length')[2])
#         all_responses.append(naming_standard_results.get('pipe_in_name')[2])
#         all_responses.append(naming_standard_results.get('pipe_number')[2])
#         for character_result in naming_standard_results.get('check_paths'):
#             all_responses.append(character_result[2])
#     except IndexError:
#         pass
#
#     if 'WRONG' in all_responses:
#         worksheet.write(f'{column_location}{start}', f'{pipe_name}', structure.left_neutral_cell)
#     else:
#         if check_color == 1:
#             worksheet.write(f'{column_location}{start}', f'{pipe_name}', structure.blue_left)
#         elif check_color == 0:
#             worksheet.write(f'{column_location}{start}', f'{pipe_name}', structure.alt_blue_left)
#
#
# def write_description(column_name, structure, worksheet, start, check_color, column_location, description_name: str):
#     """
#     Checks description field of the Console Server in Host Group page.
#     Warns user if description field is away from standard agreed on in VSE
#     :param column_name:
#     :param structure:
#     :param worksheet:
#     :param start:
#     :param check_color:
#     :param column_location:
#     :param description_name:
#     :return:
#     """
#     description_results: list = check_description(description_name)
#
#     if 'WRONG' in description_results:
#         worksheet.write(f'{column_location}{start}', f'{description_name}', structure.left_neutral_cell)
#     else:
#         if check_color == 1:
#             worksheet.write(f'{column_location}{start}', f'{description_name}', structure.blue_left)
#         else:
#             worksheet.write(f'{column_location}{start}', f'{description_name}', structure.alt_blue_left)
#
#
# def write_checked_out_to(column_name, structure, worksheet, start, check_color, column_location, write_data):
#     """
#     Add information to the
#     :param column_name:
#     :param structure:
#     :param worksheet:
#     :param start:
#     :param check_color:
#     :param column_location:
#     :param write_data:
#     :return:
#     """
#     clean_name: list = []
#
#     if '.' in write_data:
#         new_name = str(write_data).replace('.', ' ')
#         capital_name = new_name.title()
#         clean_name.append(capital_name)
#
#     if check_color == 1:
#         if write_data is None or write_data == '':
#             worksheet.write(f'{column_location}{start}', f'', structure.missing_cell)
#         else:
#             worksheet.write(f'{column_location}{start}', f'{clean_name[0]}', structure.blue_middle)
#     else:
#         if write_data is None or write_data == '':
#             worksheet.write(f'{column_location}{start}', f'', structure.missing_cell)
#         else:
#             worksheet.write(f'{column_location}{start}', f'{clean_name[0]}', structure.alt_blue_middle)
#
#
# def write_status(column_name, structure, worksheet, start, check_color, column_location, write_data):
#     """
#
#     :param column_name:
#     :param structure:
#     :param worksheet:
#     :param start:
#     :param check_color:
#     :param column_location:
#     :param write_data:
#     :return:
#     """
#     if check_color == 1:
#         if write_data is None or write_data == '':
#             worksheet.write(f'{column_location}{start}', f'', structure.missing_cell)
#         else:
#             worksheet.write(f'{column_location}{start}', f'{write_data}', structure.blue_middle)
#     else:
#         if write_data is None or write_data == '':
#             worksheet.write(f'{column_location}{start}', f'', structure.missing_cell)
#         else:
#             worksheet.write(f'{column_location}{start}', f'{write_data}', structure.alt_blue_middle)
#
#
# def check_due_dates(due_date_data) -> str:
#     check_for_one = list(set(due_date_data))
#
#     if len(check_for_one) > 1:
#         return 'More Than One Due Date'
#     else:
#         try:
#             return check_for_one[0]
#         except IndexError:
#             return 'None'
#
#
# def write_due_date(column_name, structure, worksheet, start, check_color, column_location, write_data: dict):
#     """
#     Checks for multiple due dates
#     :param column_name:
#     :param structure:
#     :param worksheet:
#     :param start:
#     :param check_color:
#     :param column_location:
#     :param write_data: dict
#     :return:
#     """
#
#     try:
#         expected_task_start = write_data.get('expected_task_start')
#         actual_date = parsed_date(expected_task_start)
#
#         if check_color == 1:
#             if actual_date is None or actual_date == '' or actual_date == 'None':
#                 worksheet.write(f'{column_location}{start}', f'', structure.missing_cell)
#             else:
#                 worksheet.write(f'{column_location}{start}', f'{actual_date}', structure.blue_middle)
#
#         elif check_color == 0:
#             if actual_date is None or actual_date == '' or actual_date == 'None':
#                 worksheet.write(f'{column_location}{start}', f'', structure.missing_cell)
#             else:
#                 worksheet.write(f'{column_location}{start}', f'{actual_date}', structure.alt_blue_middle)
#
#     except AttributeError:
#         worksheet.write(f'{column_location}{start}', f'', structure.missing_cell)
#
#     try:
#         actual_qual_start_date = write_data.get('actual_qual_start_date')
#         actual_date = parsed_date(actual_qual_start_date)
#
#         if check_color == 1:
#
#             if actual_date is None or actual_date == '' or actual_date == 'None':
#                 worksheet.write(f'J{start}', f'', structure.missing_cell)
#             else:
#                 worksheet.write(f'J{start}', f'{actual_date}', structure.blue_middle)
#
#         elif check_color == 0:
#
#             if actual_date is None or actual_date == '' or actual_date == 'None':
#                 worksheet.write(f'J{start}', f'', structure.missing_cell)
#             else:
#                 worksheet.write(f'J{start}', f'{actual_date}', structure.alt_blue_middle)
#
#     except AttributeError:
#         worksheet.write(f'J{start}', f'', structure.missing_cell)
#
#
# def parsed_date(due_date: str) -> str:
#     """
#     Convert weird CSI due date into actual due date.
#     :param due_date:
#     :return:
#     """
#
#     try:
#         actual_date = due_date[0:10]
#         raw_year = actual_date.split('-')[0]
#         raw_month = actual_date.split('-')[1]
#         raw_day = actual_date.split('-')[2]
#
#         actual_month = convert_month(raw_month)
#
#         convert_date = f'{actual_month} {raw_day}, {raw_year}'
#         return convert_date
#     except TypeError:
#         pass
#
#
# def convert_month(month_in_number: str) -> str:
#     """
#
#     :param month_in_number:
#     :return:
#     """
#     if month_in_number == '01':
#         return 'January'
#     elif month_in_number == '02':
#         return 'February'
#     elif month_in_number == '03':
#         return 'March'
#     elif month_in_number == '04':
#         return 'April'
#     elif month_in_number == '05':
#         return 'May'
#     elif month_in_number == '06':
#         return 'June'
#     elif month_in_number == '07':
#         return 'July'
#     elif month_in_number == '08':
#         return 'August'
#     elif month_in_number == '09':
#         return 'September'
#     elif month_in_number == '10':
#         return 'October'
#     elif month_in_number == '11':
#         return 'November'
#     elif month_in_number == '12':
#         return 'December'
#
#
# def convert_day(day_in_number: str) -> str:
#     """
#
#     :param day_in_number:
#     :return:
#     """
#     if day_in_number == '01':
#         return '1st'
#     elif day_in_number == '02':
#         return '2nd'
#     elif day_in_number == '03':
#         return '3rd'
#
#
# def write_tech(column_name, structure, worksheet, start, check_color, column_location, write_data: dict,
#                current_pipe_name):
#     """
#
#     :param column_name:
#     :param structure:
#     :param worksheet:
#     :param start:
#     :param check_color:
#     :param column_location:
#     :param write_data:
#     :return:
#     """
#     # Unpack data
#     total_tally: int = write_data.get('total_tally')
#     match_tally: int = write_data.get('match_tally')
#
#     try:
#         if total_tally == 'None' or total_tally is None or total_tally == 0 or \
#                 match_tally == 'None' or match_tally is None or match_tally == 0:
#             worksheet.write(f'{column_location}{start}', f'', structure.missing_cell)
#         else:
#             tally_division: str = str((match_tally / total_tally) * 100)
#             print(f'{current_pipe_name} - {tally_division}')
#
#             if check_color == 1:
#                 if tally_division == '100.0':
#                     worksheet.write(f'{column_location}{start}', f'100 %',
#                                     structure.good_cell)
#                 else:
#                     worksheet.write(f'{column_location}{start}', f'{tally_division[0:2]} %',
#                                     structure.blue_middle)
#             elif check_color == 0:
#                 if tally_division == '100.0':
#                     worksheet.write(f'{column_location}{start}', f'100 %',
#                                     structure.good_cell)
#                 else:
#                     worksheet.write(f'{column_location}{start}', f'{tally_division[0:2]} %',
#                                     structure.alt_blue_middle)
#     except TypeError:
#         worksheet.write(f'{column_location}{start}', f'', structure.missing_cell)
#
#
# def setup_write(column_name, structure, worksheet, start, check_color, column_location, write_data,
#                       current_pipe_name):
#     if column_name == 'pipe_name':
#         write_pipe_name(column_name, structure, worksheet, start, check_color, column_location, write_data)
#
#     elif column_name == 'description':
#         write_description(column_name, structure, worksheet, start, check_color, column_location, write_data)
#
#     elif column_name == 'checked_out_to':
#         write_checked_out_to(column_name, structure, worksheet, start, check_color, column_location, write_data)
#
#     elif column_name == 'status':
#         write_status(column_name, structure, worksheet, start, check_color, column_location, write_data)
#
#     elif column_name == 'due_date_column':
#         write_due_date(column_name, structure, worksheet, start, check_color, column_location, write_data)
#
#     elif column_name == 'eng_column':
#         worksheet.write(f'{column_location}{start}', f'{write_data}', structure.missing_cell)
#
#     elif column_name == 'tech_column':
#         write_tech(column_name, structure, worksheet, start, check_color, column_location,
#                    write_data, current_pipe_name)
#
#     elif column_name == 'pm_column':
#         worksheet.write(f'{column_location}{start}', f'{write_data}', structure.missing_cell)
#
#     elif column_name == 'setup_column':
#         worksheet.write(f'{column_location}{start}', f'{write_data}', structure.missing_cell)
