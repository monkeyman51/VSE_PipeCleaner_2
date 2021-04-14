import xlsxwriter
import xlsxwriter.exceptions
# from pipe_cleaner.sheet_4 import create_sheet_4
# from pipe_cleaner.sheet_5 import create_sheet_5
# from pipe_cleaner.sheet_6 import create_sheet_6
# from pipe_cleaner.sheet_7 import create_sheet_7
from colorama import Fore, Style

from pipe_cleaner.src.sheet_1 import create_sheet_1
# from pipe_cleaner.sheet_2 import create_sheet_2
from pipe_cleaner.src.sheet_3 import create_sheet_3


def check_excel_open(pipe_name, path):
    try:
        write_book = xlsxwriter.Workbook(path)
        write_book.close()
    except xlsxwriter.exceptions.FileCreateError:
        print(f'\n  Permission Denied: Cannot create excel_{pipe_name}.xlsx if excel_{pipe_name}.xlsx '
              f'is already opened.')
        print(f'\n  Close excel_{pipe_name}.xlsx first...')
        input(f'\n  Press enter to try again...')
        check_excel_open(pipe_name, path)


def all_excel_sheets(pipe_info: dict, unique_tickets: list, name_to_ticket: dict, name_to_id: dict, ticket_to_ado: dict,
                     document_filepath: dict):
    """
    Main method for creating excel sheet that includes all sheets.
    :param document_filepath:
    :param ticket_to_ado:
    :param pipe_info:
    :param name_to_id:
    :param name_to_ticket:
    :param unique_tickets:
    :return: return timer for creating excel sheets
    """

    path = f'pipes/{pipe_info["full_name"]}/excel_{pipe_info["pipe_name"]}.xlsx'
    check_excel_open(pipe_info["pipe_name"], path)
    write_book = xlsxwriter.Workbook(path)

    title_1 = 'TRR vs Console Server'
    # title_2 = 'CRD vs TRR'
    title_3 = 'PM Section'
    # title_4 = 'Technician Section'
    # title_5 = 'VSS Section'
    # title_6 = 'Engineer Section'
    # title_7 = 'Surface Room'

    print(f'\t{pipe_info["full_name"]} | {pipe_info["description"]}: {Fore.GREEN}Creating Excel File{Style.RESET_ALL}')

    component_to_status = create_sheet_3(pipe_info['full_name'], title_3, write_book, unique_tickets, ticket_to_ado)
    create_sheet_1(pipe_info, name_to_id, name_to_ticket, unique_tickets, write_book, title_1, component_to_status,
                   document_filepath)
    # create_sheet_2(name_to_id, name_to_ticket, unique_tickets, pipe_name, write_book, title_2)
    # create_sheet_4(pipe_name, host_ids, title_4, write_book, CRD, unique_tickets)
    # create_sheet_5(pipe_name, host_ids, title_5, write_book, CRD, unique_tickets)
    # create_sheet_6(pipe_name, host_ids, title_6, write_book, CRD, unique_tickets)
    # create_sheet_7(pipe_name, host_ids, title_7, write_book, CRD, unique_tickets)

    write_book.close()

    return path
