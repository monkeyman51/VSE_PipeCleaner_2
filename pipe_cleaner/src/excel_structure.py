def layout(worksheet, structure):
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
    worksheet.set_column('B:B', 25, structure.white)
    worksheet.set_column('C:C', 15, structure.white)
    worksheet.set_column('D:D', 16, structure.white)
    worksheet.set_column('E:E', 24, structure.white)
    worksheet.set_column('F:F', 15, structure.white)
    worksheet.set_column('G:G', 30, structure.white)
    worksheet.set_column('H:H', 30, structure.white)
    worksheet.set_column('I:I', 70, structure.white)
    worksheet.set_column('J:J', 70, structure.white)
    worksheet.set_column('K:K', 25, structure.white)
    worksheet.set_column('L:L', 25, structure.white)
    worksheet.set_column('M:M', 25, structure.white)
    worksheet.set_column('N:N', 25, structure.white)
    worksheet.set_column('O:O', 25, structure.white)
    worksheet.set_column('P:P', 25, structure.white)


def trr_vs_console_server(start: int, worksheet, structure):
    """
    :param start:
    :param worksheet:
    :param structure:
    :return:
    """
    name_to_number: dict = {}

    table_names: list = ['Machine Name',
                         'TRR ID',
                         'Type',
                         'Component',
                         'Status',
                         'Ticket',
                         'System']

    # Image
    worksheet.insert_image('A1', 'pipe_cleaner/img/vse_logo.png')

    # Number part of the excel position
    num = str(start)

    initial = 0
    while initial < len(table_names):
        little = chr(ord('b') + initial)
        let = str(little).upper()

        worksheet.write(f'{let}{num}', f'{table_names[initial]}', structure.teal_middle)

        # Create key for dictionary
        name = str(table_names[initial]).lower().replace(' ', '_')
        number = initial + 1

        name_to_number[name] = str(number)

        initial += 1

    return name_to_number
