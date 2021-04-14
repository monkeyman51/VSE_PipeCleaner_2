from colorama import Fore, Style

numbers = ['0',
           '1',
           '2',
           '3',
           '4',
           '5',
           '6',
           '7',
           '8',
           '9']

alphabet = ['a',
            'b',
            'c',
            'd',
            'e',
            'f',
            'g',
            'h',
            'i',
            'j',
            'k',
            'l',
            'm',
            'n',
            'o',
            'p',
            'q',
            'r',
            's',
            't',
            'u',
            'v',
            'w',
            'x',
            'y',
            'z']


def process_bios(csv_line: str):
    """
    Find Bios in CSV Line
    :param csv_line:
    :return:
    """
    potential_components: list = []

    # BIOS - 17 Characters Total
    position_00 = []
    position_01 = []
    position_02 = []
    position_03 = []
    position_04 = []
    position_05 = []  # Period starts here
    position_06 = 'B'
    position_07 = 'S'
    position_08 = '.'
    position_09 = []
    position_10 = []
    position_11 = []
    position_12 = []
    position_13 = '.'
    position_14 = []
    position_15 = []
    position_16 = []

    upper_bios = csv_line.upper()

    if csv_line == '' or csv_line is None:
        print(f'   WARNING: {Fore.RED}No BIOS{Style.RESET_ALL}')

    elif '.BS.' in upper_bios:
        slice_line = [character for character in csv_line]
        for index, character in enumerate(slice_line, start=1):
            position_00.clear()
            position_01.clear()
            position_02.clear()
            position_03.clear()
            position_04.clear()
            position_05.clear()
            # position_06.clear()
            # position_07.clear()
            # position_08.clear()
            position_09.clear()
            position_10.clear()
            position_11.clear()
            position_12.clear()
            # position_13.clear()
            position_14.clear()
            position_15.clear()
            position_16.clear()

            if '.' in character:

                index_00 = index - 6
                index_01 = index - 5
                index_02 = index - 4
                index_03 = index - 3
                index_04 = index - 2
                index_05 = index - 1
                index_06 = index
                index_07 = index + 1
                index_08 = index + 2
                index_09 = index + 3
                index_10 = index + 4
                index_11 = index + 5
                index_12 = index + 6
                index_13 = index + 7
                index_14 = index + 8
                index_15 = index + 9
                index_16 = index + 10

                if slice_line[index_06] != 'B':
                    continue
                if slice_line[index_07] != 'S':
                    continue
                if slice_line[index_08] != '.':
                    continue
                if slice_line[index_13] != '.':
                    continue

                if index_00 < 0:
                    continue

                position_00.append(slice_line[index_00])
                position_01.append(slice_line[index_01])
                position_02.append(slice_line[index_02])
                position_03.append(slice_line[index_03])
                position_04.append(slice_line[index_04])
                position_05.append(slice_line[index_05])
                # position_06.append(slice_line[index_06])
                # position_07.append(slice_line[index_07])
                # position_08.append(slice_line[index_08])
                position_09.append(slice_line[index_09])
                position_10.append(slice_line[index_10])
                position_11.append(slice_line[index_11])
                position_12.append(slice_line[index_12])
                # position_13.append(slice_line[index_13])
                position_14.append(slice_line[index_14])
                position_15.append(slice_line[index_15])
                position_16.append(slice_line[index_16])

                potential_components.append(f'{position_00[0]}'
                                            f'{position_01[0]}'
                                            f'{position_02[0]}'
                                            f'{position_03[0]}'
                                            f'{position_04[0]}'
                                            f'{position_05[0]}'
                                            f'{position_06}'
                                            f'{position_07}'
                                            f'{position_08}'
                                            f'{position_09[0]}'
                                            f'{position_10[0]}'
                                            f'{position_11[0]}'
                                            f'{position_12[0]}'
                                            f'{position_13}'
                                            f'{position_14[0]}'
                                            f'{position_15[0]}'
                                            f'{position_16[0]}')
        return potential_components

    else:
        return None


def process_bmc(csv_line: str):
    """
    Find BMC in CSV Line
    :param csv_line:
    :return:
    """
    potential_components: list = []

    # BMC - 16 Characters Total
    position_00 = []
    position_01 = []
    position_02 = []
    position_03 = []
    position_04 = []
    position_05 = []  # Period starts here
    position_06 = 'B'
    position_07 = 'C'
    position_08 = '.'
    position_09 = []
    position_10 = []
    position_11 = []
    position_12 = []
    position_13 = '.'
    position_14 = []
    position_15 = []

    upper_bmc = csv_line.upper()

    if csv_line == '' or csv_line is None:
        print(f'   WARNING: {Fore.RED}No BMC{Style.RESET_ALL}')

    elif '.BC.' in upper_bmc:
        slice_line = [character for character in csv_line]
        for index, character in enumerate(slice_line, start=1):
            position_00.clear()
            position_01.clear()
            position_02.clear()
            position_03.clear()
            position_04.clear()
            position_05.clear()
            # position_06.clear()
            # position_07.clear()
            # position_08.clear()
            position_09.clear()
            position_10.clear()
            position_11.clear()
            position_12.clear()
            # position_13.clear()
            position_14.clear()
            position_15.clear()

            if '.' in character:

                index_00 = index - 6
                index_01 = index - 5
                index_02 = index - 4
                index_03 = index - 3
                index_04 = index - 2
                index_05 = index - 1
                index_06 = index
                index_07 = index + 1
                index_08 = index + 2
                index_09 = index + 3
                index_10 = index + 4
                index_11 = index + 5
                index_12 = index + 6
                index_13 = index + 7
                index_14 = index + 8
                index_15 = index + 9
                index_16 = index + 10

                if slice_line[index_06] != 'B':
                    continue
                if slice_line[index_07] != 'C':
                    continue
                if slice_line[index_08] != '.':
                    continue
                if slice_line[index_13] != '.':
                    continue

                if index_00 < 0:
                    continue

                position_00.append(slice_line[index_00])
                position_01.append(slice_line[index_01])
                position_02.append(slice_line[index_02])
                position_03.append(slice_line[index_03])
                position_04.append(slice_line[index_04])
                position_05.append(slice_line[index_05])
                # position_06.append(slice_line[index_06])
                # position_07.append(slice_line[index_07])
                # position_08.append(slice_line[index_08])
                position_09.append(slice_line[index_09])
                position_10.append(slice_line[index_10])
                position_11.append(slice_line[index_11])
                position_12.append(slice_line[index_12])
                # position_13.append(slice_line[index_13])
                position_14.append(slice_line[index_14])
                position_15.append(slice_line[index_15])

                potential_components.append(f'{position_00[0]}'
                                            f'{position_01[0]}'
                                            f'{position_02[0]}'
                                            f'{position_03[0]}'
                                            f'{position_04[0]}'
                                            f'{position_05[0]}'
                                            f'{position_06}'
                                            f'{position_07}'
                                            f'{position_08}'
                                            f'{position_09[0]}'
                                            f'{position_10[0]}'
                                            f'{position_11[0]}'
                                            f'{position_12[0]}'
                                            f'{position_13}'
                                            f'{position_14[0]}'
                                            f'{position_15[0]}')

        return potential_components

    else:
        return None


def process_tpm(csv_line: str):
    """
    Find TPM in CSV Line
    :param csv_line:
    :return:
    """
    potential_components: list = []

    # TPM - 11 Characters Total
    # ex. 7.85.4555.0-
    position_00 = []
    position_01 = '.'
    position_02 = []
    position_03 = []
    position_04 = '.'
    position_05 = []
    position_06 = []
    position_07 = []
    position_08 = []
    position_09 = '.'
    position_10 = []

    upper_bmc = csv_line.upper()

    if csv_line == '' or csv_line is None:
        print(f'   WARNING: {Fore.RED}No TPM{Style.RESET_ALL}')

    elif 'TPM' in upper_bmc:
        slice_line = [character for character in csv_line]
        for index, character in enumerate(slice_line, start=1):
            position_00.clear()
            # position_01.clear()
            position_02.clear()
            position_03.clear()
            # position_04.clear()
            position_05.clear()
            position_06.clear()
            position_07.clear()
            position_08.clear()
            # position_09.clear()
            position_10.clear()

            if '.' in character:

                index_00 = index - 2
                index_01 = index - 1
                index_02 = index
                index_03 = index + 1
                index_04 = index + 2
                index_05 = index + 3
                index_06 = index + 4
                index_07 = index + 5
                index_08 = index + 6
                index_09 = index + 7
                index_10 = index + 8

                if slice_line[index_01] != '.':
                    continue
                if slice_line[index_04] != '.':
                    continue
                if slice_line[index_09] != '.':
                    continue

                if index_00 < 0:
                    continue

                position_00.append(slice_line[index_00])
                # position_01.append(slice_line[index_01])
                position_02.append(slice_line[index_02])
                position_03.append(slice_line[index_03])
                # position_04.append(slice_line[index_04])
                position_05.append(slice_line[index_05])
                position_06.append(slice_line[index_06])
                position_07.append(slice_line[index_07])
                position_08.append(slice_line[index_08])
                # position_09.append(slice_line[index_09])
                position_10.append(slice_line[index_10])

                potential_components.append(f'{position_00[0]}'
                                            f'{position_01}'
                                            f'{position_02[0]}'
                                            f'{position_03[0]}'
                                            f'{position_04}'
                                            f'{position_05[0]}'
                                            f'{position_06[0]}'
                                            f'{position_07[0]}'
                                            f'{position_08[0]}'
                                            f'{position_09}'
                                            f'{position_10[0]}')

        return potential_components

    else:
        return None


def process_cpld(csv_line: str):
    """
    Find CPLD in CSV Line
    :param csv_line:
    :return:
    """
    potential_components: list = []

    # TPM - 11 Characters Total
    # ex. V042
    position_00 = 'V'
    position_01 = []
    position_02 = []
    position_03 = []

    upper_bmc = csv_line.upper()

    if csv_line == '' or csv_line is None:
        print(f'   WARNING: {Fore.RED}No CPLD{Style.RESET_ALL}')

    elif 'CPLD' in upper_bmc:
        slice_line = [character for character in csv_line]
        for index, character in enumerate(slice_line):
            # position_00.clear()
            position_01.clear()
            position_02.clear()
            position_03.clear()

            if 'V' in character:

                index_00 = index  # Start Here
                index_01 = index + 1
                index_02 = index + 2
                index_03 = index + 3

                print(slice_line[index_00])

                if slice_line[index_00] != 'V':
                    continue

                if index_00 < 0:
                    continue

                # position_00.append(slice_line[index_00])
                position_01.append(slice_line[index_01])
                position_02.append(slice_line[index_02])
                position_03.append(slice_line[index_03])

                potential_components.append(f'{position_00}'
                                            f'{position_01[0]}'
                                            f'{position_02[0]}'
                                            f'{position_03[0]}')
        return potential_components

    else:
        return None


def process_cerberus(csv_line: str):
    """
    Find Cerberus Version in CSV Line
    :param csv_line:
    :return:
    """
    potential_components: list = []

    # Cerberus - 7 Characters Total
    position_00 = []
    position_01 = '.'  # Period starts here
    position_02 = []
    position_03 = '.'
    position_04 = []
    position_05 = '.'
    position_06 = []

    upper_csv = csv_line.upper()

    if csv_line == '' or csv_line is None:
        print(f'   WARNING: {Fore.RED}No Cerberus{Style.RESET_ALL}')

    elif 'CERBERUS' in upper_csv:
        slice_line = [character for character in csv_line]
        for index, character in enumerate(slice_line, start=1):
            position_00.clear()
            # position_01.clear()
            position_02.clear()
            # position_03.clear()
            position_04.clear()
            # position_05.clear()
            position_06.clear()

            if '.' in character:

                index_00 = index - 2
                index_01 = index - 1
                index_02 = index
                index_03 = index + 1
                index_04 = index + 2
                index_05 = index + 3
                index_06 = index + 4

                if slice_line[index_01] != '.':
                    continue
                if slice_line[index_03] != '.':
                    continue
                if slice_line[index_05] != '.':
                    continue

                if index_00 < 0:
                    continue

                position_00.append(slice_line[index_00])
                # position_01.append(slice_line[index_01])
                position_02.append(slice_line[index_02])
                # position_03.append(slice_line[index_03])
                position_04.append(slice_line[index_04])
                # position_05.append(slice_line[index_05])
                position_06.append(slice_line[index_06])

                potential_components.append(f'{position_00[0]}'
                                            f'{position_01}'
                                            f'{position_02[0]}'
                                            f'{position_03}'
                                            f'{position_04[0]}'
                                            f'{position_05}'
                                            f'{position_06[0]}')

        return potential_components

    else:
        return None


def process_fpga(csv_line: str):
    """
    Find FPGA Version in CSV Line
    :param csv_line:
    :return:
    """
    potential_components: list = []

    # Cerberus - 7 Characters Total
    position_00 = []
    position_01 = '.'  # Period starts here
    position_02 = []
    position_03 = '.'
    position_04 = []
    position_05 = '.'
    position_06 = []

    upper_csv = csv_line.upper()

    if csv_line == '' or csv_line is None:
        print(f'   WARNING: {Fore.RED}No Cerberus{Style.RESET_ALL}')

    elif 'CERBERUS' in upper_csv:
        slice_line = [character for character in csv_line]
        for index, character in enumerate(slice_line, start=1):
            position_00.clear()
            # position_01.clear()
            position_02.clear()
            # position_03.clear()
            position_04.clear()
            # position_05.clear()
            position_06.clear()

            if '.' in character:

                index_00 = index - 2
                index_01 = index - 1
                index_02 = index
                index_03 = index + 1
                index_04 = index + 2
                index_05 = index + 3
                index_06 = index + 4

                if slice_line[index_01] != '.':
                    continue
                if slice_line[index_03] != '.':
                    continue
                if slice_line[index_05] != '.':
                    continue

                if index_00 < 0:
                    continue

                position_00.append(slice_line[index_00])
                # position_01.append(slice_line[index_01])
                position_02.append(slice_line[index_02])
                # position_03.append(slice_line[index_03])
                position_04.append(slice_line[index_04])
                # position_05.append(slice_line[index_05])
                position_06.append(slice_line[index_06])

                potential_components.append(f'{position_00[0]}'
                                            f'{position_01}'
                                            f'{position_02[0]}'
                                            f'{position_03}'
                                            f'{position_04[0]}'
                                            f'{position_05}'
                                            f'{position_06[0]}')

        return potential_components

    else:
        return None


def process_rack_manager_version(csv_line: str):
    """
    Find Rack Manager Version in CSV Line
    :param csv_line:
    :return:
    """
    potential_components: list = []

    # Rack Manager Version - 9 Characters Total
    position_00 = []
    position_01 = '.'  # Period starts here
    position_02 = []
    position_03 = '.'
    position_04 = []
    position_05 = []
    position_06 = '.'
    position_07 = []
    position_08 = []

    upper_csv = csv_line.upper()

    if csv_line == '' or csv_line is None:
        print(f'   WARNING: {Fore.RED}No Rack Manager Version{Style.RESET_ALL}')

    elif 'RACK' in upper_csv and 'MANAGER' in upper_csv:
        slice_line = [character for character in csv_line]
        for index, character in enumerate(slice_line, start=1):
            position_00.clear()
            # position_01.clear()
            position_02.clear()
            # position_03.clear()
            position_04.clear()
            position_05.clear()
            # position_06.clear()
            position_07.clear()
            position_08.clear()

            if '.' in character:

                index_00 = index - 2
                index_01 = index - 1
                index_02 = index
                index_03 = index + 1
                index_04 = index + 2
                index_05 = index + 3
                index_06 = index + 4
                index_07 = index + 5
                index_08 = index + 6

                if slice_line[index_01] != '.':
                    continue
                if slice_line[index_03] != '.':
                    continue
                if slice_line[index_06] != '.':
                    continue

                if index_00 < 0:
                    continue

                position_00.append(slice_line[index_00])
                # position_01.append(slice_line[index_01])
                position_02.append(slice_line[index_02])
                # position_03.append(slice_line[index_03])
                position_04.append(slice_line[index_04])
                position_05.append(slice_line[index_05])
                # position_06.append(slice_line[index_06])
                position_07.append(slice_line[index_07])
                position_08.append(slice_line[index_08])

                potential_components.append(f'{position_00[0]}'
                                            f'{position_01}'
                                            f'{position_02[0]}'
                                            f'{position_03}'
                                            f'{position_04[0]}'
                                            f'{position_05[0]}'
                                            f'{position_06}'
                                            f'{position_07[0]}'
                                            f'{position_08[0]}')

        return potential_components

    else:
        return None


string = ',CPLD,C2030C2030.BS.1D07.GN1	C2030.BS.1D07.GN1.zip,,V042,,,'

foo = process_cpld(string)
print(foo)
