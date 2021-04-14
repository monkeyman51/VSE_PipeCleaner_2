import pandas as pd

gen_list = ['GEN5.0',
            'GEN5.1',
            'GEN5.2',
            'GEN5.3',
            'GEN5.4',
            'GEN5.5',
            'GEN5.6',
            'GEN5.7',
            'GEN5.8',
            'GEN5.9',
            'GEN6.0',
            'GEN6.1',
            'GEN6.2',
            'GEN6.3',
            'GEN6.4',
            'GEN6.5',
            'GEN6.6',
            'GEN6.7',
            'GEN6.8',
            'GEN6.9',
            'GEN7.0',
            'GEN7.1',
            'GEN7.2',
            'GEN7.3',
            'GEN7.4',
            'GEN7.5',
            'GEN7.6',
            'GEN7.7',
            'GEN7.8',
            'GEN7.9']

supplier_list = ['[DELL]',
                 '[WIWYNN]',
                 '[ZT]',
                 '[LENOVO]']


def get_gen_number(target):
    parsed_target = str(target).replace(' ', '')
    parsed_target = parsed_target.upper()

    for item in gen_list:
        if item in parsed_target:
            replace_item = item.replace('GEN', '')
            return replace_item


def get_generation(target):
    parsed_target = str(target).replace(' ', '')
    parsed_target = parsed_target.upper()

    for item in gen_list:
        if item in parsed_target:
            replace_item = item[:4:]
            return replace_item


def get_supplier(target):
    for item in supplier_list:
        if item in target:
            parsed_item = item.replace('[', '')
            parsed_item = parsed_item.replace(']', '')
            return parsed_item


def create_csv_from_qcl_list():
    df = pd.read_excel(f'../input/commodity_numbers.xlsx', sheet_name='Main')
    df.to_csv('all_gen.csv')
    configuration_list = df['Configuration'].to_list()
    return configuration_list


def get_column_headers():
    df = pd.read_excel(f'../input/commodity_numbers.xlsx', sheet_name='Main')
    headers = df.columns.tolist()
    return headers


def get_configuration(configuration_list, target, gen_number):
    target_components = []
    possible_configuration = []

    parsed_target = str(target).split('[')[2]
    parsed_target = parsed_target.replace(']', '')

    target_1 = parsed_target.split(' ')[0]
    try:
        target_2 = parsed_target.split(' ')[1]
    except IndexError:
        target_2 = target_1
    target_3 = parsed_target.split(' ')[-1]

    upper_target_1 = target_1.upper()
    upper_target_2 = target_2.upper()
    upper_target_3 = target_3.upper()

    target_components.append(target_1)
    target_components.append(target_2)
    target_components.append(target_3)

    for item in configuration_list:
        upper_item = str(item).upper()
        if upper_target_1 in upper_item:
            possible_configuration.append(item)
        if upper_target_2 in upper_item:
            possible_configuration.append(item)
        if upper_target_3 in upper_item:
            possible_configuration.append(item)

    def list_duplicates(seq):
        seen = set()
        seen_add = seen.add
        # adds all elements it doesn't know yet to seen and all other to seen_twice
        seen_twice = set(x for x in seq if x in seen or seen_add(x))
        # turn the set into a list (as requested)
        return list(seen_twice)

    reduce_configuration = list_duplicates(possible_configuration)

    print(reduce_configuration)

    for item in reduce_configuration:
        upper_item = str(item).upper()
        if gen_number in upper_item:
            return item


def double_check(configuration_list, gen_number, supplier, configuration):
    possible_configuration = []
    upper_generation = gen_number.upper()
    upper_supplier = supplier.upper()
    upper_configuration = configuration.upper()

    for item in configuration_list:
        upper_item = str(item).upper()
        if upper_generation in upper_item:
            possible_configuration.append(item)
        if upper_supplier in upper_item:
            possible_configuration.append(item)
        if upper_configuration in upper_item:
            possible_configuration.append(item)

    if len(possible_configuration) == 3:
        return possible_configuration[0]
    else:
        return 'No Match'


def store_headers(headers):
    headers_list = []
    for item in headers:
        headers_list.append(item)
    return headers_list


def get_the_configuration_row(final_configuration):
    with open('all_gen.csv', 'r') as f:
        for item in f:
            if final_configuration in item:
                return item


def parse_configuration_row(row):
    row_len = len(str(row).split(',')[0]) + 1
    new_row = row[row_len:]
    row_len = len(str(new_row).split(',')[0]) + 1
    new_row = new_row[row_len:]
    row_len = len(str(new_row).split(',')[0]) + 1
    new_row = new_row[row_len:]
    row_len = len(str(new_row).split(',')[0]) + 1
    new_row = new_row[row_len:]

    return new_row


def get_all(just_components, headers_list):
    component_dict = {}

    headers_number = len(headers_list) - 4

    start = 0
    while start < headers_number:
        component_pos = str(just_components).split(',')
        if component_pos[start] != '':
            configuration_number = start + 3
            component = headers_list[configuration_number]
            component_dict[component] = f"{component_pos[start]}"
        start += 1

    return component_dict

    # while begin < len(configuration_list):
    #     component = just_components.split(',')[start]
    #     print(component)
    #     start += 1
    #     begin += 1


def main():
    # target = '[Azure][xStore Storage Utility][Gen 5.6][ZT] (M1150894-001)'
    # target = '[Azure][Compute Optimized][Gen 6.1][Wiwynn]'
    target = '[Azure][ xStore Storage Server][Gen 5.0][ZT] (M1150894-001)'

    configuration_list = create_csv_from_qcl_list()
    gen_number = get_gen_number(target)
    configuration = get_configuration(configuration_list, target, gen_number)
    headers = get_column_headers()
    headers_list = store_headers(headers)
    supplier = get_supplier(target)
    # final_configuration = double_check(configuration_list, gen_number, supplier, configuration)

    configuration_row = get_the_configuration_row(configuration)
    just_components = parse_configuration_row(configuration_row)
    component_dict = get_all(just_components, headers_list)

    print(component_dict)


main()
