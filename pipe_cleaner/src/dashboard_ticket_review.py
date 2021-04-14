
def check_data_table(toggle_component: str, table_data: dict, toggle_data: dict, toggle_keys: list):
    """
    Check for None returns from table data
    :param toggle_data:
    :param toggle_component:
    :param table_data:
    :param toggle_keys:
    :return:
    """
    toggle_pair: dict = toggle_data[toggle_component]

    get_table_data: dict = {}
    for key in toggle_keys:
        toggle_status: str = toggle_pair[key]

        if toggle_status == 'REQUIRED':
            get_table_data[key] = table_data.get(key)
            # print(f'\t- {key}: {table_data.get(key)}')
        elif toggle_status == 'ON':
            get_table_data[key] = table_data.get(key)
            # print(f'\t- {key}: {table_data.get(key)}')
        elif toggle_status == 'OFF':
            pass
        # Unintentional input from excel
        else:
            pass

    return get_table_data


def main_method(all_ticket_data: dict, toggle_data: dict):
    """
    For the PM column in the main dashboard
    :param all_ticket_data: TRR
    :param toggle_data: Which information to fetch
    :return:
    """

    # Unpack Toggle Keys based on Component
    dimm_toggle = list(toggle_data.get('DIMM').keys())
    nvme_toggle = list(toggle_data.get('NVME').keys())
    hdd_toggle = list(toggle_data.get('HDD').keys())
    ssd_toggle = list(toggle_data.get('SSD').keys())

    data_extracted: dict = {}

    try:
        for ticket in all_ticket_data:

            ticket_all_data: dict = {}

            # Ticket Data, reduce size of code
            table_data: dict = all_ticket_data[ticket]['table_data']

            # Raise for comparison, know which toggle component to pull from
            table_request_type = str(table_data.get('request_type')).upper()

            if 'DIMM' in table_request_type:
                ticket_data = check_data_table('DIMM', table_data, toggle_data, dimm_toggle)
                data_extracted[ticket] = ticket_data
            elif 'NVME' in table_request_type:
                ticket_data = check_data_table('NVME', table_data, toggle_data, nvme_toggle)
                data_extracted[ticket] = ticket_data
            elif 'HDD' in table_request_type:
                ticket_data = check_data_table('HDD', table_data, toggle_data, hdd_toggle)
                data_extracted[ticket] = ticket_data
            elif 'SSD' in table_request_type:
                ticket_data = check_data_table('SSD', table_data, toggle_data, ssd_toggle)
                data_extracted[ticket] = ticket_data
    except TypeError:
        pass

    return all_ticket_data
