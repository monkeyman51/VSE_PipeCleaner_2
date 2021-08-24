"""
Checks for weekly report dealing with presentation sent to the Client.
"""
from pipe_cleaner.src.data_console_server import main_method as get_console_server_data
from pipe_cleaner.src.data_ado import main_method as get_all_ticket_data


def clean_assigned_to_name(assigned_to: str) -> str:
    """
    Assure assigned to name within each TRR is clean and readable.
    """
    if not assigned_to:
        return "None"

    else:
        return assigned_to.replace(".", " ").title()


def get_people_data(ticket_data: dict) -> dict:
    """

    """



def main() -> None:
    """

    """
    console_server_data: dict = get_console_server_data()
    ticket_data: dict = get_all_ticket_data(console_server_data)

    import json
    foo = json.dumps(ticket_data, sort_keys=True, indent=4)
    print(foo)
    input()

    kirkland_report: dict = {}
    people_report: dict = {}
    trr_types: dict = {"Primary": 0,
                       "Secondary": 0}

    for trr_number in ticket_data:
        if str(trr_number).isdigit():
            trr_data: dict = ticket_data[trr_number]
            assigned_to: str = trr_data["assigned_to"]
            state: str = trr_data["state"]
            trr_type: str = trr_data["trr_type"]

            assigned_to: str = clean_assigned_to_name(assigned_to)

            if assigned_to not in people_report:
                people_report[assigned_to] = 1
            else:
                people_report[assigned_to] += 1

            if state not in kirkland_report:
                kirkland_report[state] = 1
            else:
                kirkland_report[state] += 1

            if str(trr_type) == "1":
                trr_types["Primary"] += 1
            elif str(trr_type) == "2":
                trr_types["Secondary"] += 1

    print(f'\nTRR Status: ')
    for segment in kirkland_report:
        value = kirkland_report[segment]
        print(f'\t-{segment}: {value}')

    print(f'\n\nPeople: ')
    for person in people_report:
        count = people_report[person]
        print(f'\t-{person}: {count}')

    print(f'\n\nTRR Types: ')
    for trr_type in trr_types:
        step = trr_types[trr_type]
        print(f'\t-{trr_type}: {step}')

