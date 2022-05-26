"""
Rich request.  Excel output that shows each pipe along with host states.  Omits Virtual Machines
"""
from os import system

from openpyxl import load_workbook

from pipe_cleaner.src.data_ado import main_method as get_all_ticket_data
from pipe_cleaner.src.data_console_server import main_method as get_console_server_data


def get_real_pipes(console_server: dict) -> dict:
    """
    Iterate through console_server.  Return only real pipes based on name
    :param console_server:
    :return:
    """
    real_pipes: dict = {}

    for potential_pipe_name in console_server:
        if "Pipe-" in potential_pipe_name:
            real_pipes[potential_pipe_name]: dict = console_server[potential_pipe_name]

    return real_pipes


def get_blade_count(console_server: dict) -> dict:
    """
    Get blade count for each one.
    :return:
    """
    real_pipes: dict = get_real_pipes(console_server)

    blade_count: dict = {}
    for real_pipe in real_pipes:
        pipe_data: dict = real_pipes[real_pipe]["pipe_data"]

        blade_count[real_pipe]: int = 0
        for blade_name in pipe_data:
            if "-VM-" not in blade_name \
                    and "pipe_inventory" not in blade_name \
                    and "CMA-" not in blade_name:
                print(f"blade_name: {blade_name}")
                blade_count[real_pipe] += 1

    return blade_count


def is_real_blade(blade_name: str) -> bool:
    """
    blade_name
    :param blade_name:
    :return: True is real, False is fake blade
    """
    if "-VM-" not in blade_name \
            and "pipe_inventory" not in blade_name \
            and "CMA-" not in blade_name:
        return True

    else:
        return False


def get_location_from_pipe(raw_pipe_name: str) -> str:
    """
    Extract rack location from Console Server pipe name.
    :param raw_pipe_name:
    :return:
    """
    return raw_pipe_name.split("[")[-1].replace("]", "")


def get_clean_connection(connection: str) -> str:
    """

    :param connection:
    :return:
    """
    if connection == "alive":
        return "On"
    else:
        return "Off"


def get_pipes_data(real_pipes: dict) -> dict:
    """
    Get current pipes data.
    :param real_pipes:
    :return:
    """
    pipes_data: dict = {}

    for real_pipe in real_pipes:
        pipe_data: dict = real_pipes[real_pipe]["pipe_data"]

        pipes_data[real_pipe]: dict = {}
        pipes_data[real_pipe]["location"]: str = get_location_from_pipe(real_pipe)

        total: int = 0
        alive: int = 0
        dead: int = 0

        for blade_name in pipe_data:
            if is_real_blade(blade_name):
                connection: str = pipe_data[blade_name]["connection_status"]

                if connection == "dead":
                    dead += 1
                elif connection == "alive":
                    alive += 1
                elif connection == "mostly_dead":
                    dead += 1

                total += 1

        if total == alive:
            pipes_data[real_pipe]["state"]: str = "All On"

        elif total == dead:
            pipes_data[real_pipe]["state"]: str = "All Off"

        else:
            pipes_data[real_pipe]["state"]: str = "Mix On / Off"

        pipes_data[real_pipe]["count"]: int = total

        owner_name: str = real_pipes[real_pipe]["checked_out_to"]
        pipes_data[real_pipe]["owner"]: str = clean_owner_name(owner_name)
        pipes_data[real_pipe]["tickets"]: str = real_pipes[real_pipe]["group_unique_tickets"]

    return pipes_data


def clean_owner_name(owner_name: str) -> str:
    """

    :param owner_name:
    :return:
    """
    return owner_name.replace(".", " ").title()


def clean_ticket_number(ticket_raw: str) -> str:
    """
    Assures ticket number or None return for consistency.
    :param ticket_raw: ticket number taken per blade within Console Server
    :return: ticket number or None
    """
    if ticket_raw.isdigit():
        return ticket_raw
    else:
        return "None"


def get_blades_data(console_server: dict) -> list:
    """
    Get pipes including count, hosts state.
    :return: list of individual blade data
    """
    real_pipes: dict = get_real_pipes(console_server)
    pipes_data: dict = get_pipes_data(real_pipes)

    blades_data: list = []
    for real_pipe in real_pipes:
        pipe_data: dict = real_pipes[real_pipe]["pipe_data"]

        for blade_name in pipe_data:

            if is_real_blade(blade_name):

                ticket_number: str = pipe_data[blade_name]["ticket"]
                clean_ticket: str = clean_ticket_number(ticket_number)

                current_blade: dict = {"blade_name": blade_name,
                                       "pipe_name": clean_pipe_name(real_pipe),
                                       "pipe_location": pipes_data[real_pipe]["location"],
                                       "pipe_state": pipes_data[real_pipe]["state"],
                                       "pipe_count": pipes_data[real_pipe]["count"],
                                       "tickets": ", ".join(pipes_data[real_pipe]["tickets"]),
                                       "pipe_owner": pipes_data[real_pipe]["owner"],
                                       "blade_state": pipe_data[blade_name]["connection_status"],
                                       "blade_ticket": clean_ticket}

                blades_data.append(current_blade)

    return blades_data


def clean_pipe_name(pipe_name: str) -> str:
    """
    Shorten pipe name to fit into excel output
    :param pipe_name:
    :return:
    """
    clean_data: str = pipe_name. \
        replace('[', ''). \
        replace(']', ''). \
        replace("'", '')

    last_part: str = clean_data.split(' ')[-1]

    return str(clean_data.replace('Pipe-', '').replace(last_part, '')).strip()


def add_pipe_state(index, pipe_name, pipes_data, worksheet):
    """

    :param index:
    :param pipe_name:
    :param pipes_data:
    :param worksheet:
    :return:
    """
    alive_count: int = pipes_data[pipe_name]["alive"]
    dead_count: int = pipes_data[pipe_name]["dead"]
    total_count: int = pipes_data[pipe_name]["count"]

    if total_count == alive_count:
        worksheet[f"B{index}"].value = "All On"

    elif total_count == dead_count:
        worksheet[f"B{index}"].value = "All Off"

    else:
        worksheet[f"B{index}"].value = "Mix On / Off"


def get_ticket_state(ticket_number: str, azure_devops_data: dict) -> str:
    """
    Clean ticket state as TRR from ADO.
    :param ticket_number: raw ticket number / None
    :param azure_devops_data: collected ADO data from available ticket numbers
    :return:
    """
    if ticket_number == "None":
        return "No TRR Assigned"
    else:
        try:
            ticket_state: str = azure_devops_data[ticket_number]["state"]

            clean_state: str = ticket_state.replace('InProgress', 'In Progress'). \
                replace('Test completed', 'Test Completed'). \
                replace('Ready To Review', 'Ready to Review'). \
                replace('Ready to start', 'Ready to Start')

            return clean_state
        except KeyError:
            print(f"\n\n\tTRR {ticket_number} does not exist in ADO.  Please get rid of Ticket Field {ticket_number}.")
            print(f"\tTIP: Use Console Server All Hosts tab to quickly find TRR {ticket_number}")
            input(f"Press enter to exit.")


def get_pipe_states(azure_devops_data: dict, blades_data: list) -> dict:
    """
    Get pipe to ticket (TRR) state / states
    :param azure_devops_data: gathered from Console Server tickets to ADO data
    :param blades_data: per blade data
    :return: pipe to ticket states
    """
    ticket_to_state: dict = {}

    for blade_data in blades_data:

        trr_number: str = blade_data["blade_ticket"]
        ticket_state: str = get_ticket_state(trr_number, azure_devops_data)

        if ticket_state not in ticket_to_state:
            ticket_to_state[trr_number]: list = []
            ticket_to_state[trr_number].append(ticket_state)

        elif ticket_state in ticket_to_state:
            given_states: list = ticket_to_state[trr_number]
            given_states.append(ticket_state)

            unique_states = list(set(given_states))

            ticket_to_state[trr_number]: list = unique_states

    return ticket_to_state


def main() -> None:
    """
    Main Method.
    :return:
    """
    console_server: dict = get_console_server_data()
    azure_devops_data: dict = get_all_ticket_data(console_server)
    blades_data: list = get_blades_data(console_server)

    workbook = load_workbook("settings/inventory/pipes_state_template.xlsx")
    worksheet = workbook["Sheet1"]

    pipe_to_states: dict = get_pipe_states(azure_devops_data, blades_data)

    for index, blade_data in enumerate(blades_data, start=2):

        trr_number: str = blade_data["blade_ticket"]
        ticket_state: str = get_ticket_state(trr_number, azure_devops_data)

        worksheet[f"A{index}"].value = blade_data["pipe_name"]
        worksheet[f"B{index}"].value = blade_data["pipe_location"]
        worksheet[f"C{index}"].value = blade_data["pipe_state"]
        worksheet[f"D{index}"].value = blade_data["pipe_count"]
        worksheet[f"E{index}"].value = blade_data["pipe_owner"]
        worksheet[f"F{index}"].value = blade_data["blade_name"]
        worksheet[f"G{index}"].value = get_clean_connection(blade_data["blade_state"])
        worksheet[f"H{index}"].value = trr_number
        worksheet[f"I{index}"].value = ticket_state
        worksheet[f"J{index}"].value = ", ".join(pipe_to_states[trr_number])
        worksheet[f"K{index}"].value = blade_data["tickets"]

    workbook.save("pipes_state_output.xlsx")
    system(fr'start EXCEL.EXE pipes_state_output.xlsx')
