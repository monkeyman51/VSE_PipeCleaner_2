"""
Rich request.  Excel output that shows each pipe along with host states.  Omits Virtual Machines
"""
from pipe_cleaner.src.data_console_server import main_method as get_console_server_data
from openpyxl import load_workbook
from os import system


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


def get_blades_data(console_server: dict) -> list:
    """
    Get pipes including count, hosts state.
    :return:
    """
    real_pipes: dict = get_real_pipes(console_server)
    pipes_data: dict = get_pipes_data(real_pipes)

    blades_data: list = []
    for real_pipe in real_pipes:
        pipe_data: dict = real_pipes[real_pipe]["pipe_data"]

        for blade_name in pipe_data:
            if is_real_blade(blade_name):

                current_blade: dict = {"blade_name": blade_name,
                                       "pipe_name": clean_pipe_name(real_pipe),
                                       "pipe_location": pipes_data[real_pipe]["location"],
                                       "pipe_state": pipes_data[real_pipe]["state"],
                                       "pipe_count": pipes_data[real_pipe]["count"],
                                       "tickets": ", ".join(pipes_data[real_pipe]["tickets"]),
                                       "pipe_owner": pipes_data[real_pipe]["owner"],
                                       "blade_state": pipe_data[blade_name]["connection_status"]}

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


def main() -> None:
    """
    Main Method.
    :return:
    """
    console_server: dict = get_console_server_data()
    blades_data: list = get_blades_data(console_server)

    workbook = load_workbook("settings/inventory/pipes_state_template.xlsx")
    worksheet = workbook["Sheet1"]

    for index, blade_data in enumerate(blades_data, start=2):
        worksheet[f"A{index}"].value = blade_data["pipe_name"]
        worksheet[f"B{index}"].value = blade_data["pipe_location"]
        worksheet[f"C{index}"].value = blade_data["pipe_state"]
        worksheet[f"D{index}"].value = blade_data["pipe_count"]
        worksheet[f"E{index}"].value = blade_data["pipe_owner"]
        worksheet[f"F{index}"].value = blade_data["blade_name"]
        worksheet[f"G{index}"].value = get_clean_connection(blade_data["blade_state"])
        worksheet[f"H{index}"].value = blade_data["tickets"]

    workbook.save("pipes_state_output.xlsx")
    system(fr'start EXCEL.EXE pipes_state_output.xlsx')
