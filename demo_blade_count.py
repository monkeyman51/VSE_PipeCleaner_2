"""
Rich request.  Excel output that shows each pipe along with blade count for each.
"""
from pipe_cleaner.src.data_console_server import main_method as get_console_server_data
from openpyxl import load_workbook


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


def main() -> None:
    """
    Main Method.
    :return:
    """
    console_server: dict = get_console_server_data()
    blade_count: dict = get_blade_count(console_server)

    workbook = load_workbook("settings/inventory/blade_count.xlsx")
    worksheet = workbook["Sheet1"]

    for index, pipe_name in enumerate(blade_count, start=2):
        worksheet[f"A{index}"].value = clean_pipe_name(pipe_name)
        worksheet[f"B{index}"].value = blade_count[pipe_name]

    workbook.save("blade_count_output.xlsx")


main()
