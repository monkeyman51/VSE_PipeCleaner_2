"""
Intended for proof of concept for accessing Rack Manager, sending commands, and retrieving text back as response.

Program starts at main() at the bottom of the file.
"""


from paramiko import SSHClient, AutoAddPolicy, ssh_exception
from requests import post
from json import loads
from datetime import datetime, date
from openpyxl import load_workbook
from os import system
from openpyxl.styles import Alignment


def build_command_result(dhcp_name: str, rack_ip_address: str) -> dict:
    """
    Construct command result for later parsing.
    """
    current_time: str = datetime.now().strftime("%H:%M:%S")
    current_date = str(date.today())

    return {"dhcp_name": dhcp_name,
            "ip_address": rack_ip_address,
            "current_time": current_time,
            "current_date": current_date,
            "group": get_dhcp_group(dhcp_name),
            "max_power": "None",
            "power_drawn": "None",
            "real_power": "None"}


def get_dhcp_group(dhcp_name: str) -> str:
    """

    """
    if "KRK-RB" in dhcp_name or "KRK-RC" in dhcp_name or "-RB" in dhcp_name or "-RC" in dhcp_name:
        return "Einstein"
    elif "KRK-RD" in dhcp_name or "-RD" in dhcp_name:
        return "DaVinci"
    else:
        return "Other"


def send_rack_manage_command(credentials: dict, command: str, dhcp_name: str) -> dict:
    """
    Attempting to make SSH connection work

    :param dhcp_name: Dynamic Host Configuration Protocol
    :param credentials: Contains console server and rack manager's usernames, passwords, and ip addresses
    :param command:
    """
    console_server: dict = credentials['console_server']
    rack_manager: dict = credentials['rack_manager']
    port_number: int = credentials['port_number']
    rack_ip_address: str = rack_manager.get('ip_address')

    result: dict = build_command_result(dhcp_name, rack_ip_address)

    # Login Credentials to Console Server network
    virtual_machine = SSHClient()
    virtual_machine.set_missing_host_key_policy(AutoAddPolicy())
    virtual_machine.connect(console_server.get('ip_address'),
                            port=port_number,
                            username=console_server.get('username'),
                            password=console_server.get('password'))

    # Set up channel
    virtual_machine_transport = virtual_machine.get_transport()
    destination_address = (rack_manager.get('ip_address'), port_number)
    source_address = (console_server.get('ip_address'), port_number)

    # Errors out mean unavailable
    try:
        virtual_machine_channel = virtual_machine_transport.open_channel('direct-tcpip',
                                                                         destination_address,
                                                                         source_address)

        # Log SSH Console Server
        virtual_machine_status = virtual_machine.get_transport().is_active()

        # Login Credentials to Rack Manager network
        machine_host = SSHClient()
        machine_host.set_missing_host_key_policy(AutoAddPolicy())
        machine_host.connect(rack_manager.get('ip_address'),
                             port=port_number,
                             username=rack_manager.get('username'),
                             password=rack_manager.get('password'),
                             sock=virtual_machine_channel)

        # Log SSH Rack Manager
        machine_host_status = machine_host.get_transport().is_active()
        print(f'\t\t- Response Received')

        command_response: list = get_command_response(command, machine_host)

        for line in command_response:
            if "MaxPowerInWatts" in line:
                clean_line: str = clean_response_line(line)
                result["max_power"]: str = clean_line

            elif "PowerDrawnInWatts" in line:
                clean_line: str = clean_response_line(line)
                result["power_drawn"]: str = clean_line

            elif "RealPowerInWatts" in line:
                clean_line: str = clean_response_line(line)
                result["real_power"]: str = clean_line

        machine_host.close()
        virtual_machine.close()

        return result

    except ssh_exception.ChannelException:
        print(f"\t\t- Failed")
        return result

    except ssh_exception.AuthenticationException:
        print(f"\t\t- Failed")
        return result


def clean_response_line(line_response: str) -> str:
    """
    Given current line from the command sent to the RM, clean data left to obtain only relevant infomration.

    For this case,sh manager powermeter reading command for RM
    :param line_response:
    :return:
    """
    return line_response.replace("MaxPowerInWatts: ", "").\
        replace("PowerDrawnInWatts", "").\
        replace("RealPowerInWatts", "").\
        replace("    : ", "").replace(r"\r\n", "").replace("\r\n", "").replace("    ", "")


def get_command_response(command: str, machine_host) -> list:
    """
    Grabs all text from command sent to Rack Manager.
    """
    all_text: list = []
    stdin, stdout, stderr = machine_host.exec_command(command, get_pty=True)
    initial = 0
    while initial < 5_000:
        feedback = stdout.readline()
        if feedback == "" or feedback is None:
            pass
        else:
            all_text.append(feedback)

        if stdout.channel.exit_status_ready():
            break

        initial += 1

    return all_text


def get_credentials(rack_manager_ip_address: str) -> dict:
    """
    Get Console Server and Rack manager basic information to get access Console Server.
    """
    # Personal Information
    return {'console_server': {'username': 'joe.ton',
                               'password': 'kn1f3loc321',
                               'ip_address': '172.30.1.100'},

            'rack_manager': {'username': 'root',
                             'password': '$pl3nd1D',
                             'ip_address': rack_manager_ip_address},

            'port_number': 22}


def get_dhcp_data():
    """
    Get Dynamic Host Configuration Protocol from Console Server page containing the rack manager information
    :return:
    """
    data: dict = {
        'action': 'get_reservations',
    }
    response = post(url=f'http://172.30.1.100/console/console_js.php', json=data)
    return loads(response.text)


def output_excel(results: list) -> None:
    """
    Output excel given
    """
    workbook = load_workbook("settings/rack_managers_power_template.xlsx")
    worksheet = workbook["Sheet1"]

    for index, blade_data in enumerate(results, start=2):

        worksheet[f"A{index}"].value = blade_data["dhcp_name"]
        worksheet[f"B{index}"].value = blade_data["ip_address"]
        worksheet[f"C{index}"].value = blade_data["group"]
        worksheet[f"D{index}"].value = blade_data["current_date"]
        worksheet[f"E{index}"].value = blade_data["current_time"]
        worksheet[f"F{index}"].value = blade_data["max_power"]
        worksheet[f"G{index}"].value = blade_data["power_drawn"]
        worksheet[f"H{index}"].value = blade_data["real_power"]

    workbook.save("rack_managers_power.xlsx")
    system(fr'start EXCEL.EXE rack_managers_power.xlsx')


def store_excel_data(position: str, key_name: str, worksheet, blade_data: dict) -> None:
    """

    :param position:
    :param index:
    :param key_name:
    :return:
    """
    value = blade_data.get(key_name)

    if type(value) == str:
        worksheet[position].value = value

    elif type(value) == int:
        worksheet[position].value = value


def main():
    dhcp_nodes: dict = get_dhcp_data()
    command: str = "sh manager powermeter reading"

    results: list = []

    for dhcp_node in dhcp_nodes:
        dhcp_name: str = dhcp_node.get("name", "")
        ip_address: str = dhcp_node.get("ip", "")

        print(f"\t{dhcp_name} | {ip_address} - {command}")
        print(f"\t\t- Sending Command")

        if dhcp_name and ip_address:
            credentials: dict = get_credentials(ip_address)
            command_response: dict = send_rack_manage_command(credentials, command, dhcp_name)
            results.append(command_response)

    output_excel(results)


if __name__ == '__main__':
    import time

    start = time.time()
    main()
    end = time.time()
    print(end - start)
