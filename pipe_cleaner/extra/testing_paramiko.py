from paramiko import SSHClient, AutoAddPolicy, channel


show_manager_info_data: list = ['Manager Uptime',
                                'Firmware Version',
                                'Manufacturer',
                                'FW Version',
                                'Host Name',
                                'IP Address',
                                'Power',
                                'Humidity',
                                'Temperature']


# def parse_rack_manager(ssh_output: str, component: str) -> dict:
#     """
#     Parse information from rack manager
#     :param ssh_output: line of SSH from Rack Manager
#     :param component: individual component of data
#     :return:
#     """
#     if ssh_output == "" or ssh_output is None:
#         continue
#
#     if "FW Version:" in ssh_output:
#         parsed_fw_version = str(ssh_output).replace("FW Version:", "").replace(" ", "")
#         rack_manager_data["fw_version"] = parsed_fw_version
#
#     if "Host Name:" in ssh_output:
#         parsed_host_name = str(ssh_output).replace("Host Name:", "").replace(" ", "")



# def data_from_rack_manager(stdout: channel.ChannelFile, ssh_line: str) -> dict:
#     """
#     System output from Rack Manager data
#     :param stdout: SSH output via Rack Manager
#     :param ssh_line: command_line from Rack Manager SSH
#     :return:
#     """
#     rack_manager_data = {}
#
#     initial = 0
#     while initial < 200:
#         ssh_output = stdout.readline()
#
#         for component in show_manager_info_data:
#             rack_manager_data[]


def testing_paramiko(credentials: dict, port_number: int) -> None:
    """
    Attempting to make SSH connection work
    :param credentials: Contains console server and rack manager's usernames, passwords, and ip addresses
    :param port_number: Should be 22
    :return: None
    """
    # Unpack Credentials Dictionary
    console_server: dict = credentials['console_server']
    rack_manager: dict = credentials['rack_manager']

    command_1 = 'wcscli'
    command_2 = 'show manager info'

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
    virtual_machine_channel = virtual_machine_transport.open_channel('direct-tcpip',
                                                                     destination_address,
                                                                     source_address)
    # scp = SCPClient(virtual_machine_transport)

    # Log SSH Console Server
    virtual_machine_status = virtual_machine.get_transport().is_active()
    print(f'\tVirtual Machine Transport Status: {virtual_machine_status}')

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
    print(f'\tMachine Host Status: {machine_host_status}')

    stdin, stdout, stderr = machine_host.exec_command('show manager info', get_pty=True)

    initial = 0
    while initial < 200:
        feedback = stdout.readline()
        if feedback == "" or feedback is None:
            pass
        else:
            print(feedback)

        if stdout.channel.exit_status_ready():
            break

        initial += 1

    machine_host.close()
    virtual_machine.close()


def main():
    # Setting up login credentials for paramiko
    credentials: dict = {'console_server': {}, 'rack_manager': {}}
    credentials['console_server']['username'] = 'joe.ton'
    credentials['console_server']['password'] = 'kn1f3loc321'
    credentials['console_server']['ip_address'] = '172.30.1.100'
    credentials['rack_manager']['username'] = 'root'
    credentials['rack_manager']['password'] = '$pl3nd1D'
    # credentials['rack_manager']['ip_address'] = '192.168.237.181'
    credentials['rack_manager']['ip_address'] = '192.168.0.16'  # Worked
    # credentials['rack_manager']['ip_address'] = '192.168.0.16'

    port_number = 22

    testing_paramiko(credentials, port_number)


if __name__ == '__main__':
    main()
