# from pipe_cleaner.src.data_access import write_host_groups_json
# from pipe_cleaner.src.data_access import get_all_host_ids, request_ado_json
# from pipe_cleaner.src.data_access import generate_console_server_json, get_console_server_json
# import json
# import requests
# import aiohttp
# import asyncio
#
# tickets_in_console_server: list = []
#
#
# def get_unique_tickets_console_server() -> list:
#     """
#     Clean tickets and return only unique tickets
#     :return:
#     """
#     unique_tickets: list = []
#     for ticket in tickets_in_console_server:
#         if ticket.isdigit():
#             unique_tickets.append(ticket)
#
#     return list(set(unique_tickets))
#
#
# def get_host_groups_data(host_groups) -> dict:
#     """
#     Get all host group data
#     :return:
#     """
#     dashboard_data = {}
#
#     print(f'\n\t**** Please WAIT - Gathering LOTS OF DATA for Main Dashboard ****\n')
#     input(f'\tClose dashboard.xlsx if open. Press Enter to continue...')
#
#     print(f'\n\tGathering Data:')
#
#     for host_group_name in host_groups:
#         if 'Pipe-' in host_group_name['name'] and 'OFFLINE' not in host_group_name['comment'] \
#                 and '[' in host_group_name['name'] and ']' in host_group_name['name']:
#             pipe_name = host_group_name.get('name')
#             description = host_group_name.get('description')
#             host_id = host_group_name.get('id')
#             host_ids = host_group_name.get('host_ids')
#             status = host_group_name.get('comment')
#             checked_out_to = host_group_name.get('checked_out_to')
#
#             dashboard_data[pipe_name] = {}
#             current_pipe = dashboard_data[pipe_name]
#
#             print(f'\t- Current Pipe: {pipe_name}')
#
#             pipe_data = get_pipe_data(host_ids, pipe_name)
#             current_pipe['pipe_data'] = pipe_data
#
#             current_pipe['description'] = description
#             current_pipe['host_group_status'] = status
#             current_pipe['host_id'] = host_id
#             current_pipe['host_ids'] = host_ids
#             current_pipe['checked_out_to'] = checked_out_to
#
#             group_unique_tickets = get_group_unique_tickets(host_id, pipe_name)
#             due_dates: dict = access_due_dates(group_unique_tickets)
#
#             current_pipe['group_unique_tickets'] = group_unique_tickets
#             current_pipe['due_dates'] = due_dates
#
#     return dashboard_data
#
#
# def get_group_unique_tickets(host_id: str, pipe_name: str) -> list:
#     """
#     Get the Host Groups unique tickets for later extraction
#     :param pipe_name:
#     :param host_id: Host Group ID
#     :return:
#     """
#     host_group_unique_tickets: list = []
#
#     host_group_json = get_all_host_ids(host_id, pipe_name)
#
#     for system in host_group_json:
#         ticket = str(system['ticket'])
#
#         if ticket.isdigit() is True:
#             host_group_unique_tickets.append(ticket)
#
#     unique_tickets = list(set(host_group_unique_tickets))
#     # print(f'Host Group Unique Tickets: ( {pipe_name} ) - {unique_tickets}')
#
#     return unique_tickets
#
#
# def get_pipe_data(host_ids, pipe_name) -> dict:
#     all_host_ids: dict = {}
#
#     for host_id in host_ids:
#         product_serial = generate_console_server_json(host_id)
#         get_console_server_json(product_serial, host_id)
#
#         host_json_path = f'pipe_cleaner/data/{host_id}.json'
#
#         try:
#             # Console Server
#             with open(host_json_path, 'r') as f:
#                 console_server_host_json = json.loads(f.read())
#                 host_data = system_decoder(console_server_host_json)
#                 machine_name = host_data['machine_name']
#
#                 # Account for mis-capitalization in Machine Name
#                 upper_machine_name = str(machine_name).upper()
#                 if upper_machine_name == 'NONE' or upper_machine_name is None or '-VM-' in upper_machine_name:
#                     pass
#                 elif 'VSE' in upper_machine_name:
#                     all_host_ids[upper_machine_name] = host_data
#         except json.decoder.JSONDecodeError:
#             pass
#
#     return all_host_ids
#
#
# def system_decoder(console_server_json: dict) -> dict:
#     """
#     Decodes and console server host
#     :param console_server_json:
#     :return:
#     """
#     system_data: dict = {}
#
#     try:
#         machine_name = console_server_json['machine_name']
#         system_data['machine_name'] = machine_name
#     except KeyError:
#         system_data['machine_name'] = 'None'
#     except TypeError:
#         system_data['machine_name'] = 'None'
#
#     try:
#         ticket = str(console_server_json['ticket'])
#         system_data['ticket'] = ticket
#
#         if ticket != '' and ticket.isdigit() is True:
#             tickets_in_console_server.append(ticket)
#
#     except KeyError:
#         system_data['ticket'] = 'None'
#     except TypeError:
#         system_data['ticket'] = 'None'
#
#     try:
#         # system_bios = console_server_json['dmi']['bios']['version'][-8:]
#         system_bios = console_server_json['dmi']['bios']['version']
#         system_data['server_bios'] = system_bios
#     except KeyError:
#         system_data['server_bios'] = 'None'
#     except TypeError:
#         system_data['server_bios'] = 'None'
#
#     try:
#         system_bios = console_server_json['dmi']['procs']
#         system_data['processors'] = system_bios
#     except KeyError:
#         system_data['processors'] = 'None'
#     except TypeError:
#         system_data['processors'] = 'None'
#
#     try:
#         system_bmc = console_server_json['bmc']['mc']['firmware']
#         # system_data['server_bmc'] = str(system_bmc).replace('.', '')
#         system_data['server_bmc'] = str(system_bmc)
#     except KeyError:
#         system_data['server_bmc'] = 'None'
#     except TypeError:
#         system_data['server_bmc'] = 'None'
#
#     try:
#         system_cpld = console_server_json['cpld']['secure_cpld_version']
#         # system_data['server_cpld'] = str(system_cpld)[-2:]
#         system_data['server_cpld'] = str(system_cpld)
#     except KeyError:
#         system_data['server_cpld'] = 'None'
#     except TypeError:
#         system_data['server_cpld'] = 'None'
#
#     try:
#         system_cpld = console_server_json['cpld']['sequence_cpld_version']
#         # system_data['server_cpld'] = str(system_cpld)[-2:]
#         system_data['sequence_cpld_version'] = str(system_cpld)
#     except KeyError:
#         system_data['sequence_cpld_version'] = 'None'
#     except TypeError:
#         system_data['sequence_cpld_version'] = 'None'
#
#     try:
#         # system_os = console_server_json['platform']['version'][-5:]
#         system_os = console_server_json['platform']['version']
#         system_data['server_os'] = system_os
#     except KeyError:
#         system_data['server_os'] = 'None'
#     except TypeError:
#         system_data['server_os'] = 'None'
#
#     try:
#         system_tpm = console_server_json['tpm']['version']
#         # system_data['server_tpm'] = str(system_tpm).replace('V', '').replace('v', '')[0:2]
#         system_data['server_tpm'] = str(system_tpm)
#     except KeyError:
#         system_data['server_tpm'] = 'None'
#     except TypeError:
#         system_data['server_tpm'] = 'None'
#
#     try:
#         system_status = console_server_json['status']
#         system_data['system_status'] = system_status
#     except KeyError:
#         system_data['system_status'] = 'None'
#     except TypeError:
#         system_data['system_status'] = 'None'
#
#     try:
#         host_id = console_server_json['id']
#         system_data['host_id'] = host_id
#     except KeyError:
#         system_data['host_id'] = 'None'
#     except TypeError:
#         system_data['host_id'] = 'None'
#
#     try:
#         location = console_server_json['location']
#         system_data['location'] = location
#     except KeyError:
#         system_data['location'] = 'None'
#     except TypeError:
#         system_data['location'] = 'None'
#
#     try:
#         host_id = console_server_json['host_id']
#         system_data['host_id'] = host_id
#     except KeyError:
#         system_data['host_id'] = 'None'
#     except TypeError:
#         system_data['host_id'] = 'None'
#
#     try:
#         comment = console_server_json['comment']
#         system_data['comment'] = comment
#     except KeyError:
#         system_data['comment'] = 'None'
#     except TypeError:
#         system_data['comment'] = 'None'
#
#     try:
#         username = console_server_json['username']
#         system_data['username'] = username
#     except KeyError:
#         system_data['username'] = 'None'
#     except TypeError:
#         system_data['username'] = 'None'
#
#     try:
#         system_dimms = console_server_json['dmi']['dimms']
#         system_data['system_dimms'] = system_dimms
#     except KeyError:
#         system_data['system_dimms'] = 'None'
#     except TypeError:
#         system_data['system_dimms'] = 'None'
#
#     try:
#         unique_dimms = console_server_json['dmi']['unique_dimms']
#         system_data['unique_dimms'] = unique_dimms
#     except KeyError:
#         system_data['unique_dimms'] = 'None'
#     except TypeError:
#         system_data['unique_dimms'] = 'None'
#
#     try:
#         system_nvmes = console_server_json['nvme']['nvmes']
#         system_data['system_nvmes'] = system_nvmes
#     except KeyError:
#         system_data['system_nvmes'] = 'None'
#     except TypeError:
#         system_data['system_nvmes'] = 'None'
#
#     try:
#         unique_nvmes = console_server_json['nvme']['unique_nvmes']
#         system_data['unique_nvmes'] = unique_nvmes
#     except KeyError:
#         system_data['unique_nvmes'] = 'None'
#     except TypeError:
#         system_data['unique_nvmes'] = 'None'
#
#     try:
#         system_disks = console_server_json['disk']['disks']
#         system_data['system_disks'] = system_disks
#     except KeyError:
#         system_data['system_disks'] = 'None'
#     except TypeError:
#         system_data['system_disks'] = 'None'
#
#     try:
#         unique_disks = console_server_json['disk']['unique_disks']
#         system_data['unique_disks'] = unique_disks
#     except KeyError:
#         system_data['unique_disks'] = 'None'
#     except TypeError:
#         system_data['unique_disks'] = 'None'
#
#     return system_data
#
#
# def ticket_decoder(component: str, ticket_data: dict) -> str:
#     """
#     Gets value based on component given from TRR ID. Cleans data for later comparison to Console Server
#     :param component:
#     :param ticket_data:
#     :return:
#     """
#
#     if 'Server BIOS' in component:
#         server_bios = ticket_data.get('server_bios')
#         return server_bios[-8:]
#
#     elif 'Server BMC' in component:
#         server_bmc = ticket_data.get('server_bmc')
#         return (server_bmc.replace('.00', '')[-4:])[-3:]
#
#     elif 'Server CPLD' in component:
#         server_cpld = ticket_data.get('server_cpld')
#         for index, character in enumerate(server_cpld):
#             if 'V' in character:
#                 parsed_component = character + server_cpld[index + 1] + server_cpld[index + 2]
#                 return str(parsed_component).replace('V', '').replace('v', '')
#
#     elif 'Server TPM' in component:
#         server_tpm = ticket_data.get('server_tpm')
#         return str(server_tpm).replace('V', '').replace('v', '')[0:2]
#
#     elif 'Server OS' in component:
#         server_os = ticket_data.get('server_os')
#         if '17763' in server_os:
#             return '2019'
#         elif '2019' in server_os:
#             return '2019'
#         else:
#             return server_os
#
#     else:
#         return component
#
#
# def access_due_dates(unique_tickets) -> dict:
#     """
#     Get actual due date from ADO
#     :param unique_tickets:
#     :return:
#     """
#     due_dates: dict = {}
#
#     for ticket in unique_tickets:
#         json_file = request_ado_json(ticket)
#
#         try:
#             expected_task_start = json_file['fields']['AzureCSI-V1.1.ExpectedTaskStart']
#             due_dates['expected_task_start'] = expected_task_start
#         except KeyError:
#             pass
#
#         try:
#             expected_task_completion = json_file['fields']['AzureCSI-V1.1.ExpectedTaskCompletion']
#             due_dates['expected_task_completion'] = expected_task_completion
#         except KeyError:
#             pass
#
#         try:
#             actual_qual_start_date = json_file['fields']['Custom.ActualQualStartDate']
#             due_dates['actual_qual_start_date'] = actual_qual_start_date
#         except KeyError:
#             pass
#
#         try:
#             actual_qual_end_date = json_file['fields']['Custom.ActualQualEndDate']
#             due_dates['actual_qual_end_date'] = actual_qual_end_date
#         except KeyError:
#             pass
#
#     return due_dates
#
#
# def get_all_ticket_data():
#     all_tickets_data: dict = {}
#
#     for ticket in tickets_in_console_server:
#         print(f'Ticket: {ticket}')
#         ticket_data: dict = {}
#         all_tickets_data[ticket] = ticket_data
#
#         ticket_json = request_ado_json(ticket)
#
#         # BIOS
#         # BMC
#         # CPLD
#         # OS System
#         # TPM
#         # Due Dates
#
#
# async def main():
#     async with aiohttp.ClientSession() as session:
#         async with session.get('http://httpbin.org/get') as resp:
#             print(resp.status)
#             print(await resp.text())
#
#
# def main_method():
#     """
#     Get All relevant data from Console Server based on Host Groups
#     :return:
#     """
#
#     # Create single Host Group data
#     file_path = 'pipe_cleaner/data/all_host_groups.json'
#     json_file = write_host_groups_json(file_path)
#     host_groups = json_file['host_groups']
#
#     event_loop = asyncio.get_event_loop()
#     event_loop.run_until_complete(main())
#     event_loop.close()
#
#     return get_host_groups_data(host_groups)
