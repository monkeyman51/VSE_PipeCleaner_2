import pandas as pd


def access_crd(ticket_number: str, file_path: str, sheet_name: str):
    """
    Access inventory via local file within Pipe Cleaner.
    WARNING: Must update local excel file in order to be up to date
    :param ticket_number: TTR ID
    :param sheet_name: commodity inventory
    :param file_path: file path of inventory maintained by Traci, Bruce, or inventory person
    :return:
    """

    commodity_inventory = pd.read_excel(f'{file_path}', sheet_name=f'{sheet_name}')
    commodity_inventory.to_csv(f'crd_{ticket_number}.csv', index=False)

    # pd.read_csv(f'crd_{ticket_number}.csv', sep=',', header=None, chunksize=1)
    pd.read_csv(f'crd_{ticket_number}.csv')

    return f'crd_{ticket_number}.csv'


# def main(ticket_to_crd: dict, inventory):
def main():
    """
    Scans through CRD
    :return:
    """
    csv_match = []

    ticket_number = '111111'
    sheet_name = 'FW-SW Configuration'

    inventory = {'item_type': {'M393A8G40AB2-CWE': 'DIMM', 'M393A4G40AB3-CVF': 'DIMM', 'M393A4G40AB3-CWE': 'DIMM', 'HMA42GR7AFR4N-TF': 'DIMM', 'HMN82GR7AFR4N-UH': 'DIMM', 'HMA84GR7AFR4N-UH': 'DIMM', 'HMA82GR7AFR8N-VK': 'DIMM', 'HMA84GR7AFR4N-VK': 'DIMM', 'HMAA8GL7AMR4N-VK': 'DIMM', 'HMAA4GR7AJR4N-WM (T8, T4, TG)': 'DIMM', 'HMAA8GR7AJR4N-WM': 'DIMM', 'M391A2K43BB1-CTD': 'DIMM', 'M393A2K40CB2-CVF (2015/2016)': 'DIMM', 'M393A2K40CB2-CTD': 'DIMM', 'M393A4K40CB2-CVF (2015/2016)': 'DIMM', 'M393A4K40CB2-CTD (2006)': 'DIMM', 'M393A4K40CB2-CTD (1915)': 'DIMM', 'M393A4K40CB2-CTD (1902/1903))': 'DIMM', 'HMA82GR7CJR4N-VK': 'DIMM', 'HMA84GR7CJR4N-VK': 'DIMM', 'HMAA8GL7CPR4N-WM': 'DIMM', 'M393A2K40DB2-CVF (2035)': 'DIMM', 'M393A4K40DB2-CVF (2001)': 'DIMM', 'M393A4K40DB2-CVF (2030)/(2035)': 'DIMM', 'M393A2K40DB3-CWE': 'DIMM', 'M393A4K40DB3-CWE': 'DIMM', 'HMA82GR7DJR4N-WM': 'DIMM', 'HMA84GR7DJR4N-WM': 'DIMM', 'HMA82GR7JJR4N-VK': 'DIMM', 'HMA84GR7JJR4N-VK': 'DIMM', 'HMA84GR7MFR4N-UH': 'DIMM', 'NT16GA72D4PBX3P-HR (old)': 'DIMM', 'NT16GA72D4PBX3P-IX': 'DIMM', 'NT32GA72D4NBX3P-HR': 'DIMM', 'NT32GA72D4NBX3P-IX': 'DIMM', 'NT32GA72D4NXA3P-HR(old)': 'DIMM', 'MTA36ASF2G72PZ-2G6E1': 'DIMM', 'MTA18ASF2G72PDZ-2G6D1': 'DIMM', 'MTA36ASF2G72PZ-2G6F1': 'DIMM', 'MTA18ASF2G72PDZ-2G6E1': 'DIMM', 'MTA18ASF2G72PZ-2G9E1': 'DIMM', 'MTA36ASF4G72PZ-2G6E1': 'DIMM', 'MTA36ASF4G72PZ-2G9E2': 'DIMM', 'MTA18ASF2G72PZ-2G6J1': 'DIMM', 'MTA18ASF2G72PZ-2G9J3': 'DIMM', 'MTA36ASF4G72PZ-2G6J1': 'DIMM', 'MTA18ASF4G72PZ-2G6B1': 'DIMM', 'MTA18ASF4G72PZ-2G9B1': 'DIMM', 'MTA36ASF8G72PZ-2G9B2': 'DIMM', 'MTA36ASF8G72PZ-3G2B2': 'DIMM', 'MTA18ASF4G72PZ-2G9E1': 'DIMM', 'MTA18ASF4G72PZ-3G2E1': 'DIMM', 'MTA36ASF8G72PZ-2G9E1': 'DIMM', 'MTA36ASF8G72PZ-3G2E1': 'DIMM', 'MTA36ASF4G72PZ-3G2J3': 'DIMM', 'MTA18ASF2G72PZ-3G2J3': 'DIMM', 'MS16D432R22S8MEX': 'DIMM', 'MS32D432R22S4MEX': 'DIMM', '0F27479': 'HDD', '0F29630': 'HDD', '0F29866': 'HDD', '0F34623': 'HDD', '0F38314': 'HDD', '0F14688': 'HDD', '1TT101-002': 'HDD', '1TT101-401': 'HDD', '2MU103-402': 'HDD', '2K2101-401': 'HDD', '2K2101-402': 'HDD', '2HZ100-401': 'HDD', '1V4107-002': 'HDD', '2RM102-402': 'HDD', '2KG103-401': 'HDD', '2K8122-402': 'HDD', '2LQ202-403': 'HDD', '2KH103-402': 'HDD', '2RK202-401': 'HDD', '2MQ101-402': 'HDD', '3AY212-401': 'HDD', 'HDEPV10SMA51': 'HDD', 'HDEPR03GEA51': 'HDD', 'HDEPR01SMA51': 'HDD', 'HDEPW21SMA51': 'HDD', '0F31114': 'HDD', '0F38313': 'HDD', '0B35950': 'HDD', 'AD2-KW960 ': 'M.2', 'AD2-KW960': 'M.2', 'EP2-KB960': 'M.2', 'EP3-KW960': 'M.2', 'EP4-KW960': 'M.2', 'EPX-KW960': 'M.2', 'SSDPELKX019T8D': 'M.2', 'SSDPELKX960G8D-201': 'M.2', 'SSDPELKX960G8D-203': 'M.2', 'SSDPELKX038T8D': 'M.2', 'SSDPELKX020T8D-201': 'M.2', 'HFS960GD0TEG-6410A': 'M.2', 'HFS1T9GD0FEH-6410A BA': 'M.2', 'HFS960GD0FEG-A430A (Purple dot; common PSID)': 'M.2', 'HFS1T9GD0FEH-A430A': 'M.2', 'HFS3T8GD0FEH-A430A': 'M.2', 'HFS1T9GD0FEI-A430A (new)': 'M.2', 'HFS3T8GD0FEI-A430A': 'M.2', 'HFS960GD0FEI-A430A (Purple dot: Common PSID)': 'M.2', 'MZ1LV960HCJH-000MU': 'M.2', 'MZ1LW1T9HMLS-000MV': 'M.2', 'MZ1LW960HMJP-000MV': 'M.2', 'MZ1LB1T9HALS-00AMV': 'M.2', 'MZ1LB1T9HALS-000MV': 'M.2', 'MZ1LB3T8HMLA-000MV': 'M.2', 'MZ1LB960HAJQ-000MV': 'M.2', 'MZ1LB960HAJQ-00AMV': 'M.2', 'MZ1LB960HBJR-00AMV': 'M.2', 'MZ1LB1T9HBLS-00AMV': 'M.2', 'MZ1LB3T8HALA-00AMV': 'M.2', 'MZ1L2960HCJR-00AMV': 'M.2', 'MZ1L21T9HCLS-00AMV': 'M.2', 'KXD5DLN13T84': 'M.2', 'M1113049-001': 'M.2', 'MZELB15THMLA-000MV': 'RULER', 'SSDPEXNV153T8M2': 'RULER', '0TS2003': 'RULER', 'HFS15T3DGLX070N': 'RULER', 'MTFDDAK960TCB': 'SSD', 'MTFDDAK960TCC': 'SSD', 'MTFDDAK960TDD': 'SSD', 'MZ7LM960HCHP-000MV': 'SSD', 'MZ7LM960HMJP-000MV': 'SSD', 'MZ7LH960HAJR-00AMV': 'SSD', 'MZ7LH960HAJR-000MV': 'SSD', 'HFS960G32MED-3410A': 'SSD', 'HFS960G32MFD-3410A': 'SSD', 'HFS960G32FEH-7A10A': 'SSD', 'MZ7WD960HMHP-00003': 'SSD', 'MZPLJ6T4HALA-00AMV': 'SSD'}, 'item_supplier': {'M393A8G40AB2-CWE': 'Samsung', 'M393A4G40AB3-CVF': 'Samsung', 'M393A4G40AB3-CWE': 'Samsung', 'HMA42GR7AFR4N-TF': 'SK Hynix', 'HMN82GR7AFR4N-UH': 'SK Hynix', 'HMA84GR7AFR4N-UH': 'SK Hynix', 'HMA82GR7AFR8N-VK': 'SK Hynix', 'HMA84GR7AFR4N-VK': 'SK Hynix', 'HMAA8GL7AMR4N-VK': 'SK Hynix', 'HMAA4GR7AJR4N-WM (T8, T4, TG)': 'SK Hynix', 'HMAA8GR7AJR4N-WM': 'SK Hynix', 'M391A2K43BB1-CTD': 'Samsung', 'M393A2K40CB2-CVF (2015/2016)': 'Samsung', 'M393A2K40CB2-CTD': 'Samsung', 'M393A4K40CB2-CVF (2015/2016)': 'Samsung', 'M393A4K40CB2-CTD (2006)': 'Samsung', 'M393A4K40CB2-CTD (1915)': 'Samsung', 'M393A4K40CB2-CTD (1902/1903))': 'Samsung', 'HMA82GR7CJR4N-VK': 'SK Hynix', 'HMA84GR7CJR4N-VK': 'SK Hynix', 'HMAA8GL7CPR4N-WM': 'SK Hynix', 'M393A2K40DB2-CVF (2035)': 'Samsung', 'M393A4K40DB2-CVF (2001)': 'Samsung', 'M393A4K40DB2-CVF (2030)/(2035)': 'Samsung', 'M393A2K40DB3-CWE': 'Samsung', 'M393A4K40DB3-CWE': 'Samsung', 'HMA82GR7DJR4N-WM': 'SK Hynix', 'HMA84GR7DJR4N-WM': 'SK Hynix', 'HMA82GR7JJR4N-VK': 'SK Hynix', 'HMA84GR7JJR4N-VK': 'SK Hynix', 'HMA84GR7MFR4N-UH': 'SK Hynix', 'NT16GA72D4PBX3P-HR (old)': 'Nanya', 'NT16GA72D4PBX3P-IX': 'Nanya', 'NT32GA72D4NBX3P-HR': 'Nanya', 'NT32GA72D4NBX3P-IX': 'Nanya', 'NT32GA72D4NXA3P-HR(old)': 'Nanya', 'MTA36ASF2G72PZ-2G6E1': 'Micron', 'MTA18ASF2G72PDZ-2G6D1': 'Micron', 'MTA36ASF2G72PZ-2G6F1': 'Micron', 'MTA18ASF2G72PDZ-2G6E1': 'Micron', 'MTA18ASF2G72PZ-2G9E1': 'Micron', 'MTA36ASF4G72PZ-2G6E1': 'Micron', 'MTA36ASF4G72PZ-2G9E2': 'Micron', 'MTA18ASF2G72PZ-2G6J1': 'Micron', 'MTA18ASF2G72PZ-2G9J3': 'Micron', 'MTA36ASF4G72PZ-2G6J1': 'Micron', 'MTA18ASF4G72PZ-2G6B1': 'Micron', 'MTA18ASF4G72PZ-2G9B1': 'Micron', 'MTA36ASF8G72PZ-2G9B2': 'Micron', 'MTA36ASF8G72PZ-3G2B2': 'Micron', 'MTA18ASF4G72PZ-2G9E1': 'Micron', 'MTA18ASF4G72PZ-3G2E1': 'Micron', 'MTA36ASF8G72PZ-2G9E1': 'Micron', 'MTA36ASF8G72PZ-3G2E1': 'Micron', 'MTA36ASF4G72PZ-3G2J3': 'Micron', 'MTA18ASF2G72PZ-3G2J3': 'Micron', 'MS16D432R22S8MEX': 'Kingston', 'MS32D432R22S4MEX': 'Kingston', '0F27479': 'HGST/Western Digital', '0F29630': 'HGST/Western Digital', '0F29866': 'HGST/Western Digital', '0F34623': 'HGST/Western Digital', '0F38314': 'HGST/Western Digital', '0F14688': 'HGST/Western Digital', '1TT101-002': 'Seagate', '1TT101-401': 'Seagate', '2MU103-402': 'Seagate', '2K2101-401': 'Seagate', '2K2101-402': 'Seagate', '2HZ100-401': 'Seagate', '1V4107-002': 'Seagate', '2RM102-402': 'Seagate', '2KG103-401': 'Seagate', '2K8122-402': 'Seagate', '2LQ202-403': 'Seagate', '2KH103-402': 'Seagate', '2RK202-401': 'Seagate', '2MQ101-402': 'Seagate', '3AY212-401': 'Seagate', 'HDEPV10SMA51': 'Toshiba', 'HDEPR03GEA51': 'Toshiba', 'HDEPR01SMA51': 'Toshiba', 'HDEPW21SMA51': 'Toshiba', '0F31114': 'HGST/Western Digital', '0F38313': 'HGST/Western Digital', '0B35950': 'HGST/Western Digital', 'AD2-KW960 ': 'Lite-On', 'AD2-KW960': 'Lite-On', 'EP2-KB960': 'Lite-On', 'EP3-KW960': 'Lite-On', 'EP4-KW960': 'Lite-On', 'EPX-KW960': 'Lite-On', 'SSDPELKX019T8D': 'Intel', 'SSDPELKX960G8D-201': 'Intel', 'SSDPELKX960G8D-203': 'Intel', 'SSDPELKX038T8D': 'Intel', 'SSDPELKX020T8D-201': 'Intel', 'HFS960GD0TEG-6410A': 'SK Hynix', 'HFS1T9GD0FEH-6410A BA': 'SK Hynix', 'HFS960GD0FEG-A430A (Purple dot; common PSID)': 'SK Hynix', 'HFS1T9GD0FEH-A430A': 'SK Hynix', 'HFS3T8GD0FEH-A430A': 'SK Hynix', 'HFS1T9GD0FEI-A430A (new)': 'SK Hynix', 'HFS3T8GD0FEI-A430A': 'SK Hynix', 'HFS960GD0FEI-A430A (Purple dot: Common PSID)': 'SK Hynix', 'MZ1LV960HCJH-000MU': 'Samsung', 'MZ1LW1T9HMLS-000MV': 'Samsung', 'MZ1LW960HMJP-000MV': 'Samsung', 'MZ1LB1T9HALS-00AMV': 'Samsung', 'MZ1LB1T9HALS-000MV': 'Samsung', 'MZ1LB3T8HMLA-000MV': 'Samsung', 'MZ1LB960HAJQ-000MV': 'Samsung', 'MZ1LB960HAJQ-00AMV': 'Samsung', 'MZ1LB960HBJR-00AMV': 'Samsung', 'MZ1LB1T9HBLS-00AMV': 'Samsung', 'MZ1LB3T8HALA-00AMV': 'Samsung', 'MZ1L2960HCJR-00AMV': 'Samsung', 'MZ1L21T9HCLS-00AMV': 'Samsung', 'KXD5DLN13T84': 'Toshiba/Kioxia', 'M1113049-001': 'Microsoft', 'MZELB15THMLA-000MV': 'Samsung', 'SSDPEXNV153T8M2': 'Intel', '0TS2003': 'HGST/Western Digital', 'HFS15T3DGLX070N': 'SK Hynix', 'MTFDDAK960TCB': 'Micron', 'MTFDDAK960TCC': 'Micron', 'MTFDDAK960TDD': 'Micron', 'MZ7LM960HCHP-000MV': 'Samsung', 'MZ7LM960HMJP-000MV': 'Samsung', 'MZ7LH960HAJR-00AMV': 'Samsung', 'MZ7LH960HAJR-000MV': 'Samsung', 'HFS960G32MED-3410A': 'SK Hynix', 'HFS960G32MFD-3410A': 'SK Hynix', 'HFS960G32FEH-7A10A': 'SK Hynix', 'MZ7WD960HMHP-00003': 'Samsung', 'MZPLJ6T4HALA-00AMV': 'Samsung'}, 'description': {'M393A8G40AB2-CWE': 'Samsung 64GB', 'M393A4G40AB3-CVF': 'Samsung 32GB', 'M393A4G40AB3-CWE': 'Samsung 32GB', 'HMA42GR7AFR4N-TF': 'SK Hynix 16GB', 'HMN82GR7AFR4N-UH': 'SK Hynix 16GB  NN4', 'HMA84GR7AFR4N-UH': 'SK Hynix 32GB ', 'HMA82GR7AFR8N-VK': 'SK Hynix 16GB', 'HMA84GR7AFR4N-VK': 'SK Hynix 32GB ', 'HMAA8GL7AMR4N-VK': 'SK Hynix 64GB (Lenovo OEM)', 'HMAA4GR7AJR4N-WM (T8, T4, TG)': 'SK Hynix 32GB', 'HMAA8GR7AJR4N-WM': 'Sk Hynix 64GB', 'M391A2K43BB1-CTD': 'Samsung 16GB Lenovo Beast Utility DIMM', 'M393A2K40CB2-CVF (2015/2016)': 'Samsung 16GB CB2  ', 'M393A2K40CB2-CTD': 'Samsung 16GB', 'M393A4K40CB2-CVF (2015/2016)': 'Samsung 32G', 'M393A4K40CB2-CTD (2006)': 'Samsung 32GB ', 'M393A4K40CB2-CTD (1915)': 'Samsung 32GB ', 'M393A4K40CB2-CTD (1902/1903))': 'Samsung 32GB (QCL use only)', 'HMA82GR7CJR4N-VK': 'SK Hynix 16GB ', 'HMA84GR7CJR4N-VK': 'SK Hynix 32GB ', 'HMAA8GL7CPR4N-WM': 'SK Hynix 64GB ', 'M393A2K40DB2-CVF (2035)': 'Samsung 16GB', 'M393A4K40DB2-CVF (2001)': 'Samsung 32GB', 'M393A4K40DB2-CVF (2030)/(2035)': 'Samsung 32GB', 'M393A2K40DB3-CWE': 'Samsung 16GB DB3', 'M393A4K40DB3-CWE': 'Samsung 32GB ', 'HMA82GR7DJR4N-WM': 'SK Hynix 16GB 2933 (Davinci)', 'HMA84GR7DJR4N-WM': 'SK Hynix 32GB 2933', 'HMA82GR7JJR4N-VK': 'SK Hynix 16GB ', 'HMA84GR7JJR4N-VK': 'SK Hynix 32GB ', 'HMA84GR7MFR4N-UH': 'SK Hynix 32GB ', 'NT16GA72D4PBX3P-HR (old)': 'Nanya 16GB', 'NT16GA72D4PBX3P-IX': 'Nanya 16GB', 'NT32GA72D4NBX3P-HR': 'Nanya 32GB ', 'NT32GA72D4NBX3P-IX': 'Nanya 32GB ', 'NT32GA72D4NXA3P-HR(old)': 'Nanya 32GB ', 'MTA36ASF2G72PZ-2G6E1': 'Micron 16GB ', 'MTA18ASF2G72PDZ-2G6D1': 'Micron 16GB', 'MTA36ASF2G72PZ-2G6F1': 'Micron 16GB ', 'MTA18ASF2G72PDZ-2G6E1': 'Micron 16GB ', 'MTA18ASF2G72PZ-2G9E1': 'Micron 16GB', 'MTA36ASF4G72PZ-2G6E1': 'Micron 32GB ', 'MTA36ASF4G72PZ-2G9E2': 'Micron 32GB ', 'MTA18ASF2G72PZ-2G6J1': 'Micron 16GB', 'MTA18ASF2G72PZ-2G9J3': 'Micron 16GB', 'MTA36ASF4G72PZ-2G6J1': 'Micron 32GB ', 'MTA18ASF4G72PZ-2G6B1': 'Micron 32GB', 'MTA18ASF4G72PZ-2G9B1': 'Micron 32GB', 'MTA36ASF8G72PZ-2G9B2': 'Micron 64GB', 'MTA36ASF8G72PZ-3G2B2': 'Micron 64GB', 'MTA18ASF4G72PZ-2G9E1': 'Micron 32GB ', 'MTA18ASF4G72PZ-3G2E1': 'Micron 32GB', 'MTA36ASF8G72PZ-2G9E1': 'Micron 64GB', 'MTA36ASF8G72PZ-3G2E1': 'Micron 64GB', 'MTA36ASF4G72PZ-3G2J3': 'Micron 32GB', 'MTA18ASF2G72PZ-3G2J3': 'Micron 16GB', 'MS16D432R22S8MEX': 'Kingston 16GB', 'MS32D432R22S4MEX': 'Kingston 32GB', '0F27479': 'HGST10TB HDD ', '0F29630': 'HGST12TB HDD ', '0F29866': 'HGST 14TB HDD ', '0F34623': 'HGST 15TB ', '0F38314': 'HGST 18TB', '0F14688': 'HGST 4TB ', '1TT101-002': 'Seagate  10TB HDD ', '1TT101-401': 'Seagate  10TB HDD', '2MU103-402': 'Seagate  10TB HDD ', '2K2101-401': 'Seagate 12TB HDD', '2K2101-402': 'Seagate 12TB HDD ', '2HZ100-401': 'Seagate  4TB HDD ', '1V4107-002': 'Seagate  4TB HDD ', '2RM102-402': 'Seagate 12TB HDD', '2KG103-401': 'Seagate 14TB HDD ', '2K8122-402': 'Seagate 14TB HDD ', '2LQ202-403': 'Seagate 14TB HDD ', '2KH103-402': 'Seagate 16TB HDD ', '2RK202-401': 'Seagate 16TB HDD ', '2MQ101-402': 'Seagate  6TB HDD ', '3AY212-401': 'Seagate 18TB HDD', 'HDEPV10SMA51': 'Toshiba  10TB HDD ', 'HDEPR03GEA51': 'Toshiba 2TB', 'HDEPR01SMA51': 'Toshiba  4TB HDD', 'HDEPW21SMA51': 'Toshiba 12TB ', '0F31114': 'WD 14TB HDD', '0F38313': 'WD 16TB HDD', '0B35950': 'WD 4TB HDD ', 'AD2-KW960 ': 'Lite-On CNEX Labs NVME M.2(old)', 'AD2-KW960': 'Lite-On CNEX Labs NVME M.2(New)', 'EP2-KB960': 'Lite-On EP2-960', 'EP3-KW960': 'Lite-On EP3', 'EP4-KW960': 'Lite-On EP4 960GB', 'EPX-KW960': 'Lite-On EPX KW960-PLP', 'SSDPELKX019T8D': 'Intel P4511 1.9TB Gen 6', 'SSDPELKX960G8D-201': 'Intel P4511 960GB Gen5/Gen 6', 'SSDPELKX960G8D-203': 'Intel P4511 960GB Gen5/Gen 6', 'SSDPELKX038T8D': 'Intel P4511 3.8TB ', 'SSDPELKX020T8D-201': 'Intel P4511 1.9TB Gen 5', 'HFS960GD0TEG-6410A': 'SK Hynix PE4010 1TB', 'HFS1T9GD0FEH-6410A BA': 'SK Hynix PE4011 2TB', 'HFS960GD0FEG-A430A (Purple dot; common PSID)': 'SK Hynix PE6010 1TB', 'HFS1T9GD0FEH-A430A': 'SK Hynix PE6011 2TB', 'HFS3T8GD0FEH-A430A': 'SK Hynix PE6011 4TB', 'HFS1T9GD0FEI-A430A (new)': 'SK Hynix PE6110 2TB', 'HFS3T8GD0FEI-A430A': 'SK Hynix PE6110 4TB', 'HFS960GD0FEI-A430A (Purple dot: Common PSID)': 'SK Hynix PE6110 1TB', 'MZ1LV960HCJH-000MU': 'Samsung PM953 960GB', 'MZ1LW1T9HMLS-000MV': 'Samsung PM963 2TB', 'MZ1LW960HMJP-000MV': 'Samsung PM963 960GB', 'MZ1LB1T9HALS-00AMV': 'Samsung PM983 2TB', 'MZ1LB1T9HALS-000MV': 'Samsung PM983 2TB', 'MZ1LB3T8HMLA-000MV': 'Samsung PM983 4TB', 'MZ1LB960HAJQ-000MV': 'Samsung PM983 960GB', 'MZ1LB960HAJQ-00AMV': 'Samsung PM983 960GB', 'MZ1LB960HBJR-00AMV': 'Samsung PM983a 1TB', 'MZ1LB1T9HBLS-00AMV': 'Samsung PM983a 2TB', 'MZ1LB3T8HALA-00AMV': 'Samsung PM983a 4TB', 'MZ1L2960HCJR-00AMV': 'Samsung PM9A3 960GB', 'MZ1L21T9HCLS-00AMV': 'Samsung PM9A3 1.9TB', 'KXD5DLN13T84': 'Toshiba 3840GB SSD', 'M1113049-001': 'Twin Peak 960GB', 'MZELB15THMLA-000MV': 'Samsung E1.L PM983', 'SSDPEXNV153T8M2': 'Intel SSDPEXNV153T8D', '0TS2003': 'WD 1600GB', 'HFS15T3DGLX070N': 'PE8111 E1.L 18mm 16TB', 'MTFDDAK960TCB': 'Micron MTFDDAK960TCB', 'MTFDDAK960TCC': 'Micron MTFDDAK960TCC', 'MTFDDAK960TDD': 'Micron MTFDDAK960TDD 960GB SSD', 'MZ7LM960HCHP-000MV': 'Samsung PM863 SSD', 'MZ7LM960HMJP-000MV': 'Samsung PM863a 960GB SSD', 'MZ7LH960HAJR-00AMV': 'Samsung PM883 960GB SSD', 'MZ7LH960HAJR-000MV': 'Samsung PM883 960GB SSD', 'HFS960G32MED-3410A': 'SK Hynix 960GB SSD', 'HFS960G32MFD-3410A': 'SK Hynix 960GB SSD', 'HFS960G32FEH-7A10A': 'SK Hynix 960GB SSD', 'MZ7WD960HMHP-00003': 'Samsung SSD', 'MZPLJ6T4HALA-00AMV': 'Samsung SSD'}, 'quantity': {'M393A8G40AB2-CWE': 400, 'M393A4G40AB3-CVF': 138, 'M393A4G40AB3-CWE': 400, 'HMA42GR7AFR4N-TF': 28, 'HMN82GR7AFR4N-UH': 99, 'HMA84GR7AFR4N-UH': 2, 'HMA82GR7AFR8N-VK': 205, 'HMA84GR7AFR4N-VK': 25, 'HMAA8GL7AMR4N-VK': 241, 'HMAA4GR7AJR4N-WM (T8, T4, TG)': 406, 'HMAA8GR7AJR4N-WM': 400, 'M391A2K43BB1-CTD': 100, 'M393A2K40CB2-CVF (2015/2016)': 400, 'M393A2K40CB2-CTD': 76, 'M393A4K40CB2-CVF (2015/2016)': 21, 'M393A4K40CB2-CTD (2006)': 10, 'M393A4K40CB2-CTD (1915)': 128, 'M393A4K40CB2-CTD (1902/1903))': 439, 'HMA82GR7CJR4N-VK': 284, 'HMA84GR7CJR4N-VK': 962, 'HMAA8GL7CPR4N-WM': 313, 'M393A2K40DB2-CVF (2035)': 300, 'M393A4K40DB2-CVF (2001)': 9, 'M393A4K40DB2-CVF (2030)/(2035)': 31, 'M393A2K40DB3-CWE': 250, 'M393A4K40DB3-CWE': 742, 'HMA82GR7DJR4N-WM': 658, 'HMA84GR7DJR4N-WM': 565, 'HMA82GR7JJR4N-VK': 310, 'HMA84GR7JJR4N-VK': 753, 'HMA84GR7MFR4N-UH': 21, 'NT16GA72D4PBX3P-HR (old)': 250, 'NT16GA72D4PBX3P-IX': 176, 'NT32GA72D4NBX3P-HR': 32, 'NT32GA72D4NBX3P-IX': 248, 'NT32GA72D4NXA3P-HR(old)': 0, 'MTA36ASF2G72PZ-2G6E1': 316, 'MTA18ASF2G72PDZ-2G6D1': 177, 'MTA36ASF2G72PZ-2G6F1': 43, 'MTA18ASF2G72PDZ-2G6E1': 136, 'MTA18ASF2G72PZ-2G9E1': 397, 'MTA36ASF4G72PZ-2G6E1': 272, 'MTA36ASF4G72PZ-2G9E2': 740, 'MTA18ASF2G72PZ-2G6J1': 312, 'MTA18ASF2G72PZ-2G9J3': 300, 'MTA36ASF4G72PZ-2G6J1': 180, 'MTA18ASF4G72PZ-2G6B1': 400, 'MTA18ASF4G72PZ-2G9B1': 400, 'MTA36ASF8G72PZ-2G9B2': 400, 'MTA36ASF8G72PZ-3G2B2': 336, 'MTA18ASF4G72PZ-2G9E1': 115, 'MTA18ASF4G72PZ-3G2E1': 0, 'MTA36ASF8G72PZ-2G9E1': 400, 'MTA36ASF8G72PZ-3G2E1': 0, 'MTA36ASF4G72PZ-3G2J3': 365, 'MTA18ASF2G72PZ-3G2J3': 460, 'MS16D432R22S8MEX': 128, 'MS32D432R22S4MEX': 56, '0F27479': 500, '0F29630': 197, '0F29866': 310, '0F34623': 60, '0F38314': 98, '0F14688': 20, '1TT101-002': 373, '1TT101-401': 35, '2MU103-402': 0, '2K2101-401': 62, '2K2101-402': 9, '2HZ100-401': 146, '1V4107-002': 34, '2RM102-402': 94, '2KG103-401': 19, '2K8122-402': 60, '2LQ202-403': 250, '2KH103-402': 84, '2RK202-401': 22, '2MQ101-402': 40, '3AY212-401': 64, 'HDEPV10SMA51': 274, 'HDEPR03GEA51': 60, 'HDEPR01SMA51': 42, 'HDEPW21SMA51': 44, '0F31114': 551, '0F38313': 412, '0B35950': 10, 'AD2-KW960 ': 70, 'AD2-KW960': 89, 'EP2-KB960': 6, 'EP3-KW960': 560, 'EP4-KW960': 0, 'EPX-KW960': 67, 'SSDPELKX019T8D': 40, 'SSDPELKX960G8D-201': 123, 'SSDPELKX960G8D-203': 34, 'SSDPELKX038T8D': 119, 'SSDPELKX020T8D-201': 10, 'HFS960GD0TEG-6410A': 66, 'HFS1T9GD0FEH-6410A BA': 90, 'HFS960GD0FEG-A430A (Purple dot; common PSID)': 16, 'HFS1T9GD0FEH-A430A': 2, 'HFS3T8GD0FEH-A430A': 118, 'HFS1T9GD0FEI-A430A (new)': 61, 'HFS3T8GD0FEI-A430A': 179, 'HFS960GD0FEI-A430A (Purple dot: Common PSID)': 20, 'MZ1LV960HCJH-000MU': 88, 'MZ1LW1T9HMLS-000MV': 49, 'MZ1LW960HMJP-000MV': 202, 'MZ1LB1T9HALS-00AMV': 42, 'MZ1LB1T9HALS-000MV': 70, 'MZ1LB3T8HMLA-000MV': 171, 'MZ1LB960HAJQ-000MV': 975, 'MZ1LB960HAJQ-00AMV': 43, 'MZ1LB960HBJR-00AMV': 590, 'MZ1LB1T9HBLS-00AMV': 190, 'MZ1LB3T8HALA-00AMV': 190, 'MZ1L2960HCJR-00AMV': 0, 'MZ1L21T9HCLS-00AMV': 200, 'KXD5DLN13T84': 156, 'M1113049-001': 0, 'MZELB15THMLA-000MV': 76, 'SSDPEXNV153T8M2': 35, '0TS2003': 21, 'HFS15T3DGLX070N': 180, 'MTFDDAK960TCB': 25, 'MTFDDAK960TCC': 2, 'MTFDDAK960TDD': 79, 'MZ7LM960HCHP-000MV': 1, 'MZ7LM960HMJP-000MV': 18, 'MZ7LH960HAJR-00AMV': 1, 'MZ7LH960HAJR-000MV': 55, 'HFS960G32MED-3410A': 40, 'HFS960G32MFD-3410A': 1, 'HFS960G32FEH-7A10A': 163, 'MZ7WD960HMHP-00003': 2, 'MZPLJ6T4HALA-00AMV': 0}}

    crd_path = r'\\172.30.1.100\pxe\Kirkland_Lab\Microsoft_CSI\Documentation\PipeCleaner_Attachments\CRD\Gen_7.x'
    crd_file = r'\M1117080 CRD,FY19Q4,AZURE,GEN7.0,COMPUTE GP-MED,INTEL,WIWYNN,104_REVK.XLSX'
    crd_complete = f'{crd_path}{crd_file}'

    csv_file_name = access_crd(ticket_number, crd_complete, sheet_name)

    with open(f'{csv_file_name}', 'r') as f:
        for line in f:
            for part_number in inventory['item_type']:
                if part_number in line:
                    csv_match.append(part_number)

    # print(csv_match)


    # for item in data_frame:
    #     print(item)

    # for line in content:
    #     print(line)

    # for item in ticket_to_crd:
    #     print(item)
    # for item in inventory:
    #     print(item)


main()