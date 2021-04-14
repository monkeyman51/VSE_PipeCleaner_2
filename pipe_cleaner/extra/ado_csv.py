from bs4 import BeautifulSoup
import csv
import re


data = {}

def request_ado(trr_id):
    base_url = 'https://azurecsi.visualstudio.com/'
    path_url = 'CSI%20Commodity%20Qualification/_apis/wit/workitems?'
    query_parameter = f'id={str(trr_id)}&$expand=all&api-version=5.1'
    user_password = ADO.token_name + ':' + ADO.personal_access_token
    web_address = base_url + path_url + query_parameter
    base64_user_password = base64.b64encode(user_password.encode()).decode()
    headers = {'Authorization': 'Basic %s' % base64_user_password}

    try:
        ado_response = requests.get(
            web_address, headers=headers, timeout=1)
        ado_response.raise_for_status()
        get_csv(trr_id, ado_response)
        get_json(trr_id)
    except requests.exceptions.Timeout:
        print('ADO Response: Timeout Occurred')
        ado_response = requests.get(
            web_address, headers=headers, timeout=5)
        ado_response.raise_for_status()
        get_csv(trr_id, ado_response)
        get_json(trr_id)


def get_csv(trr_id, ado_response):
    with open(f'{Path.info}{trr_id}_description.csv', 'w', newline='', encoding='utf-8') as f:
        soup = BeautifulSoup(ado_response.text, 'html.parser')
        table = soup.findAll('table')[0]
        rows = table.findAll('tr')
        writer = csv.writer(f)
        for row in rows:
            csv_row = []
            for cell in row.findAll(['td', 'th']):
                csv_row.append(cell.get_text())
                cell = [re.sub(r'\\n  ', '', i) for i in csv_row]
                row = [re.sub(r'\xa0', '', i) for i in cell]
            writer.writerow(row)
    with open(f'{Path.info}{trr_id}_component.csv', 'w', newline='', encoding='utf-8') as f:
        soup = BeautifulSoup(ado_response.text, 'html.parser')
        table = soup.findAll('table')[1]
        rows = table.findAll('tr')
        writer = csv.writer(f)
        for row in rows:
            csv_row = []
            for cell in row.findAll(['td', 'th']):
                csv_row.append(cell.get_text())
                cell = [re.sub(r'\\n  ', '', i) for i in csv_row]
                row = [re.sub(r'\xa0', '', i) for i in cell]
            writer.writerow(row)