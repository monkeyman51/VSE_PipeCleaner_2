from pipe_cleaner.src.credentials import Path
import requests


def request_console_server(host_id):
    data = {
        'action': 'get_json_data',
        'host_id': f'{host_id}'
    }
    response = requests.post(url=f'http://172.30.1.100/results/{host_id}.json', data=data)

    with open(f'{Path.info}{host_id}.json', 'w') as f:
        f.write(response.text)


request_console_server('C2012-T0100989290003AK')
