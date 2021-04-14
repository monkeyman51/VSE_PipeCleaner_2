import requests


def get_console_server_access(host_id):
    data = {
        'action': 'get_json_data',
        'host_id': f'{host_id}'
    }
    response = requests.post(url=f'http://172.30.1.100/results/{host_id}.json', data=data)

    with open(f'{host_id}.json', 'w') as f:
        f.write(response.text)
        host = response.text

    return host

host = get_console_server_access('C104EH-11S01PG420Y150QT87T7B6')

print(host)
