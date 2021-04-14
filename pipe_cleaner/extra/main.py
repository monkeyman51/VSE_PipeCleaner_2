import requests
from requests.auth import HTTPBasicAuth

username = 'root'
password = '$pl3nd1D'  # PAT expires 7/26/2020
r = requests.get('https://192.168.237.216:8080/redfish/v1', auth=HTTPBasicAuth(username, password))
print(r.status_code)

# def get_bmc_access(ids):
#     url = 'https://192.168.237.216:8080/redfish/v1'
#     user_pass = token_name + ':' + personal_access_token
#     web_address = base_url + path_url + query_parameter
#     base64_user_pass = base64.b64encode(user_pass.encode()).decode()
#     headers = {'Authorization': 'Basic %s' % base64_user_pass}
#
#     try:
#         ado_response = requests.get(
#             web_address, headers=headers, timeout=1)
#         ado_response.raise_for_status()
#
#     except (requests.exceptions.HTTPError, IndexError):
#         pass
#
# url = 'https://192.168.237.216:8080/redfish/v1'
# r = requests.get(url)
# print(r.status_code)




# get_bmc_access()