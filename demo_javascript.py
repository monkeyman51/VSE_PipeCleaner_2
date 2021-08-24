from selenium import webdriver

from pipe_cleaner.src.credentials import AccessADO as Ado
from base64 import b64encode


def get_response_text_from_ado():
    """
    Get content from all recent ADO work items for grabbing TRR IDs. This is an effort towards automating part number
    library.
    """
    browser = webdriver.Chrome()
    site_url: str = 'https://azurecsi.visualstudio.com/CSI%20Commodity%20Qualification/_queries/query/' \
                    '5233817c-b790-4482-8cb7-200aae92f508/'
    user_password: str = f'{Ado.token_name}:{Ado.personal_access_token}'

    headers = {'Authorization': f'Basic {b64encode(user_password.encode()).decode()}'}

    return browser.get(site_url)


response = get_response_text_from_ado()
print(response)
