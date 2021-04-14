from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext, ClientCredential
from office365.sharepoint.files.file import File

import sys


sharepoint_file: str = 'https://veritasdcservices.sharepoint.com/:x:/r/sites/Veritas-KirklandWA/_layouts/15/' \
                       'Doc.aspx?sourcedoc=%7BA91E72E6-2453-4873-9515-FD1BE40866C1%7D&file=' \
                       'Veritas%20Kirkland%20Commodity%20Inventory.xlsx'

main_url: str = 'https://veritasdcservices.sharepoint.com/sites/Veritas-KirklandWA'
relative_url: str = '/sites/Veritas-KirklandWA/_layouts/15/' \
                    'Doc.aspx?sourcedoc=%7BA91E72E6-2453-4873-9515-FD1BE40866C1%7D&file=Veritas%20Kirkland%20Commodity%' \
                    '20Inventory.xlsx'

file_url: str = 'https://veritasdcservices.sharepoint.com/:x:/r/sites/Veritas-KirklandWA/_layouts/15/' \
                'Doc.aspx?sourcedoc=%7BA91E72E6-2453-4873-9515-FD1BE40866C1%7D&file=Veritas%20Kirkland%20Commodity%' \
                '20Inventory.xlsx'

download_url: str = 'https://veritasdcservices.sharepoint.com/sites/Veritas-KirklandWA/_layouts/15/download.aspx?UniqueId=a91e72e6-2453-4873-9515-fd1be40866c1'


username: str = 'joe.ton@VeritasDCservices.com'
password: str = 'FordFocus24'
# client_id: str = '208d1001-5758-4a4f-a92b-48e746da6ec8'
# client_secret: str = 'Os8viotO1h204K5DI1CMq8cOMBwv/J43Hifqm6tgkQ8='

new_client_id: str = '321f5249-3384-4edf-b0dc-c2d47d69be38'
new_client_secret: str = 'Tps-UM8j1vQy5R_PV90ENwk9~n6l~J70Kk'

file_path_2: str = r"C:\Users\joe.ton\Documents\test_22.xlsx"
file_path_3: str = r"C:\Users\joe.ton\Documents\test_3.xlsx"

# ctx_auth = AuthenticationContext(sharepoint_file)
# if ctx_auth.acquire_token_for_user(username, password):
#     ctx = ClientContext(sharepoint_file, ctx_auth)
#     web = ctx.web
#     ctx.load(web)
#     ctx.execute_query()
#
#     # response = File.open_binary(ctx, relative_url).content
#     response = File.
#     foo = sys.getsizeof(response)
#     print(foo)

credentials = ClientCredential(new_client_id, new_client_secret)
ctx = ClientContext(sharepoint_file).with_credentials(credentials)
target_web = ctx.web.get().execute_query()
ctx.load(target_web)
ctx.execute_query()
response = File.open_binary(ctx, relative_url).content
foo = sys.getsizeof(response)
print(foo)

    # with open(file_path_2, 'wb') as local_file:
    #     if not response.content:
    #         print(f'NULL')
    #     else:
    #         local_file.write(response.content)
    #
    # read_file = open(file_path_2, 'rb')
    # data = read_file.read()
    #
    # b64 = base64.b64encode(data)
    #
    # # Save file
    # decode_b64 = base64.b64decode(b64)
    # out_file = open(file_path_3, 'wb')
    # out_file.write(decode_b64)

# import json
#
# foo = json.loads(json_string)
# print(json.dumps(foo, sort_keys=True, indent=4))

    # with open(file_path_2, "wb") as local_file:
    #     local_file.write(response)

# response_2 = requests.get(download_url)
# print(response_2.status_code)
#
# with open(file_path_2, 'wb') as f:
#     f.write(response_2)
