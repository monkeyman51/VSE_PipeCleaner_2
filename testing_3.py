from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext, ClientCredential
from office365.sharepoint.files.file import File

username: str = 'joe.ton@VeritasDCservices.com'
password: str = 'FordFocus24'

sharepoint_file: str = 'https://veritasdcservices.sharepoint.com/:x:/r/sites/Veritas-KirklandWA/_layouts/15/' \
                       'Doc.aspx?sourcedoc=%7BA91E72E6-2453-4873-9515-FD1BE40866C1%7D&file=' \
                       'Veritas%20Kirkland%20Commodity%20Inventory.xlsx'

relative_url: str = '/sites/Veritas-KirklandWA/_layouts/15/' \
                    'Doc.aspx?sourcedoc=%7BA91E72E6-2453-4873-9515-FD1BE40866C1%7D&file=Veritas%20Kirkland%20Commodity%' \
                    '20Inventory.xlsx'

main: str = 'https://veritasdcservices.sharepoint.com/sites/Veritas-KirklandWA'

new_client_id: str = '321f5249-3384-4edf-b0dc-c2d47d69be38'
# new_client_id: str = '751040bd-e908-46a8-827c-ec32248b3e4d'
new_client_secret: str = 'Tps-UM8j1vQy5R_PV90ENwk9~n6l~J70Kk'

credentials = ClientCredential(new_client_id, new_client_secret)

ctx = ClientContext(sharepoint_file).with_credentials(credentials)

file_path_2: str = r"C:\Users\joe.ton\Documents\test_22.xlsx"
with open(file_path_2, "wb") as local_file:
    file = ctx.web.get_file_by_server_relative_url(sharepoint_file).download(local_file).execute_query()
