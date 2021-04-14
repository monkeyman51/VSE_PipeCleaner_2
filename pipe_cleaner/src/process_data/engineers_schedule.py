"""
2/12/2021

Takes all unique TRRs gathered from the all the Host Groups (Pipes) within Console Server. Structures the information
into a data structure to create an Excel output for engineer schedule particularly for quals.  This is also in hopes to
provide a small data base within the shared VSE Z: Drive based on the engineer schedule.
"""
import base64
from json import loads

from bs4 import BeautifulSoup
from requests_html import HTMLSession

from execute_credentials import get_personal_access_token

hyperlink_example_1: str = 'https://azurecsi.visualstudio.com/CSI%20Commodity%20Qualification/_testPlans/' \
                           'execute?planId=368793&suiteId=368794'
hyperlink_example_2: str = 'https://azurecsi.visualstudio.com/CSI%20Commodity%20Qualification/_testPlans/' \
                           'execute?planId=368793&suiteId=368794TestCase?witFields=' \
                           'System.Id&expand=true&returnIdentityRef=true&excludeFlags=0&isRecursive=false'


class CleanAssignedToName:
    """
    Responsible for cleaning Assigned To names found in TRRs within Microsoft's Azure DevOps for Commodity Testing.
    Cleaning these names are for output within the VSE Pipe Cleaner
    """

    def __init__(self, assigned_to_name: str):
        self.assigned_to_name: str = assigned_to_name

    def clean_assigned_to_field(self) -> str:
        """
        Cleans name of the engineer getting rid of unnecessary strings such as periods or email information.
        """
        if self.check_empty_names_as_none() == 'None':
            return 'None'

        else:
            self.check_email_in_name()
            self.gets_first_and_last_from_full_name()
            self.replace_unnecessary_characters_in_name()

            return self.assigned_to_name

    def replace_email_from_name(self) -> str:
        """
        If engineer has email in it, gets rid of string characters after @ ie. @ character, mail server, domain
        """
        irrelevant_part: str = '@' + self.assigned_to_name.split('@')[-1]
        self.assigned_to_name: str = self.assigned_to_name.replace(irrelevant_part, '')

        return self.assigned_to_name

    def check_email_in_name(self) -> str:
        """
        Checks @ in engineer name.
        """
        if '@' in self.assigned_to_name:
            self.assigned_to_name: str = self.replace_email_from_name()

            return self.assigned_to_name

        else:
            return self.assigned_to_name

    def replace_periods_in_name_to_spaces(self) -> str:
        """
        After cleaning engineer name, convert name to title.
        """
        self.assigned_to_name: str = self.assigned_to_name.replace('.', ' ')

        return self.assigned_to_name

    def convert_name_to_title(self) -> str:
        """
        Converts the engineer name from ADO to title.
        """
        self.assigned_to_name = self.assigned_to_name.title()

        return self.assigned_to_name

    def gets_first_and_last_from_full_name(self) -> str:
        """
        Assumes that full name has periods that separate the different names as per standard.
        Only accounts for first and last name of full name. Excluding middle names.
        """
        first_name: str = self.assigned_to_name.split('.')[0]
        last_name: str = self.assigned_to_name.split('.')[-1]
        self.assigned_to_name: str = f'{first_name}.{last_name}'

        return self.assigned_to_name

    def remove_ext_in_name(self) -> str:
        """
        Remove EXT in name which denotes as a current contractor or previous contractor.
        """
        assigned_to_name: str = self.assigned_to_name.upper()

        self.assigned_to_name: str = assigned_to_name.replace('-EXT', '')

        return self.assigned_to_name

    def replace_space_to_period_in_name(self) -> str:
        """
        Converts the spaces in the name to period. Assures that the name is standard before processing.
        """
        return self.assigned_to_name.replace(' ', '.')

    def replace_unnecessary_characters_in_name(self) -> str:
        """
        Replace unnecessary information in full name
        """
        self.replace_periods_in_name_to_spaces()
        self.remove_unnecessary_spaces_in_name()
        self.remove_ext_in_name()
        self.convert_name_to_title()

        return self.assigned_to_name

    def replace_first_character_with_empty_space(self) -> str:
        """
        Checks for empty characters and replaces them with empty space within the assigned to name.
        """
        while self.assigned_to_name[0].isspace():
            self.assigned_to_name: str = self.assigned_to_name[1:]

        return self.assigned_to_name

    def replace_last_character_with_empty_space(self) -> str:
        """
        Checks for empty characters and replaces them with empty space within the assigned to name.
        """
        while self.assigned_to_name[-1].isspace():
            self.assigned_to_name: str = self.assigned_to_name[0:-1]

        return self.assigned_to_name

    def remove_unnecessary_spaces_in_name(self) -> str:
        """
        Account for names that have empty spaces before or after the full name.
        """
        self.assigned_to_name: str = self.replace_first_character_with_empty_space()
        self.assigned_to_name: str = self.replace_last_character_with_empty_space()

        return self.assigned_to_name

    def check_empty_names_as_none(self) -> str:
        """
        Account for TRRs that have empty names.
        """
        if not self.assigned_to_name:
            return 'None'


class TestCaseResponse:
    """
    Responsible for getting a specific test case as a HTML response.
    Not responsible for parsing data from response yet.
    """

    def __init__(self, test_plan_hyperlink: str):
        self.test_case_api: dict = {}
        self.test_plan_hyperlink: str = test_plan_hyperlink

    def main_method(self) -> str:
        """
        Setup proper data to access data with correct API credentials to retrieve html response.
        :return: Test Case API data
        """
        self.setup_api_credentials()
        self.request_test_case_text()

        return self.convert_test_case_to_json()

    def setup_api_credentials(self) -> dict:
        """
        Get basic API information to later request test case data. 
        """
        self.store_test_plan_hyperlink()
        self.store_personal_access_token()
        self.store_user_password()
        self.replace_define_with_execute_hyperlink()

        self.encode_base64_user_password()
        self.decode_base64_user_password()

        return self.test_case_api

    def store_test_plan_hyperlink(self) -> dict:
        """
        Store the actual URL that will be called
        """
        self.test_case_api['test_plan_hyperlink']: str = self.test_plan_hyperlink

        return self.test_case_api

    def store_personal_access_token(self) -> dict:
        """
        Get Access Token which has a expiration date
        """
        self.test_case_api['personal_access_token']: str = get_personal_access_token()

        return self.test_case_api

    def store_user_password(self) -> dict:
        """
        User is none. So it would be an empty string.
        """
        personal_access_token: str = self.test_case_api.get('personal_access_token')
        self.test_case_api['user_password'] = f':{personal_access_token}'

        return self.test_case_api

    def replace_define_with_execute_hyperlink(self) -> dict:
        """
        Default test plan hyperlink found in the TRR is for define URL.
        """
        test_plan_hyperlink: str = self.test_case_api.get('test_plan_hyperlink')
        self.test_case_api['test_plan_hyperlink']: str = test_plan_hyperlink.replace('define?planId=',
                                                                                     'execute?planId=')

        return self.test_case_api

    def encode_base64_user_password(self) -> dict:
        """
        Decode Base64 User Password
        """
        self.test_case_api['encode_user_password']: bytes = self.test_case_api.get('user_password').encode()

        return self.test_case_api

    def decode_base64_user_password(self) -> dict:
        """
        Decode to string for correct credentials for test plans access.
        """
        base64_user_password: bytes = self.test_case_api.get('encode_user_password')
        self.test_case_api['decode_user_password']: bytes = base64.b64encode(base64_user_password).decode()

        return self.test_case_api

    def request_test_case(self):
        """

        """
        decode_user_password: str = self.test_case_api.get('decode_user_password')
        test_plan_hyperlink: str = self.test_case_api.get('test_plan_hyperlink')

        headers: dict = {'Authorization': f'Basic {decode_user_password}'}
        session = HTMLSession().get(test_plan_hyperlink, headers=headers)
        # print(f'TESTING... {session.html.render()}')

        return session

    def request_test_case_text(self) -> str:
        """
        Request HTML response, not yet parsed.
        """
        self.test_case_api['response_text']: str = self.request_test_case().text
        return self.request_test_case().text

    def convert_test_case_to_json(self):
        """
        Convert the HTML text to JSON
        """
        response_text: str = self.test_case_api.get('response_text')
        print(response_text)
        beautiful_soup = BeautifulSoup(response_text, "html.parser")

        return loads(str(beautiful_soup.find('script', type='application/json').contents[0]))


class ParseTestCaseJSON:
    """
    After retrieving test case HTML response as JSON, grabs essential information dealing with test case.
    """

    def __init__(self, test_case_json: str):
        self.test_case_progress: dict = {}
        self.test_case_json: str = test_case_json

    def main_method(self):
        """

        """
        self.get_test_points()
        self.collect_data()

        return self.test_case_progress

    def get_test_points(self):
        """
        Get
        """
        self.test_case_progress['test_points']: list = self.test_case_json.get('data', {}). \
            get('ms.vss-test-web.test-plans-hub-refresh-data-provider', {}).get('testPoints')

        return self.test_case_progress

    def collect_data(self):
        """
        Store data into data structure
        """
        self.get_name_of_test_point()

    def get_name_of_test_point(self):
        """
        5 Character code id for a test point
        """
        test_points: list = self.test_case_progress.get('test_points')
        self.test_case_progress['progress']: dict = {}
        for test_case in test_points:
            test_case_name: str = test_case.get('testCaseReference', {}).get('name')[1:6]
            self.test_case_progress['progress'][test_case_name]: dict = {}

            outcome: int = test_case.get('results', {}).get('outcome')
            state: int = test_case.get('results', {}).get('state')

            if outcome == 2 and state == 2:
                return 'Passed'
            elif outcome == 3 and state == 3:
                return 'Failed'
            elif outcome == 11 and state == 2:
                return 'Not Applicable'
            elif outcome == 7 and state == 3:
                return 'Blocked'
            # elif

        return self.test_case_progress

    # def analyze_test_points(self):
    #     """
    #
    #     """
    #     self.test_case_progress.get('test_points')


response = TestCaseResponse(test_plan_hyperlink=hyperlink_example_1).main_method()
# data = ParseTestCaseJSON(test_case_json=response).main_method()
# print(dumps(data, sort_keys=True, indent=4))
# print(dumps(response, sort_keys=True, indent=4))
