import pytest

import engineers_schedule as schedule


@pytest.fixture()
def defined_hyperlink_example() -> str:
    """
    Example of test case hyperlink based on TRR
    """
    return 'https://azurecsi.visualstudio.com/CSI%20Commodity%20Qualification/_testPlans/' \
           'execute?planId=368793&suiteId=368794'


@pytest.fixture()
def test_case_data_example(defined_hyperlink_example: str) -> str:
    """
    Based on defined_hyperlink_example, extracts from TestCaseResponse class
    """
    return schedule.TestCaseResponse(test_plan_hyperlink=defined_hyperlink_example).main_method()


def test_capitalization_in_name():
    """
    The first character in the first name and last name of a name should be capitalized.
    """
    clean_assigned_to = schedule.CleanAssignedToName(assigned_to_name='  joe.ton-ext@VeritasDCservices.com  ')

    result: str = clean_assigned_to.clean_assigned_to_field()

    assert result.istitle()


def test_check_email_in_name():
    """
    Checks to make sure VSE email is not in name.
    """
    clean_assigned_to = schedule.CleanAssignedToName(assigned_to_name='joe.ton@VeritasDCservices.com')

    result: str = clean_assigned_to.check_email_in_name()

    assert '@' not in result


def test_checks_space_between_first_and_last_name():
    """
    Checks space in between.
    """
    clean_assigned_to = schedule.CleanAssignedToName(assigned_to_name='joe.ton')

    result: str = clean_assigned_to.replace_periods_in_name_to_spaces()

    assert ' ' in result


def test_assigned_to_name_is_title():
    """
    Checks assigned to name is properly title
    """
    clean_assigned_to = schedule.CleanAssignedToName(assigned_to_name='joe.ton')

    result: str = clean_assigned_to.convert_name_to_title()

    assert result.istitle()


def test_checks_for_two_names_in_full_name():
    """
    Possible first, middle, and last name scenario
    """
    clean_assigned_to = schedule.CleanAssignedToName(assigned_to_name='joe.that.ton')

    result: str = clean_assigned_to.gets_first_and_last_from_full_name()
    result: list = result.split('.')

    assert len(result) == 2


def test_check_for_ext_in_name():
    """
    Accounts for EXT in name for contractors
    """
    clean_assigned_to = schedule.CleanAssignedToName(assigned_to_name='joe.ton-EXT')

    result: str = clean_assigned_to.remove_ext_in_name()

    assert '-EXT' not in result


def test_first_character_in_assigned_to_name_for_empty_space():
    """
    Accounts for possible names with spaces as a default state for first character.
    """
    clean_assigned_to = schedule.CleanAssignedToName(assigned_to_name='   joe ton')

    result: str = clean_assigned_to.replace_first_character_with_empty_space()

    assert not result[0].isspace()


def test_last_character_in_assigned_to_name_for_empty_space():
    """
    Accounts for possible names with spaces as a default state for last character.
    """
    clean_assigned_to = schedule.CleanAssignedToName(assigned_to_name='joe ton  ')

    result: str = clean_assigned_to.replace_last_character_with_empty_space()

    assert not result[-1].isspace()


def test_none_in_assigned_to_field():
    """
    Account for TRRs that have empty assigned to names.
    """
    clean_assigned_to = schedule.CleanAssignedToName(assigned_to_name='')

    result: str = clean_assigned_to.clean_assigned_to_field()

    assert result == 'None'


def test_available_personal_access_token(defined_hyperlink_example: str):
    """
    Check to make API call to URL for test cases within ADO works
    """
    test_case_response = schedule.TestCaseResponse(test_plan_hyperlink=defined_hyperlink_example)

    test_case_api: dict = test_case_response.store_personal_access_token()
    result: str = test_case_api.get('personal_access_token')

    assert result.isascii()


def test_base_url_for_test_case_url(defined_hyperlink_example: str):
    """
    Assure that correct base url is given
    """
    test_case_response = schedule.TestCaseResponse(test_plan_hyperlink=defined_hyperlink_example)

    test_case_api: dict = test_case_response.store_test_plan_hyperlink()
    result: str = test_case_api.get('test_plan_hyperlink')

    assert '_testPlans' in result


# Automation tests for HTTP response takes way too long
# def test_200_response_from_test_case_hyperlink():
#     """
#     See if response gives JSON file back
#     """
#     test_case_url: str = 'https://azurecsi.visualstudio.com/CSI%20Commodity%20Qualification/_testPlans/' \
#                          'execute?planId=206199&suiteId=206200'
#     personal_access_token: str = execute_credentials.get_personal_access_token()
#     user_password = f':{personal_access_token}'
#     base64_user_password = base64.b64encode(user_password.encode()).decode()
#     headers = {'Authorization': 'Basic %s' % base64_user_password}
#
#     test_case_response = requests.get(test_case_url, headers=headers)
#     result = test_case_response.status_code
#
#     assert result == 200


def test_define_to_execute_hyperlink(defined_hyperlink_example):
    """
    Assure that define in the test case hyperlink to execute test case hyperlink
    """
    test_case_response = schedule.TestCaseResponse(test_plan_hyperlink=defined_hyperlink_example)

    test_case_response.store_test_plan_hyperlink()
    test_case_api: dict = test_case_response.replace_define_with_execute_hyperlink()
    result: str = test_case_api.get('test_plan_hyperlink')

    assert '/execute?planId=' in result


def test_user_password_with_colon(defined_hyperlink_example):
    """
    Assure User Password to access ADO API is formatted well with a colon.
    """
    test_case_response = schedule.TestCaseResponse(test_plan_hyperlink=defined_hyperlink_example)

    test_case_response.store_personal_access_token()
    test_case_api: dict = test_case_response.store_user_password()
    result: str = test_case_api.get('user_password')

    assert ':' in result


def test_base64_user_password_decoding_as_str(defined_hyperlink_example):
    """
    Azure DevOps passwords for some reason like base64 encoding before decoding.
    """
    test_case_response = schedule.TestCaseResponse(test_plan_hyperlink=defined_hyperlink_example)
    test_case_response.store_personal_access_token()
    test_case_response.store_user_password()
    test_case_response.encode_base64_user_password()

    test_case_api: dict = test_case_response.decode_base64_user_password()
    result: bytes = test_case_api.get('decode_user_password')

    assert type(result) is str


def test_base64_user_password_encoding_as_bytes():
    """
    Azure DevOps passwords decoded from encoded.  So user password should no longer be encoded as bytes.
    """
    test_case_response = schedule.TestCaseResponse(test_plan_hyperlink=defined_hyperlink_example)
    test_case_response.store_personal_access_token()
    test_case_response.store_user_password()

    test_case_api: dict = test_case_response.encode_base64_user_password()
    result: str = test_case_api.get('encode_user_password')

    assert type(result) is bytes


def test_get_test_case_response(defined_hyperlink_example: str):
    """
    Assure getting 200 response from test case request
    """
    test_case_response = schedule.TestCaseResponse(test_plan_hyperlink=defined_hyperlink_example)

    test_case_response.setup_api_credentials()
    result = test_case_response.request_test_case()

    assert result.status_code == 200

# def test_get_json_part(defined_hyperlink_example: str):
#     """
#
#     """
#     schedule.TestCaseResponse(test_plan_hyperlink=defined_hyperlink_example).main_method()
#     result: str = schedule.ParseTestCaseJSON(test_case_data_example)
#
#     assert len(result) == 5 and result.isdigit()


# def test_get_json_from_html(test_case_data_example):
#     assert
