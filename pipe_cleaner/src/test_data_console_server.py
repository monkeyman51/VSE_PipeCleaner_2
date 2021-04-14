import data_console_server as data


def test_clean_username_is_none_false():
    """
    Checks for username as none
    """
    empty_variance: str = 'none'

    result: bool = data.check_username_is_empty(empty_variance)

    assert result is False


def test_clean_username_is_empty_false():
    """
    Checks for username as empty
    """
    empty_variance: str = ''

    result: bool = data.check_username_is_empty(empty_variance)

    assert result is False


def test_clean_username_is_empty_true():
    """
    Checks method for True output
    """
    personal_name: str = 'joe_ton'

    result: bool = data.check_username_is_empty(personal_name)

    assert result is True


# def test_check_username_in_userbase():
#     """
#
#     """
#     personal_name: str = 'joe_ton'
#
#     result: bool = data.check_username_in_userbase(personal_name)
#
#     assert True
