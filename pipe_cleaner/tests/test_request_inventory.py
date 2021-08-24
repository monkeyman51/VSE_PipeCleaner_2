from pipe_cleaner.src import request_inventory as inventory


def test_default_username_for_ext() -> None:
    """
    Checks -EXT in username for contractors.
    """
    default_username: str = 'joe.ton-EXT'

    result: str = inventory.clean_default_username(default_username)

    assert '-EXT' not in result.upper()


def test_default_username_for_none() -> None:
    """
    Checks empty default username.
    """
    default_username: str = ''

    result: str = inventory.clean_default_username(default_username)

    assert 'None' in result


def test_default_username_for_non_string_input() -> None:
    """
    Checks for non-string input.
    """
    default_username: str = '0'

    result: str = inventory.clean_default_username(default_username)

    assert 'None' in result


def test_default_username_for_first_character_being_empty() -> None:
    """
    Checks to make sure there is no space before the name.
    """
    default_username: str = ' joe.ton-EXT'

    result: str = inventory.clean_default_username(default_username)

    assert result[0] != ' '


def test_default_username_for_last_character_being_empty() -> None:
    """
    Checks to make sure last character is not empty.
    """
    default_username: str = 'joe.ton-EXT '

    result: str = inventory.clean_default_username(default_username)

    assert result[-1] != ' '