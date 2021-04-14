from getpass import getuser


def get_username() -> str:
    """
    Gets user from network information after logging in successfully
    :return:
    """
    return getuser()


def convert_username(username: str) -> dict:
    """
    Converts Username to actual looking name.
    :param username: user name from network logged to Z: Drive
    :return: If no period separator, then returns default user_name
    """
    username_info: dict = {}

    if '.' in username:
        first_name = username.split('.')[0].capitalize()
        last_name = username.split('.')[1].capitalize()
        full_name = f'{first_name} {last_name}'

        username_info['first_name'] = first_name
        username_info['last_name'] = last_name
        username_info['full_name'] = full_name

        return username_info

    else:
        username_info['default_username'] = username
        username_info['full_name'] = username

        return username_info
