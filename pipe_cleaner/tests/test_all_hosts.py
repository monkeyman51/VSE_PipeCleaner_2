from pipe_cleaner.src import report_all_hosts as all_hosts


def test_machine_last_online_in_days_for_zero_input() -> None:
    """
    Console Server may return machines that have no information.
    """
    empty_input: float = 0.00

    result: str = all_hosts.convert_machine_last_online_to_days(empty_input)

    assert result == -1


def test_get_sku_name_for_empty_dictionary() -> None:
    """
    SKU name dictionary might be empty.
    """
    host_data: dict = {}

    result: str = all_hosts.get_sku_name(host_data)

    assert "None" in result


def test_get_machine_name_for_empty_dictionary() -> None:
    """
    Checks machine name for empty dictionary input.
    """
    host_data: dict = {}

    result: str = all_hosts.get_machine_name(host_data)

    assert result == ""


def test_get_machine_name_for_space_in_first_character() -> None:
    """
    Checks machine name for first character as space.
    """
    host_data: dict = {"machine_name": " VSE0G5IZXDR-303"}

    result: str = all_hosts.get_machine_name(host_data)

    assert result[0] != " "


def test_get_machine_name_for_space_in_last_character() -> None:
    """
    Checks machine name for first character as space.
    """
    host_data: dict = {"machine_name": "VSE0G5IZXDR-303 "}

    result: str = all_hosts.get_machine_name(host_data)

    assert result[-1] != " "


def test_get_last_found_online_for_empty_dictionary() -> None:
    """
    Checks to account for empty dictionary as input.
    """
    host_data: dict = {}

    result: float = all_hosts.get_last_found_online(host_data)

    assert result == 0.00


def test_get_machine_location_for_empty_dictionary() -> None:
    """
    Checks for empty dictinoary as input.
    """
    host_data: dict = {}

    result: str = all_hosts.get_machine_location(host_data)

    assert 'None' in result


def test_get_machine_location_for_uppercase() -> None:
    """
    Checks for uppercase.
    """
    host_data: dict = {"location": "r40U18N18"}

    result: str = all_hosts.get_machine_location(host_data)

    assert result.isupper()


def test_get_machine_location_for_space_in_first_character() -> None:
    """
    Checks for first character as space.
    """
    host_data: dict = {"location": " R40U18N18"}

    result: str = all_hosts.get_machine_location(host_data)

    assert result[0] != " "


def test_get_machine_location_for_space_in_last_character() -> None:
    """
    Checks for last character as space.
    """
    host_data: dict = {"location": "R40U18N18 "}

    result: str = all_hosts.get_machine_location(host_data)

    assert result[-1] != " "


def test_get_machine_name_for_vse_in_name() -> None:
    """
    Checks to see if VSE is in machine name.
    """
    host_data: dict = {"machine_name": "0G6IWBAL-021"}

    result: str = all_hosts.get_machine_name(host_data)

    assert "Invalid Name" in result


def test_get_machine_name_for_uppercase() -> None:
    """
    Checks to see if machine name is upper case.
    """
    host_data: dict = {"machine_name": "vse0G6IWBAL-021"}

    result: str = all_hosts.get_machine_name(host_data)

    assert result.isupper()


def test_get_machine_serial_for_uppercase() -> None:
    """
    Checks to see if machine serial number handles empty dictionary.
    """
    host_data: dict = {"baseboard": {"serial": "9J1000019220892J0G1"}}

    result: str = all_hosts.get_machine_serial(host_data)

    assert result.isupper()


def test_get_machine_serial_for_string_stripped() -> None:
    """
    Checks to see if machine serial number handles
    """
    host_data: dict = {"baseboard": {"serial": " 9J1000019220892J0G1 "}}

    result: str = all_hosts.get_machine_serial(host_data)

    assert result[0] != " " and result[-1] != " "


def test_get_machine_serial_for_empty_dictionary() -> None:
    """
    Checks to see if machine serial is no space at first and last character.
    """
    host_data: dict = {}

    result: str = all_hosts.get_machine_serial(host_data)

    assert 'None' in result


def test_get_machine_name_for_gen_5_cma() -> None:
    """
    Checks CMA in machine name.
    """
    host_data: dict = {"machine_name": "CMA"}

    result: str = all_hosts.get_machine_name(host_data)

    assert "CMA" not in result


def test_get_machine_host_ip_for_empty_dictionary() -> None:
    """
    Checks empty input for dictionary.
    """
    host_data: dict = {}

    result: str = all_hosts.get_machine_host_ip(host_data)

    assert "None" in result


def test_get_machine_host_ip_for_dot_in_string() -> None:
    """
    Should have multiple dots within the string.
    """
    host_data: dict = {"net": {"interfaces": [{"ip": "192168237173"}]}}

    result: str = all_hosts.get_machine_host_ip(host_data)

    assert "None" in result


def test_get_machine_host_ip_for_non_digits() -> None:
    """
    Should have multiple dots within the string.
    """
    host_data: dict = {"net": {"interfaces": [{"ip": "ABC"}]}}

    result: str = all_hosts.get_machine_host_ip(host_data)

    assert "None" in result
