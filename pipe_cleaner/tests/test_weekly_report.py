"""
Checks for weekly report dealing with presentation sent to the Client
"""

from pipe_cleaner.src.dashboard import weekly_report


def test_clean_assigned_to_for_empty() -> None:
    """
    Check TRR check out.
    """
    assigned_to: str = ""

    result: str = weekly_report.clean_assigned_to_name(assigned_to)

    assert "None" == result


def test_clean_assigned_to_for_period() -> None:
    """
    Check if period in name.
    """
    assigned_to: str = "joe.ton"

    result: str = weekly_report.clean_assigned_to_name(assigned_to)

    assert "." not in result


def test_clean_assigned_to_for_space() -> None:
    """
    Check if space in name.
    """
    assigned_to: str = "joe ton"

    result: str = weekly_report.clean_assigned_to_name(assigned_to)

    assert " " in result


def test_clean_assigned_to_for_first_name_capitalization() -> None:
    """
    Check capitalization for first name.
    """
    assigned_to: str = "joe ton"

    result: str = weekly_report.clean_assigned_to_name(assigned_to)
    first_name: str = result.split(" ")[0]
    first_letter_of_name: str = first_name[0]

    assert first_letter_of_name.istitle()


def test_clean_assigned_to_for_last_name_capitalization() -> None:
    """
    Check capitalization for last name.
    """
    assigned_to: str = "joe ton"

    result: str = weekly_report.clean_assigned_to_name(assigned_to)
    last_name: str = result.split(" ")[-1]
    last_letter_of_name: str = last_name[0]

    assert last_letter_of_name.istitle()
