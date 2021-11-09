"""
Get all history serials
"""
from pipe_cleaner.src.log_database import access_database_document
from openpyxl import load_workbook


def get_clean_alternatives(alternative, alternatives, serial_number) -> list:
    """

    :param alternative:
    :param alternatives:
    :param serial_number:
    :return:
    """
    clean_alternatives: list = [serial_number]
    for alt_name in alternatives:
        if alt_name != alternative:
            clean_alternatives.append(alt_name)

    return clean_alternatives


def access_serial_numbers_database() -> list:
    """
    Access and return serial numbers from Inventory database for testing and mocking.
    :return: List of transaction entries.
    """
    document = access_database_document('serial_numbers', 'base_line')
    return document.find({})


def clean_serial_number(serial_number: str, part_number: str) -> str:
    """

    :param serial_number:
    :param part_number:
    :return:
    """
    serial_number = str(serial_number).upper()

    if part_number in serial_number:
        clean_serial: str = serial_number.replace(f"{part_number}_", "").replace(f"{part_number} ", "").strip()
        return clean_serial
    else:
        return str(serial_number)


def clean_part_name(part_name: str) -> str:
    """
    Erase any known characters that do not belong in the part name.

    :param part_name: raw part name
    :return: clean part name
    """
    return part_name.replace(":", "").upper().strip()


def get_serial_numbers_database():
    """
    Get serial numbers for all records of Kirkland.
    :return:
    """
    document = access_database_document("serial_numbers", 'version_02')
    serial_numbers_database: list = document.find({})

    database: dict = {}
    for entry in serial_numbers_database:
        serial_number: str = entry["_id"].upper()
        current_location: str = entry["to_locations"][-1].title()
        previous_location: str = entry["from_locations"][-1].title()
        date_logged: str = entry["dates"][-1]
        part_number: str = clean_part_name(entry["part_numbers"][-1])

        clean_serial: str = clean_serial_number(serial_number, part_number)

        if clean_serial not in database:
            database[clean_serial]: dict = {}
            database[clean_serial]["current_location"]: str = current_location
            database[clean_serial]["previous_location"]: str = previous_location
            database[clean_serial]["date_logged"]: str = date_logged
            database[clean_serial]["part_number"]: str = part_number

    return database


def access_inventory_database() -> list:
    """
    Access and return all transaction entries from Inventory.
    :return: List of transaction entries.
    """
    document = access_database_document('transactions', '021')
    return document.find({})


def audit_transaction_date() -> list:
    """
    Lists of transaction dates to skip when iterating through all transactions.
    :return: dates to avoid.
    """
    return ["06/23/2021", "07/07/2021", "07/08/2021", "07/09/2021", "07/12/2021", "07/13/2021", "07/14/2021",
            "07/15/2021", "07/16/2021", "07/19/2021", "07/20/2021", "07/21/2021", "07/22/2021", "07/23/2021",
            "07/26/2021", "07/27/2021", "07/28/2021", "07/29/2021", "07/30/2021", "08/02/2021", "08/03/2021",
            "08/04/2021", "08/05/2021", "08/06/2021", "08/09/2021", "08/10/2021", "08/11/2021", "8/12/2021"]


def is_date_after(current_date: str) -> bool:
    """
    Iterates through dates that should not be included in transaction logs.
    :param current_date:
    :return: True / False
    """
    audit_dates: list = audit_transaction_date()

    for audit_date in audit_dates:
        if current_date in audit_date:
            return False
    else:
        return True


def get_transactions_from_database() -> list:
    """
    From inventory database
    :return:
    """
    transactions: list = []

    for current_entry in access_inventory_database():
        part_number: str = current_entry["part_number"]
        current_date: str = current_entry["time"]["date_logged"]

        if part_number.upper() != "TEST" and is_date_after(current_date):

            scanned: list = current_entry["scanned"]
            current_location: str = current_entry["location"]["current"]
            previous_location: str = current_entry["location"]["previous"]
            clean_part: str = clean_part_name(part_number)

            for serial_number in scanned:

                clean_serial: str = clean_serial_number(serial_number, part_number)

                current_serial: dict = {"part_number": clean_part,
                                        "serial_number": clean_serial,
                                        "current_location": current_location,
                                        "date_logged": current_date,
                                        "previous_location": previous_location}
                transactions.append(current_serial)

    return transactions


def update_serial_history(base_serials: dict, transactions: list) -> dict:
    """


    :param base_serials:
    :param transactions:
    :return:
    """
    for serial_entry in transactions:

        serial_number: str = serial_entry["serial_number"]
        part_number: str = serial_entry["part_number"]
        current_location: str = serial_entry["current_location"]
        previous_location: str = serial_entry["previous_location"]
        date_logged: str = serial_entry["date_logged"]

        if serial_number not in base_serials:

            base_serials[serial_number]: dict = {}
            base_serials[serial_number]["date_logged"]: str = date_logged
            base_serials[serial_number]["part_number"]: str = part_number
            base_serials[serial_number]["current_location"]: str = current_location
            base_serials[serial_number]["previous_location"]: str = previous_location

        elif serial_number in base_serials:

            base_serials[serial_number]["date_logged"]: str = date_logged
            base_serials[serial_number]["part_number"]: str = part_number
            base_serials[serial_number]["current_location"]: str = current_location
            base_serials[serial_number]["previous_location"]: str = previous_location

    return base_serials


def output_excel(update_serials: dict) -> None:
    """

    :return:
    """
    workbook = load_workbook("settings/inventory/serials_history_template.xlsx")
    worksheet = workbook["Sheet1"]

    for index, serial_number in enumerate(update_serials, start=2):
        serial_entry: dict = update_serials[serial_number]
        part_number: str = serial_entry["part_number"]
        current_location: str = serial_entry["current_location"]
        previous_location: str = serial_entry["previous_location"]
        date_logged: str = serial_entry["date_logged"]

        # if serial_number == part_number:
        #     print(f"serial_number: {serial_number}")

        if "_" in serial_number:
            print(f"serial_number: {serial_number}")

        worksheet[f"A{index}"].value = clean_serial_number(serial_number, part_number)
        worksheet[f"B{index}"].value = part_number
        worksheet[f"C{index}"].value = current_location
        worksheet[f"D{index}"].value = previous_location
        worksheet[f"E{index}"].value = date_logged

    workbook.save("serials_history.xlsx")


def main() -> None:
    """

    :return:
    """
    base_serial_numbers: dict = get_serial_numbers_database()
    transactions: list = get_transactions_from_database()
    update_serials: dict = update_serial_history(base_serial_numbers, transactions)
    output_excel(update_serials)

    # import json
    # foo = json.dumps(transactions, sort_keys=True, indent=4)
    # print(foo)
    # print(len(transactions))
    # input()


main()