"""
Total Kirkland based on Rich's data plus inventory tool database.
"""
from pipe_cleaner.src.log_database import access_database_document


def audit_transaction_date() -> list:
    """
    Lists of transaction dates to skip when iterating through all transactions.
    :return: dates to avoid.
    """
    return ["06/23/2021", "07/07/2021", "07/08/2021", "07/09/2021", "07/12/2021", "07/13/2021", "07/14/2021",
            "07/15/2021", "07/16/2021", "07/19/2021", "07/20/2021", "07/21/2021", "07/22/2021", "07/23/2021",
            "07/26/2021", "07/27/2021", "07/28/2021", "07/29/2021", "07/30/2021", "08/02/2021", "08/03/2021",
            "08/04/2021", "08/05/2021", "08/06/2021", "08/09/2021", "08/10/2021", "08/11/2021"]


def is_date_in_audit(current_date: str) -> bool:
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


def setup_transaction_data(inbound_name: str, outbound_name: str) -> dict:
    """
    Creates consistent naming convention for inbound and outbound names.
    :return: High-level structure of transactions of IN / OUT
    """
    return {inbound_name: {}, outbound_name: {}}


def get_transactions() -> dict:
    """
    From inventory database
    :return:
    """
    transaction_data: dict = setup_transaction_data("receipt", "shipment")

    for current_entry in access_inventory_database():
        transaction_data: dict = get_correct_transaction(current_entry, transaction_data)

    return transaction_data


def access_inventory_database() -> list:
    """
    Access and return all transaction entries from Inventory.
    :return: List of transaction entries.
    """
    document = access_database_document('transactions', '021')
    return document.find({})


def get_correct_transaction(current_entry: dict, transaction_data: dict) -> dict:
    """
    Get audited transaction logs.
    :param transaction_data: all transaction data
    :param current_entry: transaction log
    :return: combines current transaction with all transactions
    """
    current_date: str = current_entry["time"]["date_logged"]

    if is_date_in_audit(current_date):
        part_number: str = current_entry["part_number"]
        previous: str = current_entry["location"]["previous"]
        print(f'previous: {previous}')
        current: str = current_entry["location"]["current"]
        quantity: int = len(current_entry["scanned"])

        if "Receipt" in previous:
            return add_transaction_data(part_number, quantity, transaction_data)

        elif "Shipment" in current or "Customer" in current:
            return subtract_transaction_data(part_number, quantity, transaction_data)

        else:
            return transaction_data

    else:
        return transaction_data


def add_transaction_data(part_number: str, quantity: int, transaction_data: dict) -> dict:
    """
    Add to overall total based on transaction data.
    :param part_number: name of commodity
    :param quantity: quantity being moved
    :param transaction_data: total data
    :return: all transaction data
    """
    receipts: dict = transaction_data["receipt"]

    if part_number in receipts:
        receipts[part_number] += quantity

    else:
        receipts[part_number] = quantity

    return transaction_data


def subtract_transaction_data(part_number: str, quantity: int, transaction_data: dict) -> dict:
    """
    Minus to overall total based on transaction data.
    :param part_number: name of commodity
    :param quantity: quantity being moved
    :param transaction_data: total data
    :return: all transaction data
    """
    if part_number in transaction_data:
        transaction_data["shipment"][part_number] += quantity

    else:
        transaction_data["shipment"][part_number] = quantity
    return transaction_data


def main_method() -> None:
    """
    Extract data from Rich while adding / subtracting from database's transactions on plus or minus from Kirkland site.
    :return: None
    """
    transactions: dict = get_transactions()
    import json
    foo = json.dumps(transactions, sort_keys=True, indent=4)
    print(foo)
    input()
    #  Get Rich's dataset
    #  Consolidate Data
    #  Output


main_method()
