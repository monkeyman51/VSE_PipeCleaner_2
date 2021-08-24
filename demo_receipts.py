"""
Grab receipts from transaction database from 7/13 onward.
"""
from pipe_cleaner.src.log_database import access_database_document


def access_transactions() -> None:
    document = access_database_document('transactions', '021')
    cursor = document.find({})

    for current_document in cursor:
        part_number: str = current_document["part_number"]

        if "18ASF4G72PZ-2G9E1" in part_number:
            import json
            foo = json.dumps(current_document, sort_keys=True, indent=4)
            print(foo)
            input()
            scanned_serials: list = current_document["scanned"]
            for serial in scanned_serials:
                if "24C0A12A" in serial:
                    pass
                    # import json
                    # foo = json.dumps(current_document, sort_keys=True, indent=4)
                    # print(foo)
                    # input()

        current_location: str = current_document["location"]["current"].upper()
        previous_location: str = current_document["location"]["previous"].upper()

        date_logged: str = current_document["time"]["date_logged"]
        scanned_serials: list = current_document["scanned"]
        # total_count += current_count

    # print(total_count)
        # import json
        # foo = json.dumps(item, sort_keys=True, indent=4)
        # print(foo)
        # input()
        # print(item)
        # current_count: int = len(item["scanned"])
        # print(current_count)


def main_method() -> None:
    """
    Grab total of parts for receipts.
    """
    access_transactions()


main_method()
