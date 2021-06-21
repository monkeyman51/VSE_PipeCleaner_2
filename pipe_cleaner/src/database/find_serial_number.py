"""
Find serial number from database given serial number entered from user.  Print out data.
"""
from pipe_cleaner.src.log_database import access_database_document


def main_method(serial_number_entered: str) -> None:
    """

    """
    print(f'\n\tGetting serial numbers data from database...')
    all_serial_numbers: list = access_database_document('serial_numbers', 'all').find({})
    print(f'\tFinding serial number...')

    for serial_number_data in all_serial_numbers:
        serial_number_id: str = serial_number_data['_id']

        if serial_number_entered == serial_number_id:
            print(f'\t\n\n{"-"*60}\n\n')
            print(f'\tFound: {serial_number_entered}')
            print(f'\n\n\tSerial Number Data:')

            serial_number_id: str = serial_number_data['_id']

            print(f'\n\t\t- ID: {serial_number_id}')
            print(f'\t\t- Part Number: {serial_number_data["part_number"]}')

            transactions: list = serial_number_data['transactions']

            for index, transaction_entry in enumerate(transactions, start=0):

                print(f'\n\t\tTransaction #{index+1}')
                print(f'\t\t\t- Date: {transaction_entry["time"]["date_logged"]}')
                print(f'\t\t\t- Time: {transaction_entry["time"]["time_logged"]}')

                print(f'\n\t\t\t- Current: {transaction_entry["location"]["current"]}')
                print(f'\t\t\t- Previous: {transaction_entry["location"]["previous"]}')
                print(f'\t\t\t- Site: {transaction_entry["location"]["site"]}')
                print(f'\t\t\t- Rack: {transaction_entry["location"]["rack"]}')
                print(f'\t\t\t- Machine: {transaction_entry["location"]["machine"]}')
                print(f'\t\t\t- Pipe: {transaction_entry["location"]["pipe"]}')

                print(f'\n\t\t\t- Approved By: {transaction_entry["source"]["approved_by"]}')
                print(f'\t\t\t- Verified By: {transaction_entry["source"]["verified_by"]}')
                print(f'\t\t\t- Version: {transaction_entry["source"]["version"]}')
                print(f'\t\t\t- TRR: {transaction_entry["source"]["trr"]}')
                print(f'\t\t\t- Task: {transaction_entry["source"]["task"]}')
                print(f'\t\t\t- Comment: {transaction_entry["source"]["comment"]}')

            input(f'\n\tPress enter to return to menu:')

    else:
        print(f'\t\n\n{"-"*60}\n\n')
        print(f'\tNOT found.')
        input(f'\n\tPress enter to return to menu:')
