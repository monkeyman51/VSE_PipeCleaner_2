"""
Update MongoDB database via Console Server ticket fields.
"""


from pymongo import MongoClient

from pipe_cleaner.src.data_ado import main_method as get_all_ticket_data
from pipe_cleaner.src.data_console_server import main_method as get_console_server_data


def get_mongodb_client(username: str, password: str, database: str) -> MongoClient:
    """
    Get client based on username, password, and database name.  SSL is true with no certification needed.
    """
    return MongoClient(f"mongodb+srv://{username}:{password}@cluster0.fueyc.mongodb.net/"
                       f"{database}?retryWrites=true&w=majority",
                       ssl=True, ssl_cert_reqs="CERT_NONE")


def get_mongodb_document(client: MongoClient, collection_name: str, document_name: str) -> MongoClient:
    """
    Get specific record based on client, collection name, and document name.
    """
    return client[collection_name][document_name]


def clean_part_number(part_number: str) -> str:
    """
    Assure part number is consistent.
    """
    if '(' in part_number:
        main_part: str = part_number.split('(')[-1]
        return part_number.upper().strip().replace(' ', '').replace(main_part, '').replace('(', '')

    else:
        return part_number.upper().strip().replace(' ', '')


def clean_request_type(request_type: str) -> str:
    """
    Assures request type ignores other unrelated descriptions.
    """
    return request_type.upper().replace(' TEST', '').replace('TEST', '')


def get_previous_comma_from_description(description: str, current_index: int) -> int:
    """
    Get the comma before the memory size in the description.
    """
    count: int = 1
    while count < 11:
        position: int = current_index - count
        current_character: str = description[position]

        if position >= 0 and current_character == ',':
            return position

        count += 1

    else:
        return 0


def get_size_from_description(description: str) -> str:
    """
    Get memory size from description.
    """
    description: str = clean_description_from_ado(description)

    if check_empty_description_name(description):
        return 'None'

    if 'TB' in get_memory_name(0, description) or 'GB' in get_memory_name(0, description):
        print(f'description: {description}')
        return get_clean_memory_name(description, 0, 0)

    else:

        for index, character in enumerate(description, start=0):

            if ',' in character:

                memory_name: str = get_memory_name(index, description)

                if 'TB' in memory_name or 'GB' in memory_name:
                    return get_clean_memory_name(description, index, 1)


def check_empty_description_name(description: str) -> bool:
    description: str = description.replace(',', '').upper()

    if not description:
        return True

    elif description == '':
        return True

    elif description == 'NONE':
        return True

    else:
        return False


def get_clean_memory_name(description: str, index: int, next_position: int) -> str:
    """
    Attempts to get memory size name from description field from ADO.  Accounts for different description possibilities.
    """
    first_position: int = get_previous_comma_from_description(description, index) + next_position
    if next_position == 0:
        print(f'first_position: {first_position}\n')
    max_count: int = index - first_position

    memory_characters: list = []
    count: int = 0
    while count < max_count:
        memory_characters.append(description[count + first_position])
        count += 1
    return clean_memory_size_name(memory_characters)


def get_memory_name(index: int, description: str) -> str:
    """
    Memory Name
    """
    return f'{description[index - 2]}{description[index - 1]}'


def clean_memory_size_name(memory_characters: list) -> str:
    """
    Joins characters to name the memory size.  Gives spaces between memory type and number
    """
    return ''.join(memory_characters).replace('TB', ' TB').replace('GB', ' GB')


def clean_description_from_ado(description: str) -> str:
    """
    Clean the description field from Azure DevOps for memory name parsing.
    """
    return description.upper().replace(' ', ',').replace(',TB,', 'TB,').replace(',GB,', 'GB,')


def update_part_numbers(azure_devops_data: dict, document: MongoClient):

    for ticket_id in azure_devops_data:

        if str(ticket_id).isdigit():
            table_data: dict = azure_devops_data[ticket_id].get('table_data', {})

            part_number_source: str = table_data.get('part_number', 'None')

            if part_number_source == 'None':
                pass

            else:
                request_type_source: str = table_data.get('request_type', 'None')
                supplier_source: str = table_data.get('supplier', 'None')
                description_source: str = table_data.get('description', 'None')
                memory_source: str = get_size_from_description(description_source)

                part_number_source: str = clean_part_number(part_number_source)
                request_type_source: str = clean_request_type(request_type_source)

                record_database: dict = document.find_one({'_id': part_number_source})

                if not record_database:
                    document.insert_one({'_id': part_number_source,
                                         'type': request_type_source,
                                         'supplier': supplier_source,
                                         'memory': memory_source,
                                         'ticket_id': [ticket_id]})
                else:
                    update_database_commodity_type(document, part_number_source, record_database)


def update_database_commodity_type(document: MongoClient, part_number_source: str, record_database: dict) -> None:
    """
    Update the database for the commodity type ie. SSD, NVMe, HDD, etc.
    """
    db_request_type: str = record_database['type']
    if db_request_type == 'None':
        document.update_one({'_id': part_number_source},
                            {'$set': {'type': part_number_source}})


def get_part_number_document():
    client = get_mongodb_client('joton51', 'FordFocus24', 'test')
    return get_mongodb_document(client, 'part_numbers', 'azure_devops')


def get_azure_devops_data():
    console_server_data: dict = get_console_server_data()
    azure_devops_data: dict = get_all_ticket_data(console_server_data)
    return azure_devops_data


def main_method():
    document: MongoClient = get_part_number_document()
    azure_devops_data: dict = get_azure_devops_data()

    update_part_numbers(azure_devops_data, document)
