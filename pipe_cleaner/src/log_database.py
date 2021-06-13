"""
Log data inside MongoDB
"""
from pymongo import MongoClient


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


def access_database_document(database_name: str, document_name: str) -> MongoClient:
    """
    Get client from database
    """
    client = get_mongodb_client('joton51', 'FordFocus24', 'test')
    return get_mongodb_document(client, database_name, document_name)


def main_method():
    """
    Start Here.
    """
    pass
