"""
2/24/2021

Module created to address not being able to fetch personal access token necessary to access to ADO for test cases.
Unable to access upper directories containing information for personal access token and API handle.
For that reason, created basic functions here to implement code for accessing test cases functionality.
"""


def get_personal_access_token() -> str:
    """
    Personal Access Token for API calls to Azure DevOps for Commodity Testing.
    """
    return 'oa3kit3be5dlk2kfbkylhb62qn2kc4ja3363c5iogir66k5bwrwq'


def get_url() -> str:
    """
    Personal Access Token for API calls for Azure DevOps for Testing.
    """
    pass
