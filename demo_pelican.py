"""
Attempt to get Pelican data.
"""

import requests

# full_name: str = "https://mySite.officeclimatecontrol.net/api.cgi?username=myname@gmail.com&password=mypassword&request=get&object=ThermostatHistory&selection=startDateTime:2021-10-01T00:00:00;endDateTime:2021-10-03T00:00:00&&value=name;temperature"


def get_api_request(raw_username: str, raw_password: str) -> str:
    """
    Attempt to make Pelican API work.
    :return: full API string
    """
    base_url: str = f"https://veritaskirkland.officeclimatecontrol.net/api.cgi?"
    default_thermal: str = "myThermostatName"

    # Pelican API requires 6 parts
    username: str = f"username={raw_username}&"
    password: str = f"password={raw_password}&"
    request: str = "request=get&"
    # request_object: str = "object=Thermostat&"
    request_object: str = "object=ThermostatHistory&"
    # selection: str = f"selection=startDateTime:2021-10-01T00:00:00;endDateTime:2021-10-03T00:00:00&"
    selection: str = f"selection=startDateTime:2021-10-01-24T00:00;endDateTime:2021-10-02T23:59&"
    value: str = "value=name;temperature;timestamp"

    return f"{base_url}{username}{password}{request}{request_object}{selection}{value}"


def main() -> None:
    """
    Attempt to get Pelican API.
    """
    raw_username: str = "joe.ton@vsei.com"
    raw_password: str = "FordFocus24"

    api_request: str = get_api_request(raw_username, raw_password)

    response_content = requests.get(api_request).text
    print(response_content)


main()
