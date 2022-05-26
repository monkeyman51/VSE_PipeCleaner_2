"""
Attempt to get Pelican data.
"""

import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from datetime import datetime, timedelta
from time import time


def get_current_timestamp() -> str:
    """

    """
    current_time = str(datetime.now()).replace(" ", "T")[0:16]
    return f"{current_time}:00"


def get_year_ago_timestamp() -> str:
    """

    """
    year_ago = str(datetime.now() - timedelta(weeks=4)).replace(" ", "T")[0:16]
    return f"{year_ago}:00"


def get_month_date(month: str, year: str) -> str:
    """

    :param month: Three Letter string
    :param year: Four Letter string
    :return: API string
    """
    if month == "JAN":
        return f"selection=startDateTime:{year}-01-01T00:00:00;endDateTime:{year}-01-31T00:00:00&"

    elif month == "FEB":
        return f"selection=startDateTime:{year}-02-01T00:00:00;endDateTime:{year}-02-28T00:00:00&"

    elif month == "MAR":
        return f"selection=startDateTime:{year}-03-01T00:00:00;endDateTime:{year}-03-31T00:00:00&"

    elif month == "APR":
        return f"selection=startDateTime:{year}-04-01T00:00:00;endDateTime:{year}-04-30T00:00:00&"

    elif month == "MAY":
        return f"selection=startDateTime:{year}-05-01T00:00:00;endDateTime:{year}-05-31T00:00:00&"

    elif month == "JUN":
        return f"selection=startDateTime:{year}-06-01T00:00:00;endDateTime:{year}-06-30T00:00:00&"

    elif month == "JUL":
        return f"selection=startDateTime:{year}-07-01T00:00:00;endDateTime:{year}-07-31T00:00:00&"

    elif month == "AUG":
        return f"selection=startDateTime:{year}-08-01T00:00:00;endDateTime:{year}-08-31T00:00:00&"

    elif month == "SEP":
        return f"selection=startDateTime:{year}-09-01T00:00:00;endDateTime:{year}-09-30T00:00:00&"

    elif month == "OCT":
        return f"selection=startDateTime:{year}-10-01T00:00:00;endDateTime:{year}-10-31T00:00:00&"

    elif month == "NOV":
        return f"selection=startDateTime:{year}-11-01T00:00:00;endDateTime:{year}-11-30T00:00:00&"

    elif month == "DEC":
        return f"selection=startDateTime:{year}-12-01T00:00:00;endDateTime:{year}-12-31T00:00:00&"

    else:
        return "None"


def get_api_request(raw_username: str, raw_password: str) -> str:
    """
    Attempt to make Pelican API work.
    :return: full API string
    """
    base_url: str = f"https://veritaskirkland.officeclimatecontrol.net/api.cgi?"
    default_thermal: str = "myThermostatName"
    current_timestamp: str = get_current_timestamp()
    year_ago_timestamp: str = get_year_ago_timestamp()
    # print(f"{current_timestamp}")
    # print(f"{year_ago_timestamp}")
    # print(f"2021-10-01T00:00:00")

    # Pelican API requires 6 parts
    username: str = f"username={raw_username}&"
    password: str = f"password={raw_password}&"
    request: str = "request=get&"
    # request_object: str = "object=Thermostat&"
    request_object: str = "object=ThermostatHistory&"

    # selection: str = get_month_date("DEC", "2020")
    # selection: str = get_month_date("JAN", "2021")
    # selection: str = get_month_date("FEB", "2021")
    # selection: str = get_month_date("MAR", "2021")
    # selection: str = get_month_date("APR", "2021")
    # selection: str = get_month_date("MAY", "2021")
    # selection: str = get_month_date("JUN", "2021")
    # selection: str = get_month_date("JUL", "2021")
    # selection: str = get_month_date("AUG", "2021")
    # selection: str = get_month_date("SEP", "2021")
    # selection: str = get_month_date("OCT", "2021")
    selection: str = get_month_date("NOV", "2021")

    # selection: str = f"selection=startDateTime:2021-11-01T00:00:00;endDateTime:2021-11-29T00:00:00&"
    # selection: str = f"selection=startDateTime:{year_ago_timestamp};endDateTime:{current_timestamp}&"
    value: str = "value=slaves;temperature;timestamp;coolSetting;heatSetting"

    return f"{base_url}{username}{password}{request}{request_object}{selection}{value}"


def get_thermostat_history(api_request: str) -> list:
    """

    """
    response_content = requests.get(api_request).text
    print(f"response_content: {response_content}")
    soup = BeautifulSoup(response_content, "html.parser")
    thermostat_history: list = soup.find_all("thermostathistory")
    return thermostat_history


def clean_slave_name(slave_name: str) -> str:
    """
    Clean slave name.
    """
    clean_name: str = slave_name.replace(r"amp;", "")
    return clean_name


def get_remote_temperature(remote_data: list, sensor_name: str) -> dict:
    """
    Parse from XML tag
    """
    real_sensor_name: str = replace_sensor_name(sensor_name)

    if len(remote_data) >= 2:
        for slave_data in remote_data:
            slave_name: str = slave_data.find("name").text
            slave_value: str = slave_data.find("value").text

            clean_name: str = clean_slave_name(slave_name)

            if sensor_name != clean_name:
                return {"sensor_name": real_sensor_name, "remote_name": slave_name, "remote_value": slave_value}
        else:
            return {"sensor_name": real_sensor_name, "remote_name": "None", "remote_value": "None"}

    elif len(remote_data) == 0:
        return {"sensor_name": real_sensor_name, "remote_name": "None", "remote_value": "None"}

    elif len(remote_data) == 1:
        remote_name: str = remote_data[0].find("name").text
        remote_value: str = remote_data[0].find("value").text
        return {"sensor_name": real_sensor_name, "remote_name": remote_name, "remote_value": remote_value}

    else:
        return {"sensor_name": real_sensor_name, "remote_name": "None", "remote_value": "None"}


def replace_sensor_name(sensor_name: str) -> str:
    """
    Replace fake name with real sensor name.
    """
    if sensor_name == "AC-1 Supply Temp":
        return "Einstein Lab AC-1"

    elif sensor_name == "AC-2 Supply Temp":
        return "Einstein Lab AC-2"

    elif sensor_name == "AC-3 Supply Temp":
        return "Einstein Lab AC-3"

    elif sensor_name == "AC-4 Supply Temp":
        return "Einstein Lab AC-4"

    else:
        return sensor_name


def get_temperature_history(thermostat_history: list) -> list:
    """

    :param thermostat_history:
    :return:
    """
    temperature_history: list = []
    for index, thermostat_data in enumerate(thermostat_history, start=1):

        try:
            sensor_name: str = thermostat_data.find("name").text
            temperature_logs: str = thermostat_data.find_all("history")

            for temperature_instance in temperature_logs:
                temperature = temperature_instance.find("temperature").text
                timestamp = temperature_instance.find("timestamp").text
                cool_setting = temperature_instance.find("coolsetting").text
                heat_setting = temperature_instance.find("heatsetting").text
                remote_data = temperature_instance.find_all("slaves")
                temperature_case: dict = {"name": sensor_name,
                                          "temperature": temperature,
                                          "cool_setting": cool_setting,
                                          "heat_setting": heat_setting,
                                          "remote_temperature": get_remote_temperature(remote_data, sensor_name),
                                          "timestamp": timestamp}
                temperature_history.append(temperature_case)

        except AttributeError:
            # Sometimes have empty thermostat history
            pass
    return temperature_history


def get_date_from_timestamp(timestamp: str) -> str:
    """
    Parse date from temperature timestamp.
    """
    return timestamp[0:10]


def get_time_from_timestamp(timestamp: str) -> str:
    """
    Parse date from temperature timestamp.
    """
    return timestamp[11:16]


def add_excel_history_data(temperature_history: list) -> None:
    """
    Add temperature history data to the excel output.
    """
    workbook = load_workbook("settings/temperature_history_template.xlsx")
    worksheet = workbook["Sheet1"]

    for index, temperature_log in enumerate(temperature_history, start=2):

        temperature: str = temperature_log["temperature"]
        cool_setting: str = temperature_log["cool_setting"]
        heat_setting: str = temperature_log["heat_setting"]
        timestamp: str = temperature_log["timestamp"]
        remote_temperature: dict = temperature_log["remote_temperature"]

        worksheet[f"A{index}"].value = get_date_from_timestamp(timestamp)
        worksheet[f"B{index}"].value = get_time_from_timestamp(timestamp)
        worksheet[f"C{index}"].value = remote_temperature["sensor_name"]
        worksheet[f"D{index}"].value = temperature
        worksheet[f"E{index}"].value = remote_temperature["remote_name"]
        worksheet[f"F{index}"].value = remote_temperature["remote_value"]
        worksheet[f"G{index}"].value = cool_setting
        worksheet[f"H{index}"].value = heat_setting

    workbook.save("temperature_history.xlsx")


def main() -> None:
    """
    Attempt to get Pelican API.
    """
    start = time()
    raw_username: str = "joe.ton@vsei.com"
    raw_password: str = "FordFocus24"

    api_request: str = get_api_request(raw_username, raw_password)
    thermostat_history: list = get_thermostat_history(api_request)
    temperature_history: list = get_temperature_history(thermostat_history)

    add_excel_history_data(temperature_history)
    end = time()

    print(end - start)


main()
