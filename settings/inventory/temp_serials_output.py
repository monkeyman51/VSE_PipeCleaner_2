from openpyxl import load_workbook
from pipe_cleaner.src.log_database import access_database_document


def get_serial_numbers_database():
    """
    Get serial numbers for all records of Kirkland.
    :return:
    """
    document = access_database_document("serial_numbers", 'version_02')
    return document.find({})


def clean_cell(data: str) -> str:
    return ", ".join(data)


def main_method() -> None:
    """

    :return:
    """
    database: list = get_serial_numbers_database()
    workbook = load_workbook("serials_database_output.xlsx")
    worksheet = workbook["Sheet1"]

    worksheet[f"A1"].value = "Serial Number"
    worksheet[f"B1"].value = "Part Number"
    worksheet[f"C1"].value = "Supplier"
    worksheet[f"D1"].value = "In House"
    worksheet[f"E1"].value = "From Locations"
    worksheet[f"F1"].value = "To Locations"
    worksheet[f"G1"].value = "Times"
    worksheet[f"H1"].value = "Dates"

    for index, entry in enumerate(database, start=2):
        print(index)
        serial_number: str = entry["_id"]

        if " " in serial_number:
            print(f"--> {serial_number}")

        part_numbers: str = ', '.join(entry["part_numbers"])
        suppliers: str = clean_cell(entry["suppliers"])
        in_house: str = clean_cell(entry["in_house"])
        from_locations: str = clean_cell(entry["from_locations"])
        to_locations: str = clean_cell(entry["to_locations"])
        times: str = clean_cell(entry["times"])
        dates: str = clean_cell(entry["dates"])

        worksheet[f'A{index}'].value = serial_number
        worksheet[f'B{index}'].value = part_numbers
        worksheet[f'C{index}'].value = suppliers
        worksheet[f'D{index}'].value = in_house
        worksheet[f'E{index}'].value = from_locations
        worksheet[f'F{index}'].value = to_locations
        worksheet[f'G{index}'].value = times
        worksheet[f'H{index}'].value = dates

    workbook.save("sn_database.xlsx")


main_method()
