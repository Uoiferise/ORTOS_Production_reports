import openpyxl
from excel_options import COLUMN_INDEXES_INFO_DATA


class Nomenclature:
    def __init__(self, name: str, id_row: int, info: list):
        self.name = name
        self.id_row = id_row
        self.info = info


def read_main_file(main_file: str) -> dict:
    data = {
        'main_dict': {},
        'unshipped_dict': {},
        'date_dict': {}
    }

    input_book = openpyxl.load_workbook(main_file)
    input_sheet = input_book.active
    for row in range(4, input_sheet.max_row + 1):
        nomenclature_name = input_sheet.cell(row=row, column=5).value
        nomenclature_info = [input_sheet.cell(row=row, column=col).value for col in COLUMN_INDEXES_INFO_DATA]
        data['main_dict'][nomenclature_name] = Nomenclature(name=nomenclature_name, id_row=row, info=nomenclature_info)

    return data


def read_unshipped_file(unshipped_file: str) -> dict:
    pass


def read_data(main_file: str, unshipped_file: str) -> dict:
    data = read_main_file(main_file=main_file)
    result = read_unshipped_file(unshipped_file=unshipped_file)

    return data
