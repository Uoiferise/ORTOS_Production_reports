import openpyxl
from excel_options import COLUMN_INDEXES_INFO_DATA
from services.basic_report.basic_nomenclature import Nomenclature
from excel_handler.production_plan_handler import create_production_date


def read_main_file(main_file: str) -> dict:
    main_dict = dict()

    input_book = openpyxl.load_workbook(main_file)
    input_sheet = input_book.active
    for row in range(4, input_sheet.max_row + 1):
        nomenclature_name = input_sheet.cell(row=row, column=5).value
        nomenclature_info = dict()
        for index, col in enumerate(COLUMN_INDEXES_INFO_DATA):
            nomenclature_info[index + 1] = input_sheet.cell(row=row, column=col).value
        nomenclature_info[max(nomenclature_info.keys()) + 1] = create_production_date(input_sheet=input_sheet, row=row)
        main_dict[row] = Nomenclature(name=nomenclature_name, id_row=row, info=nomenclature_info)

    return main_dict


def read_unshipped_file(unshipped_file: str) -> dict:
    pass


def read_data(main_file: str, unshipped_file: str) -> dict:
    data = read_main_file(main_file=main_file)
    result = read_unshipped_file(unshipped_file=unshipped_file)

    return data
