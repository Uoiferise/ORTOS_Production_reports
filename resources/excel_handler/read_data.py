import openpyxl
from operator import add


def read_input_files(main_file, unshipped_file, options=None):
    input_book = openpyxl.load_workbook(main_file)
    input_sheet = input_book.active

    # Отбор не архивных номенклатур в main_file
    row_indexes = []
    for row in range(4, input_sheet.max_row + 1):
        if input_sheet.cell(row=row, column=6).value != 'Да' and input_sheet.cell(row=row, column=7).value != 'Да':
            row_indexes.append(row)

    # Отбор арх. поз. titanium_base, которые надо учесть
    if main_file == 'data/titanium_base/titanium_base_info.xlsx':
        tb_actual_book = openpyxl.load_workbook('data/titanium_base/titanium_base_actual.xlsx')
        tb_actual_sheet = tb_actual_book.active
        tb_actual_dict = {}
        for r in range(2, 20):
            tb_actual_dict[tb_actual_sheet.cell(row=r, column=1).value] = \
                tb_actual_dict.get(tb_actual_sheet.cell(row=r, column=1).value, []) + \
                [tb_actual_sheet.cell(row=r, column=2).value]

        tb_ar_dict = {}
        for row in range(4, input_sheet.max_row + 1):
            if input_sheet.cell(row=row, column=5).value in list(tb_actual_dict.values())[0]:
                info = []
                for col in (9, 12, 16, 17, 19, 20, 22, 23, 25, 26):
                    if input_sheet.cell(row=row, column=col).value is None:
                        x = 0
                    else:
                        x = int(input_sheet.cell(row=row, column=col).value)
                    info.append(x)

                tb_ar_dict[list(tb_actual_dict.keys())[0]] = \
                    list(
                        map(
                            add,
                            tb_ar_dict.get(list(tb_actual_dict.keys())[0], [0 for _ in range(10)]),
                            info
                        )
                    )
            elif input_sheet.cell(row=row, column=5).value in list(tb_actual_dict.values())[1]:
                info = []
                for col in (9, 12, 16, 17, 19, 20, 22, 23, 25, 26):
                    if input_sheet.cell(row=row, column=col).value is None:
                        x = 0
                    else:
                        x = int(input_sheet.cell(row=row, column=col).value)
                    info.append(x)

                tb_ar_dict[list(tb_actual_dict.keys())[1]] = \
                    list(
                        map(
                            add,
                            tb_ar_dict.get(list(tb_actual_dict.keys())[1], [0 for _ in range(10)]),
                            info
                        )
                    )
            elif input_sheet.cell(row=row, column=5).value in list(tb_actual_dict.values())[2]:
                info = []
                for col in (9, 12, 16, 17, 19, 20, 22, 23, 25, 26):
                    if input_sheet.cell(row=row, column=col).value is None:
                        x = 0
                    else:
                        x = int(input_sheet.cell(row=row, column=col).value)
                    info.append(x)

                tb_ar_dict[list(tb_actual_dict.keys())[2]] = \
                    list(
                        map(
                            add,
                            tb_ar_dict.get(list(tb_actual_dict.keys())[2], [0 for _ in range(10)]),
                            info
                        )
                    )
            elif input_sheet.cell(row=row, column=5).value in list(tb_actual_dict.values())[3]:
                info = []
                for col in (9, 12, 16, 17, 19, 20, 22, 23, 25, 26):
                    if input_sheet.cell(row=row, column=col).value is None:
                        x = 0
                    else:
                        x = int(input_sheet.cell(row=row, column=col).value)
                    info.append(x)

                tb_ar_dict[list(tb_actual_dict.keys())[3]] = \
                    list(
                        map(
                            add,
                            tb_ar_dict.get(list(tb_actual_dict.keys())[3], [0 for _ in range(10)]),
                            info
                        )
                    )

    # Дополнительный исключения номенклатур
    row_indexes_copy = row_indexes.copy()
    if main_file == 'data/screws/screws_info.xlsx':
        for i, row in enumerate(row_indexes_copy):
            if 'блистер' in input_sheet.cell(row=row, column=5).value.lower():
                del row_indexes[row_indexes.index(row)]
            elif 'упак' in input_sheet.cell(row=row, column=5).value.lower():
                del row_indexes[row_indexes.index(row)]
            elif 'проб' in input_sheet.cell(row=row, column=5).value.lower():
                del row_indexes[row_indexes.index(row)]
    elif main_file == 'data/analog/analog_info.xlsx':
        for i, row in enumerate(row_indexes_copy):
            if 'нерж' not in input_sheet.cell(row=row, column=5).value.lower():
                del row_indexes[row_indexes.index(row)]
    elif main_file == 'data/scan_body/scan_body_info.xlsx':
        for i, row in enumerate(row_indexes_copy):
            if 'нерж' in input_sheet.cell(row=row, column=5).value.lower() and \
                    'б' in input_sheet.cell(row=row, column=5).value.split()[0]:
                del row_indexes[row_indexes.index(row)]
            elif 'латунь' in input_sheet.cell(row=row, column=5).value.lower() and \
                    'б' in input_sheet.cell(row=row, column=5).value.split()[0]:
                del row_indexes[row_indexes.index(row)]
    elif main_file == 'data/titanium_base/titanium_base_info.xlsx':
        for i, row in enumerate(row_indexes_copy):
            if 'струк' in input_sheet.cell(row=row, column=5).value.lower():
                del row_indexes[row_indexes.index(row)]
            elif '2к' in input_sheet.cell(row=row, column=5).value:
                del row_indexes[row_indexes.index(row)]
            elif 'кат2' in input_sheet.cell(row=row, column=5).value.split()[0]:
                del row_indexes[row_indexes.index(row)]
    elif main_file == 'data/implants/implants_info.xlsx':
        for i, row in enumerate(row_indexes_copy):
            if input_sheet.cell(row=row, column=3).value != 'Osstem Implant':
                del row_indexes[row_indexes.index(row)]
    elif main_file == 'data/abutments/abutments_info.xlsx':
        for i, row in enumerate(row_indexes_copy):
            if input_sheet.cell(row=row, column=1).value == 'Абатмент выжигаемый':
                del row_indexes[row_indexes.index(row)]

    # Создание словаря с неотгруженными номенклатурами из unshipped_file
    if unshipped_file != 'data/implants/implants_stock.xlsx':
        unshipped_dict = {}
        unshipped_book = openpyxl.load_workbook(unshipped_file)
        unshipped_sheet = unshipped_book.active

        if unshipped_sheet.max_column == 8:
            for row in range(7, unshipped_sheet.max_row + 1):
                unshipped_dict[str(unshipped_sheet.cell(row=row, column=1).value)] = \
                    int(unshipped_sheet.cell(row=row, column=8).value)
            print(f'{unshipped_file} is loaded')
        else:
            print(f'{unshipped_file} have error: max_column != 8:')
    else:
        unshipped_dict = {}

    # Создание словаря для имплантов
    if unshipped_file == 'data/implants/implants_stock.xlsx':
        stock_dict = {}
        stock_book = openpyxl.load_workbook(unshipped_file)
        stock_sheet = stock_book.active

        for row in range(1, stock_sheet.max_row + 1):
            stock_dict[str(stock_sheet.cell(row=row, column=1).value)] = \
                int(stock_sheet.cell(row=row, column=2).value)
        print(f'{unshipped_file} is loaded')

    # Создание словаря с датами пр-ва

    print(f'{main_file} is loaded')
    return unshipped_dict
