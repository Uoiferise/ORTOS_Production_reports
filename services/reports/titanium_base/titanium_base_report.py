import openpyxl
from services.basic_report.basic_report import BasicReport
from services.basic_report.basic_report_sheet import BasicReportSheet
from settings import REPORTS_NAME_DICT
from openpyxl.workbook import Workbook
from openpyxl.styles import Alignment, Font


class ReportTitaniumBase(BasicReport):
    __slots__ = ()

    __TB_ACTUAL_BOOK_PATH = 'services/reports/titanium_base/titanium_base_actual.xlsx'

    def validation_data(self, data: dict) -> dict:
        # Aggregation of information of some archival nomenclatures
        tb_actual_book = openpyxl.load_workbook(self.__TB_ACTUAL_BOOK_PATH)
        tb_actual_sheet = tb_actual_book.active
        for r in range(2, tb_actual_sheet.max_row + 1):
            actual_nom_name = tb_actual_sheet.cell(row=r, column=1).value
            archival_nom_name = tb_actual_sheet.cell(row=r, column=2).value
            if actual_nom_name in data.keys() and archival_nom_name in data.keys():
                add_info = data[archival_nom_name].get_info()
                data[actual_nom_name].aggregate_info(add_info)

        # Selection of non-archival items in data and delete some nomenclatures / speed = O(n**2)
        data_copy = data.copy()
        for key, value in data_copy.items():
            if value.get_info()[6] == 'Да' or \
                    value.get_info()[7] == 'Да' or \
                    'струк' in value.get_info()[5].lower() or \
                    '2к' in value.get_info()[5] or \
                    'кат2' in value.get_info()[5].split()[0]:
                del data[key]

        return data

    def create_report(self) -> None:
        # divide the information into 7 sheets
        data_1 = dict()
        data_2 = dict()
        data_3 = dict()
        data_4 = dict()
        data_5 = dict()
        data_6 = dict()
        data_7 = dict()

        for key, value in self._data.items():
            if 'Patch' == value.get_info()[2]:
                data_1[key] = value
            elif 'Flat' == value.get_info()[2]:
                data_2[key] = value
            elif 'Half' == value.get_info()[2]:
                data_3[key] = value
            elif 'Bell GEO' == value.get_info()[2]:
                data_4[key] = value
            elif 'Step GEO' == value.get_info()[2]:
                data_5[key] = value
            elif 'Step ARUM' == value.get_info()[2]:
                data_6[key] = value
            else:
                data_7[key] = value

        # del_keys = []
        # for key, value in data_1.items():
        #     if 'P' in value.get_info()[8] or 'Н' in value.get_info()[8]:
        #         for k, v in data_1.items():
        #             if value.vendor_code == v.vendor_code and \
        #                     'P' not in v.get_info()[8] and \
        #                     'Н' not in v.get_info()[8] and \
        #                     value.get_info()[8][:-2] == v.get_info()[8][:-2]:
        #                 v.aggregate_info(value.get_info())
        #         del_keys.append(key)

        for key, value in data_1.items():
            if 'P' in value.get_info()[8] or 'Н' in value.get_info()[8]:
                for k, v in data_1.items():
                    if value.vendor_code == v.vendor_code and value.get_info()[8][:-2] == v.get_info()[8][:-2]:
                        print(v.vendor_code)
                        if len(v.get_info()[8]) == 8:
                            print(f'{key} == {k}')

        # for key in del_keys:
        #     del data_1[key]

        sheets_dict = {
            'Остальное': (BasicReportSheet, data_7),
            'Arum': (BasicReportSheet, data_6),
            'GEO Step': (BasicReportSheet, data_5),
            'GEO Bell': (TBReportSheet, data_4),
            'Half (ИМ Абатменты.ру)': (BasicReportSheet, data_3),
            'Flat с насечками (ИМ Ортос)': (TBReportSheet, data_2),
            'Patch': (TBReportSheet, data_1),
        }

        for name, value in sheets_dict.items():
            current_sheet = value[0](wb=self._workbook, name=name, data=value[1])
            current_sheet.create_sheet()

        self._workbook.save(filename=REPORTS_NAME_DICT['titanium_base']['report_name'])


class TBReportSheet(BasicReportSheet):
    __slots__ = ('_data_bridge', '_data_single')

    @staticmethod
    def divide_data(data: dict) -> tuple:
        data_bridge, data_single = dict(), dict()
        for key, value in data.items():
            if 'ТО bridge' == value.get_info()[1]:
                data_bridge[key] = value
            else:
                data_single[key] = value
        return data_bridge, data_single

    def __init__(self, wb: Workbook, name: str, data):
        super().__init__(wb, name, data)
        self._data_bridge, self._data_single = self.divide_data(data=self._data)

    def add_title(self, title: str, row: int) -> None:
        cell = self._sheet.cell(row=row, column=5)
        cell.value = title
        for col in range(1, self._sheet.max_column + 1):
            cell = self._sheet.cell(row=row, column=col)
            self.cell_style(cell)
            if col == 5:
                cell.font = Font(name='Arial', bold=True, size=10)
                cell.alignment = Alignment(horizontal='center', vertical='center')

    def create_sheet(self) -> None:
        self.create_sheet_header()

        self.add_title(title='Мостовидные', row=self._sheet.max_row)
        start_row = self._sheet.max_row + 1
        self.transport_date(self._data_bridge, start_row=start_row)
        self.separation_nomenclatures(start_row=start_row)
        for col in range(1, self._sheet.max_column + 1):
            cell = self._sheet.cell(row=self._sheet.max_row, column=col)
            self.cell_style(cell)

        self.add_title(title='Одиночные', row=self._sheet.max_row+1)
        start_row = self._sheet.max_row + 1
        self.transport_date(self._data_single, start_row=start_row)
        self.separation_nomenclatures(start_row=start_row)

        self.fill_small_stock()
        self.create_sheet_resul()
