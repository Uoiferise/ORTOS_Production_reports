import openpyxl
from services.basic_report.basic_report import BasicReport
from services.basic_report.basic_report_sheet import BasicReportSheet
from settings import REPORTS_NAME_DICT


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

        # Selection of non-archival items in data and delete some nomenclatures
        data_copy = data.copy()
        for key, value in data_copy.items():
            if value.get_info()[6] == 'Да' or \
                    value.get_info()[7] == 'Да' or \
                    'струк' in value.get_info()[5].lower() or \
                    '2к' in value.get_info()[5] or \
                    'кат2' in value.get_info()[5].split()[0]:
                del data[key]

        return data

    def create_report(self):
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

        sheets_dict = {
            'Остальное': (BasicReportSheet, data_7),
            'Arum': (BasicReportSheet, data_6),
            'GEO Step': (BasicReportSheet, data_5),
            'GEO Bell': (BasicReportSheet, data_4),
            'Half (ИМ Абатменты.ру)': (BasicReportSheet, data_3),
            'Flat с насечками (ИМ Ортос)': (BasicReportSheet, data_2),
            'Patch': (BasicReportSheet, data_1),
        }

        for name, value in sheets_dict.items():
            current_sheet = value[0](wb=self._workbook, name=name, data=value[1])
            current_sheet.create_sheet()

        self._workbook.save(filename=REPORTS_NAME_DICT['titanium_base']['report_name'])
