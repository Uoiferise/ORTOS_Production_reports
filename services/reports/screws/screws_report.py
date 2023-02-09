from services.basic_report.basic_report import BasicReport
from services.basic_report.basic_report_sheet import BasicReportSheet
from settings import REPORTS_NAME_DICT


class ReportScrews(BasicReport):
    __slots__ = ()

    def validation_data(self, data: dict) -> dict:
        # Selection of non-archival items in data and delete some nomenclatures
        data_copy = data.copy()
        for key, value in data_copy.items():
            if value.get_info()[6] == 'Да' or \
                    value.get_info()[7] == 'Да' or \
                    'блистер' in value.get_info()[5].lower() or \
                    'упак' in value.get_info()[5].lower() or \
                    'проб' in value.get_info()[5].lower():
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
            if 'трансфер' in value.get_info()[1]:
                data_6[key] = value
            elif 'SIRONA' == value.get_info()[2]:
                data_5[key] = value
            elif 'Аналог NT-Traiding' == value.get_info()[2]:
                data_4[key] = value
            elif 'ZIRKONZAHN' == value.get_info()[2]:
                data_3[key] = value
            elif ('3D' in value.get_info()[1]) or \
                 ('Втулка сварного винта' == value.get_info()[1]) or \
                 ('Пин' == value.get_info()[1]) or \
                 ('угл' in value.get_info()[1]) or \
                 ('Винт LM (собств. разр.)' in value.get_info()[5]):
                data_2[key] = value
            elif 'Винт LM (копия оригинала)' in value.get_info()[5]:
                data_1[key] = value
            else:
                data_7[key] = value

        sheets_dict = {
            'Лабораторные винты LM': (BasicReportSheet, data_7),
            'Для трансферов': (BasicReportSheet, data_6),
            'SIRONA': (BasicReportSheet, data_5),
            'NT-trading': (BasicReportSheet, data_4),
            'Zirkonzahn': (BasicReportSheet, data_3),
            'Собств. разработка': (BasicReportSheet, data_2),
            'Винты LM': (BasicReportSheet, data_1),
        }

        for name, value in sheets_dict.items():
            current_sheet = value[0](wb=self._workbook, name=name, data=value[1])
            current_sheet.create_sheet()

        self._workbook.save(filename=REPORTS_NAME_DICT['screws']['report_name'])
