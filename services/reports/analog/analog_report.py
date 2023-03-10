from services.basic_report.basic_report import BasicReport
from services.basic_report.basic_report_sheet import BasicReportSheet
from settings import REPORTS_NAME_DICT


class ReportAnalog(BasicReport):
    __slots__ = ()

    def validation_data(self, data: dict) -> dict:
        # Selection of non-archival items in data and delete not stainless steel
        data_copy = data.copy()
        for key, value in data_copy.items():
            if value.get_info()[6] == 'Да' or \
                    value.get_info()[7] == 'Да' or \
                    'нерж' not in value.get_info()[5].lower():
                del data[key]
        return data

    def create_report(self):
        sheets_dict = {
            'Аналоги': (ReportAnalogSheet, self._data),
        }

        for name, value in sheets_dict.items():
            current_sheet = value[0](wb=self._workbook, name=name, data=value[1])
            current_sheet.create_sheet()

        self._workbook.save(filename=REPORTS_NAME_DICT[self.report_name]['report_name'])


class ReportAnalogSheet(BasicReportSheet):
    __slots__ = ()

    __MATCH_FILE_PATH = 'services/reports/analog/analog_matching.xlsx'

    def create_sheet(self) -> None:
        self.create_sheet_header()
        self.transport_date(data=self._data)
        self.fill_small_stock()
        self.separation_nomenclatures(exception=True)
        self.create_sheet_resul()
        self.match_nomenclatures(match_file=self.__MATCH_FILE_PATH)
        self.create_sheet_resul()
