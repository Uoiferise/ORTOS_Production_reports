from services.basic_report.basic_report import BasicReport
from services.basic_report.basic_report_sheet import BasicReportSheet
from settings import REPORTS_NAME_DICT


class ReportScanBody(BasicReport):
    __slots__ = ()

    def validation_data(self, data: dict) -> dict:
        # Selection of non-archival items in data and delete not stainless steel
        data_copy = data.copy()
        for key, value in data_copy.items():
            if value.get_info()[6] == 'Да' or \
                    value.get_info()[7] == 'Да' or \
                    ('нерж' in value.get_info()[5].lower() and 'б' in value.get_info()[5].split()[0]) or \
                    ('латунь' in value.get_info()[5].lower() and 'б' in value.get_info()[5].split()[0]):
                del data[key]
        return data

    def create_report(self):
        sheets_dict = {
            'Скан_боди': (BasicReportSheet, self._data),
        }

        for name, value in sheets_dict.items():
            current_sheet = value[0](wb=self._workbook, name=name, data=value[1])
            current_sheet.create_sheet()

        self._workbook.save(filename=REPORTS_NAME_DICT['scan_body']['report_name'])