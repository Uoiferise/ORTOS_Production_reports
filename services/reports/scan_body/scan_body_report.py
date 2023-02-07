from services.basic_report.basic_report import BasicReport
from services.basic_report.basic_report_sheet import BasicReportSheet
from settings import REPORTS_NAME_DICT


class ReportSB(BasicReport):
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
        BasicReportSheet(wb=self._workbook, name='Скан_боди', data=self._data)

        self._workbook.save(filename=REPORTS_NAME_DICT['scan_body']['report_name'])