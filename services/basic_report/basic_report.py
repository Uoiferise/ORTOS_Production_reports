from services.abstractions.abstract_report import AbstractReport
from services.basic_report.basic_report_sheet import BasicReportSheet
import openpyxl
from services.basic_report.basic_settings import STYLES_DICT


class BasicReport(AbstractReport):
    """Description will be later ... maybe"""

    __slots__ = ('_data', '_sheets', '_workbook')

    def validation_data(self, data: dict) -> dict:
        # Selection of non-archival items in data
        data_copy = data.copy()
        for key, value in data_copy.items():
            if value.get_info()[6] == 'Да' or value.get_info()[7] == 'Да':
                del data[key]
        return data

    def __init__(self, data: dict, sheets: tuple = None):
        self._data = self.validation_data(data)
        self._sheets = sheets

        self._workbook = openpyxl.Workbook()
        self._workbook.remove(self._workbook.active)

        self.create_styles()
        self.create_report()

    def create_styles(self):
        # Creating a header style
        self._workbook.add_named_style(STYLES_DICT['header'])

        # Creating additional styles
        self._workbook.add_named_style(STYLES_DICT['info'])
        self._workbook.add_named_style(STYLES_DICT['date'])
        self._workbook.add_named_style(STYLES_DICT['white'])

    def create_report(self):
        name = 'basic_report'

        BasicReportSheet(wb=self._workbook, name=name, data=self._data)

        self._workbook.save(filename='output_reports/BasicReport.xlsx')
