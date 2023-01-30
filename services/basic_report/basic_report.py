from services.abstractions.abstract_report import AbstractReport
from services.basic_report.basic_report_sheet import BasicReportSheet
import openpyxl
from services.basic_report.basic_settings import STYLES_DICT


class BasicReport(AbstractReport):
    """Description will be later ... maybe"""

    __slots__ = ('_name', '_data', '_sheets', '_workbook')

    def __init__(self, name: str, data: dict, sheets: tuple):
        self._name = name
        self._data = data
        self._sheets = sheets

        self._workbook = openpyxl.Workbook()
        self._workbook.remove(self._workbook.active)

    def create_styles(self):
        # Creating a header style
        self._workbook.add_named_style(STYLES_DICT['header'])

        # Creating additional styles

        self._workbook.add_named_style(STYLES_DICT['info'])
        self._workbook.add_named_style(STYLES_DICT['date'])
        self._workbook.add_named_style(STYLES_DICT['white'])

    def create_report(self):
        self.create_styles()
        for name in self._sheets:
            BasicReportSheet(self._workbook, name=name, data=self._data[name])
