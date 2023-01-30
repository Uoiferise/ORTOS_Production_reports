from services.abstractions.abstract_report import AbstractReport
from services.basic_report.basic_report_sheet import BasicReportSheet
import openpyxl
from openpyxl.styles import NamedStyle, Font, Border, Side, PatternFill, Alignment


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
        ns_header = NamedStyle(name='header')
        ns_header.font = Font(name='Arial', bold=True, size=10)
        ns_header.fill = PatternFill("solid", fgColor="D6E5CB")
        thin = Side(border_style="thin", color="000000")
        ns_header.border = Border(top=thin, left=thin, right=thin, bottom=thin)
        ns_header.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        self._workbook.add_named_style(ns_header)

        # Creating additional styles
        ns_info = NamedStyle(name='info')
        thin = Side(border_style="thin", color="000000")
        ns_info.border = Border(top=thin, left=thin, right=thin, bottom=thin)
        ns_info.font = Font(name='Arial', bold=False, size=8)
        self._workbook.add_named_style(ns_info)

        ns_date = NamedStyle(name='date')
        thin = Side(border_style="thin", color="000000")
        ns_date.border = Border(top=thin, left=thin, right=thin, bottom=thin)
        ns_date.font = Font(name='Arial', bold=True, size=10)
        ns_date.alignment = Alignment(horizontal='center', vertical='center')
        self._workbook.add_named_style(ns_date)

        ns_white = NamedStyle(name='white')
        ns_white.font = Font(name='Arial', bold=False, size=8)
        ns_white.alignment = Alignment(horizontal='center', vertical='center')
        ns_white.fill = PatternFill("solid", fgColor="FFFFFF")
        self._workbook.add_named_style(ns_white)

    def create_report(self):
        self.create_styles()
        for name in self._sheets:
            BasicReportSheet(self._workbook, name=name, data=self._data[name])
