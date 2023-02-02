from services.abstractions.abstract_report_sheet import AbstractReportSheet
from openpyxl import Workbook
from basic_settings import HEADERS_DICT
from settings import DATE_START, DATE_STOP
from basic_functions.create_sheet_header import create_sheet_header
from basic_functions.fill_small_stock import fill_small_stock
from basic_functions.create_sheet_result import create_sheet_result
from basic_functions.separation_nomenclatures import separation_nomenclatures
from openpyxl.styles import Alignment


class BasicReportSheet(AbstractReportSheet):
    """Description will be later ... maybe"""

    __slots__ = ('_wb', '_data')

    _HEADERS_DICT = HEADERS_DICT
    _DATE_START = DATE_START
    _DATE_STOP = DATE_STOP

    def __init__(self, wb: Workbook, name: str, data):
        self._wb = wb
        self._data = data
        self._sheet = self._wb.create_sheet(title=name, index=0)
        self.create_sheet_header()
        self._start_row = self._sheet.max_row
        self.transport_date()
        self.fill_small_stock()
        self.separation_nomenclatures()
        self.create_sheet_resul()

    def create_sheet_header(self) -> None:
        create_sheet_header(sheet=self._sheet,
                            date_start=self._DATE_START,
                            date_stop=self._DATE_STOP,
                            header_dict=self._HEADERS_DICT)

    def fill_small_stock(self) -> None:
        fill_small_stock(sheet=self._sheet,
                         start_row=self._start_row)

    def create_sheet_resul(self) -> None:
        create_sheet_result(sheet=self._sheet,
                            start_row=self._start_row,
                            end_row=self._sheet.max_row)

    def separation_nomenclatures(self) -> None:
        separation_nomenclatures(sheet=self._sheet,
                                 start_row=self._start_row)

    def cell_style(self) -> None:
        pass

    def transport_date(self) -> None:
        for row in range(self._start_row, self._start_row + len(self._data)):
            nomenclature = self._data[row-3]
            for col in range(1, 20):
                cell = self._sheet.cell(row=row, column=col)
                info = nomenclature.get_info()[col]
                if 9 <= col <= 18 and info is None:
                    info = 0
                cell.value = info
                cell.style = 'info'
                if col >= 9:
                    cell.alignment = Alignment(horizontal='center', vertical='center')
