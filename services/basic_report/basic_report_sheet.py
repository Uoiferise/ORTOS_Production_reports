from services import AbstractReportSheet
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from basic_settings import HEADERS_DICT
from settings import DATE_START, DATE_STOP


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

    def create_sheet_header(self):
        self._sheet.cell(row=1, column=5).value = f'Конец периода: {self._DATE_STOP} 23:59:59'
        self._sheet.cell(row=1, column=5).font = Font(name='Arial', bold=False, size=8)

        self._sheet.cell(row=2, column=5).value = f'Начало периода: {self._DATE_START} 00:00:00'
        self._sheet.cell(row=2, column=5).font = Font(name='Arial', bold=False, size=8)

        for key in self._HEADERS_DICT.keys():
            if 9 <= key <= 10:
                self._sheet.merge_cells(start_row=5, start_column=int(key), end_row=6, end_column=int(key))
                self._sheet.cell(row=5, column=int(key)).value = self._HEADERS_DICT[key]
            else:
                self._sheet.merge_cells(start_row=4, start_column=int(key), end_row=6, end_column=int(key))
                self._sheet.cell(row=4, column=int(key)).value = self._HEADERS_DICT[key]

        self._sheet.merge_cells(start_row=4, start_column=9, end_row=4, end_column=10)
        self._sheet.cell(row=4, column=9).value = 'Итого'
        self._sheet.merge_cells(start_row=4, start_column=11, end_row=4, end_column=14)
        self._sheet.cell(row=4, column=11).value = 'ОСНОВНЫЕ СКЛАДЫ'
        self._sheet.merge_cells(start_row=4, start_column=15, end_row=4, end_column=18)
        self._sheet.cell(row=4, column=15).value = 'ПРОЧИЕ СКЛАДЫ'

        for c in [11, 15]:
            self._sheet.merge_cells(start_row=5, start_column=c, end_row=5, end_column=c + 1)
            self._sheet.cell(row=5, column=c).value = 'ОСТ'

        for c in [13, 17]:
            self._sheet.merge_cells(start_row=5, start_column=c, end_row=5, end_column=c + 1)
            self._sheet.cell(row=5, column=c).value = 'РАСХ'

        for c in range(11, 19):
            if c % 2 != 0:
                self._sheet.cell(row=6, column=c).value = 'ИЗД'
            else:
                self._sheet.cell(row=6, column=c).value = 'К/Т'

        for row in range(4, 7):
            for c in range(1, 23):
                self._sheet.cell(row=row, column=c).style = 'header'

        for c in range(1, 23):
            if c <= 4:
                self._sheet.column_dimensions[get_column_letter(c)].width = 9
            elif c == 5:
                self._sheet.column_dimensions[get_column_letter(c)].width = 90
            elif 6 <= c <= 8:
                self._sheet.column_dimensions[get_column_letter(c)].width = 7.5
            elif 9 <= c <= 18:
                self._sheet.column_dimensions[get_column_letter(c)].width = 8.25
            else:
                self._sheet.column_dimensions[get_column_letter(c)].width = 20

        self._sheet.column_dimensions.group('F', 'H', hidden=True)
        self._sheet.freeze_panes = self._sheet.cell(row=7, column=6)

    def fill_small_stock(self, start_row):
        for row in range(start_row, self._sheet.max_row + 1):
            if self._sheet.cell(row=row, column=9).value is None:
                continue
            elif self._sheet.cell(row=row, column=10).value >= self._sheet.cell(row=row, column=9).value:
                self._sheet.cell(row=row, column=5).fill = PatternFill("solid", fgColor="FFCCCC")
                for c in [9, 10]:
                    self._sheet.cell(row=row, column=c).fill = PatternFill("solid", fgColor="FFCCCC")
                    self._sheet.cell(row=row, column=c).alignment = Alignment(horizontal='center', vertical='center')
            elif self._sheet.cell(row=row, column=13).value != 0:
                if (self._sheet.cell(row=row, column=11).value / self._sheet.cell(row=row, column=13).value) < 1:
                    self._sheet.cell(row=row, column=5).fill = PatternFill("solid", fgColor="CCECFF")
                    for c in [11, 13]:
                        self._sheet.cell(row=row, column=c).fill = PatternFill("solid", fgColor="CCECFF")
                        self._sheet.cell(row=row, column=c).alignment = Alignment(horizontal='center', vertical='center')
                elif 1 <= (self._sheet.cell(row=row, column=11).value / self._sheet.cell(row=row, column=13).value) < 2.5:
                    self._sheet.cell(row=row, column=5).fill = PatternFill("solid", fgColor="CCFFCC")
                    for c in [11, 13]:
                        self._sheet.cell(row=row, column=c).fill = PatternFill("solid", fgColor="CCFFCC")
                        self._sheet.cell(row=row, column=c).alignment = Alignment(horizontal='center', vertical='center')

    def create_sheet_resul(self):
        pass

    def separation_nomenclatures(self):
        pass

    def cell_style(self):
        pass
