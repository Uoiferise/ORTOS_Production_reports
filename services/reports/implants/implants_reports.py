from services.basic_report.basic_report import BasicReport
from services.basic_report.basic_report_sheet import BasicReportSheet
from settings import REPORTS_NAME_DICT
import openpyxl
from copy import copy


class ReportImplants(BasicReport):
    __slots__ = ()

    def create_report(self):
        data_1 = dict()

        for key, value in self._data.items():
            if 'Osstem Implant' == value.get_info()[3]:
                data_1[key] = value

        sheets_dict = {
            'Implantium': (ReportImplantsSheet, self._data),
            'Nobel Active': (ReportImplantsSheet, self._data),
            'Osstem': (ReportImplantsSheet, data_1),
        }

        for name, value in sheets_dict.items():
            current_sheet = value[0](wb=self._workbook, name=name, data=value[1])
            current_sheet.create_sheet()

        self._workbook.save(filename=REPORTS_NAME_DICT['implants']['report_name'])


class ReportImplantsSheet(BasicReportSheet):
    __slots__ = ()

    __COPY_SHEET_PATH_DICT = {
        'Implantium': 'services/reports/implants/Implantium.xlsx',
        'Nobel Active': 'services/reports/implants/Nobel Active.xlsx',
        'Osstem': 'services/reports/implants/Osstem.xlsx',
    }

    def copy_sheet(self, wb_path: str) -> None:
        wb_from = openpyxl.load_workbook(wb_path)
        ws_from = wb_from.active

        for i in range(self._sheet.max_row, ws_from.max_row + 1):
            for j in range(1, ws_from.max_column + 1):
                # reading cell value from source excel file
                cell_from = ws_from.cell(row=i, column=j)

                # writing the read value to destination excel file
                self._sheet.cell(row=i, column=j).value = cell_from.value
                new_cell = self._sheet.cell(row=i, column=j)

                if cell_from.has_style:
                    new_cell.font = copy(cell_from.font)
                    new_cell.border = copy(cell_from.border)
                    new_cell.fill = copy(cell_from.fill)
                    new_cell.number_format = copy(cell_from.number_format)
                    new_cell.protection = copy(cell_from.protection)
                    new_cell.alignment = copy(cell_from.alignment)

            self._sheet.row_dimensions[i].height = 15

    def transport_date(self, data: dict, start_row: int = None) -> None:
        for row in range(7, self._sheet.max_row + 1):
            nomenclature_name = self._sheet.cell(row=row, column=5).value
            if data.get(nomenclature_name, None) is not None:
                for col in range(9, 22):
                    cell = self._sheet.cell(row=row, column=col)
                    info = data[nomenclature_name].get_info()[col]
                    if 9 <= col <= 18 and info is None:
                        info = 0
                    cell.value = info
            elif nomenclature_name is not None and nomenclature_name != 'Итого':
                for col in range(9, 22):
                    cell = self._sheet.cell(row=row, column=col)
                    if 9 <= col <= 18:
                        info = 0
                    else:
                        info = None
                    cell.value = info
        self._sheet.column_dimensions.group('A', 'D', hidden=True)

    def create_sheet(self) -> None:
        self.create_sheet_header()
        self.copy_sheet(wb_path=self.__COPY_SHEET_PATH_DICT[self.name])
        self.transport_date(data=self._data)
