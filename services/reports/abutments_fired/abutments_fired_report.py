from services.basic_report.basic_report import BasicReport
from services.basic_report.basic_report_sheet import BasicReportSheet
from settings import REPORTS_NAME_DICT


class ReportAbutmentsFired(BasicReport):
    __slots__ = ()

    def create_report(self):
        sheets_dict = {
            'Абатменты_выжигаемые': (BasicReportSheet, self._data),
        }

        for name, value in sheets_dict.items():
            current_sheet = value[0](wb=self._workbook, name=name, data=value[1])
            current_sheet.create_sheet()

        self._workbook.save(filename=REPORTS_NAME_DICT[self.report_name]['report_name'])
