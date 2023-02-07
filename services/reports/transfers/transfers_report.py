from services.basic_report.basic_report import BasicReport
from services.basic_report.basic_report_sheet import BasicReportSheet
from settings import REPORTS_NAME_DICT


class ReportTransfers(BasicReport):
    __slots__ = ()

    def create_report(self):
        BasicReportSheet(wb=self._workbook, name='Трансферы', data=self._data)

        self._workbook.save(filename=REPORTS_NAME_DICT['transfers']['report_name'])
