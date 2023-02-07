from services.basic_report.basic_report import BasicReport
from services.basic_report.basic_report_sheet import BasicReportSheet
from settings import REPORTS_NAME_DICT


class ReportSleeve(BasicReport):
    __slots__ = ()

    def create_report(self):
        BasicReportSheet(wb=self._workbook, name='Втулка', data=self._data)

        self._workbook.save(filename=REPORTS_NAME_DICT['sleeve']['report_name'])
