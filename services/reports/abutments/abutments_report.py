from services.basic_report.basic_report import BasicReport
from services.basic_report.basic_report_sheet import BasicReportSheet
from settings import REPORTS_NAME_DICT


class ReportAbutments(BasicReport):
    __slots__ = ()

    def create_report(self):
        # divide the information into 2 sheets
        data_1, data_2 = self._data.copy(), self._data.copy()
        for key, value in self._data.items():
            if value.get_info()[1] == 'Абатмент приливаемый':
                del data_1[key]
            else:
                del data_2[key]

        BasicReportSheet(wb=self._workbook, name='Приливаемый', data=data_2)
        BasicReportSheet(wb=self._workbook, name='Прямой, временный', data=data_1)

        self._workbook.save(filename=REPORTS_NAME_DICT['abutments']['report_name'])
