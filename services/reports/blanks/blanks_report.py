from services.basic_report.basic_report import BasicReport
from services.basic_report.basic_report_sheet import BasicReportSheet
from settings import REPORTS_NAME_DICT


class ReportBlanks(BasicReport):
    __slots__ = ()

    def create_report(self):
        # divide the information into 6 sheets
        data_1 = dict()
        data_2 = dict()
        data_3 = dict()
        data_4 = dict()
        data_5 = dict()
        data_6 = dict()

        sheets_dict = {
            'Остальное': (BasicReportSheet, self._data),
            'DM_Medentika': (BasicReportSheet, self._data),
            'Zirkonzahn': (BasicReportSheet, self._data),
            'ARUM_Ti': (BasicReportSheet, self._data),
            'ARUM_CoCr': (BasicReportSheet, self._data),
            'LM1': (BasicReportSheet, self._data),
        }

        for name, value in sheets_dict.items():
            current_sheet = value[0](wb=self._workbook, name=name, data=value[1])
            current_sheet.create_sheet()

        self._workbook.save(filename=REPORTS_NAME_DICT['blanks']['report_name'])
