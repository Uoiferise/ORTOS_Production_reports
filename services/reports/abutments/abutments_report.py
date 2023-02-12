from services.basic_report.basic_report import BasicReport
from services.basic_report.basic_report_sheet import BasicReportSheet
from settings import REPORTS_NAME_DICT


class ReportAbutments(BasicReport):
    __slots__ = ()

    def validation_data(self, data: dict) -> dict:
        # Selection of non-archival items in data and delete not stainless steel
        data_copy = data.copy()
        for key, value in data_copy.items():
            if value.get_info()[6] == 'Да' or \
                    value.get_info()[7] == 'Да' or \
                    value.get_info()[1] == 'Абатмент выжигаемый':
                del data[key]
        return data

    def create_report(self):
        # divide the information into 2 sheets
        data_1, data_2 = self._data.copy(), self._data.copy()
        for key, value in self._data.items():
            if value.get_info()[1] == 'Абатмент приливаемый':
                del data_1[key]
            else:
                del data_2[key]

        sheets_dict = {
            'Приливаемый': (BasicReportSheet, data_2),
            'Прямой, временный': (BasicReportSheet, data_1),
        }

        for name, value in sheets_dict.items():
            current_sheet = value[0](wb=self._workbook, name=name, data=value[1])
            current_sheet.create_sheet()

        self._workbook.save(filename=REPORTS_NAME_DICT[self.report_name]['report_name'])
