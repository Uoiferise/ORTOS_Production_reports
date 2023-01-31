from resources.resource_manager import ResourceManager
from services.reports.abutments.abutments_report import ReportAbutments


def main():
    data = ResourceManager.get_data(report_name='input_data/abutments/abutments_info.xlsx')

    # test_name = ''
    # for key, value in data[test_name].get_info().items():
    #     print(f'{key}: {value}')

    report = ReportAbutments(data=data, sheets=tuple())


if __name__ == '__main__':
    main()
