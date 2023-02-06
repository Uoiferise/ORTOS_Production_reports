from resources.resource_manager import ResourceManager
from services.reports.abutments.abutments_report import ReportAbutments
from services.reports.analog.analog_report import ReportAnalog
from services.reports.screws.screws_report import ReportScrews


def main():
    data = ResourceManager.get_data(report_name='abutments')
    ReportAbutments(data=data)

    # data = ResourceManager.get_data(report_name='analog')
    # ReportAnalog(data=data)

    # data = ResourceManager.get_data(report_name='screws')
    # ReportScrews(data=data)


if __name__ == '__main__':
    main()
