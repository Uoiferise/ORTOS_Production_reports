from resources.resource_manager import ResourceManager
from services import ReportAbutments, ReportAnalog, ReportSB, ReportScrews, ReportSleeve, ReportTransfers


def main():
    data = ResourceManager.get_data(report_name='abutments')
    ReportAbutments(data=data)

    data = ResourceManager.get_data(report_name='analog')
    ReportAnalog(data=data)

    data = ResourceManager.get_data(report_name='screws')
    ReportScrews(data=data)

    data = ResourceManager.get_data(report_name='sleeve')
    ReportSleeve(data=data)

    data = ResourceManager.get_data(report_name='transfers')
    ReportTransfers(data=data)

    data = ResourceManager.get_data(report_name='scan_body')
    ReportSB(data=data)


if __name__ == '__main__':
    main()
