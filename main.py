from resources.resource_manager import ResourceManager
from services.reports.abutments.abutments_report import ReportAbutments


def main():
    data = ResourceManager.get_data(report_name='abutments')
    ReportAbutments(data=data)


if __name__ == '__main__':
    main()
