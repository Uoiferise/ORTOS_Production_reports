from resources.resource_manager import ResourceManager
from services import ReportAbutments, ReportAbutmentsFired, ReportAnalog, ReportBlanks, ReportFormers, \
    ReportImplants, ReportScanBody, ReportScrews, ReportSleeve, ReportTitaniumBase, ReportTransfers


REPORTS_DICT = {
    'abutments': ReportAbutments,
    'abutments_fired': ReportAbutmentsFired,
    'analog': ReportAnalog,
    'blanks': ReportBlanks,
    'formers': ReportFormers,
    'implants': ReportImplants,
    'implants_new': ReportImplants,
    'scan_body': ReportScanBody,
    'screws': ReportScrews,
    'sleeve': ReportSleeve,
    'titanium_base': ReportTitaniumBase,
    'transfers': ReportTransfers,
}


def main():
    for report_name, report_class in REPORTS_DICT.items():
        if report_class is None:
            continue
        data = ResourceManager.get_data(report_name=report_name)
        report_class(data=data, report_name=report_name)
        print(f'{report_name} is done!')
        print()


if __name__ == '__main__':
    main()
