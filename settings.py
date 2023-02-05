# a report sheet header information
DATE_START = '01.01.2023'
DATE_STOP = '01.02.2023'

# a report name information
DATE_START_RN = DATE_START[0:5]
DATE_STOP_RN = DATE_STOP[0:5]

OUTPUT_PATH = 'output_reports'

# logic settings
REPORTS_NAME_DICT = {
    'abutments': {
        'main_file': 'input_data/abutments/abutments_info.xlsx',
        'unshipped_file': 'input_data/abutments/abutments_unsh.xlsx',
        'report_name': f'{OUTPUT_PATH}/Абатменты_{DATE_START_RN}-{DATE_STOP_RN}.xlsx'
    },
    'abutments_fired': {
        'main_file': 'input_data/abutments_fired/abutments_fired_info.xlsx',
        'unshipped_file': 'input_data/abutments_fired/abutments_fired_unsh.xlsx',
        'report_name': 'Абатменты_выжигаемые'
    },
    'analog': {
        'main_file': 'input_data/analog/analog_info.xlsx',
        'unshipped_file': 'input_data/analog/analog_unsh.xlsx',
        'report_name': 'Аналоги'
    },
    'blanks': {
        'main_file': 'input_data/blanks/blanks_info.xlsx',
        'unshipped_file': 'input_data/blanks/blanks_unsh.xlsx',
        'report_name': 'Заготовки'
    },
    'formers': {
        'main_file': 'input_data/formers/formers_info.xlsx',
        'unshipped_file': 'input_data/formers/formers_unsh.xlsx',
        'report_name': 'Формирователи'
    },
    'implants': {
        'main_file': 'input_data/implants/implants_info.xlsx',
        'unshipped_file': 'input_data/implants/implants_unsh.xlsx',
        'report_name': 'Импланты'
    },
    'scan_body': {
        'main_file': 'input_data/scan_body/scan_body_info.xlsx',
        'unshipped_file': 'input_data/scan_body/scan_body_unsh.xlsx',
        'report_name': 'Скан_боди'
    },
    'screws': {
        'main_file': 'input_data/screws/screws_info.xlsx',
        'unshipped_file': 'input_data/screws/screws_unsh.xlsx',
        'report_name': 'Винты'
    },
    'sleeve': {
        'main_file': 'input_data/sleeve/sleeve_info.xlsx',
        'unshipped_file': 'input_data/sleeve/sleeve_unsh.xlsx',
        'report_name': 'Втулка'
    },
    'titanium_base': {
        'main_file': 'input_data/titanium_base/titanium_base_info.xlsx',
        'unshipped_file': 'input_data/titanium_base/titanium_base_unsh.xlsx',
        'report_name': 'Титановые_основы'
    },
    'transfers': {
        'main_file': 'input_data/transfers/transfers_info.xlsx',
        'unshipped_file': 'input_data/transfers/transfers_unsh.xlsx',
        'report_name': 'Трансферы'
    },
}
