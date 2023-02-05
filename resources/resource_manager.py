from resources.excel_handler.read_data import read_data
from settings import REPORTS_NAME_DICT, OTK_PATH


class ResourceManager:

    __instance = None
    __reports_name_dict = REPORTS_NAME_DICT
    __otk_path = OTK_PATH

    def __new__(cls, *args, **kwargs):
        if cls.__instance is None:
            cls.__instance = super().__new__(cls)
        return cls.__instance

    def __init__(self):
        pass

    @classmethod
    def get_data(cls, report_name: str, resource: str = 'xlsx') -> dict:
        if report_name not in cls.__reports_name_dict.keys():
            raise ValueError('Invalid reports name')
        if resource == 'xlsx':
            return read_data(main_file=cls.__reports_name_dict[report_name]['main_file'],
                             unshipped_file=cls.__reports_name_dict[report_name]['unshipped_file'],
                             otk_file=cls.__otk_path)
        else:
            raise ValueError('Invalid data source format')
