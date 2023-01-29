from resources.excel_handler.read_data import read_data


class ResourceManager:

    __instance = None

    def __new__(cls, *args, **kwargs):
        if cls.__instance is None:
            cls.__instance = super().__new__(cls)
        return cls.__instance

    def __init__(self):
        pass

    @staticmethod
    def get_data(report_name: str, resource: str = 'xlsx') -> dict:
        if resource == 'xlsx':
            return read_data(main_file=report_name, unshipped_file=report_name)
        else:
            raise ValueError('Invalid data source format')
