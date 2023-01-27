from resources.excel_handler.read_data import read_input_files


class ResourceManager:

    __instance = None

    def __new__(cls, *args, **kwargs):
        if cls.__instance is None:
            cls.__instance = super().__new__(cls)
        return cls.__instance

    def __init__(self):
        pass

    @staticmethod
    def get_data(report_name: str, resource='xlsx') -> list:
        if resource == 'xlsx':
            return read_input_files(main_file=None, unshipped_file=None, options=None)
        else:
            raise ValueError('Неверный формат источника данных')
