from abc import ABC, abstractmethod


class AbstractReportSheet(ABC):

    @abstractmethod
    def __init__(self):
        pass

    @abstractmethod
    def create_sheet_header(self):
        pass

    @abstractmethod
    def fill_small_stock(self):
        pass

    @abstractmethod
    def create_sheet_resul(self):
        pass

    @abstractmethod
    def separation_nomenclatures(self):
        pass

    @abstractmethod
    def cell_style(self):
        pass
