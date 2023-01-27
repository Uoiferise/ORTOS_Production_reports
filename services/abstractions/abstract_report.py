from abc import ABC, abstractmethod


class AbstractReport(ABC):

    @abstractmethod
    def __init__(self):
        pass

    @abstractmethod
    def create_styles(self):
        pass

    @abstractmethod
    def create_report(self):
        pass
