from abc import ABC, abstractmethod


class AbstractNomenclature(ABC):

    @abstractmethod
    def __init__(self):
        pass

    @abstractmethod
    def get_info(self):
        pass
