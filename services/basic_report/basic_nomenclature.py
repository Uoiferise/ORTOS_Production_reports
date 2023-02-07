from services.abstractions.abstract_nomenclature import AbstractNomenclature


class Nomenclature(AbstractNomenclature):

    def __init__(self, name: str, id_row: int, info: dict):
        self.name = name
        self.id_row = id_row
        self._info = info

    def get_info(self) -> dict:
        return self._info

    def set_info(self, info: dict) -> None:
        self._info = info

    def aggregate_info(self, info: dict = None) -> None:
        if info is not None:
            for key in self._info.keys():
                if isinstance(self._info[key], int):
                    self._info[key] = self._info[key] + (0, info[key])[info[key] is not None]
