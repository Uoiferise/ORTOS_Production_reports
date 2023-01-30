from services.abstractions.abstract_nomenclature import AbstractNomenclature


class Nomenclature(AbstractNomenclature):

    def __init__(self, name: str, id_row: int, info: dict):
        self.name = name
        self.id_row = id_row
        self._info = info

    def get_info(self) -> dict:
        return self._info
