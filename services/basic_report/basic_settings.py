from basic_styles.basic_header import basic_header
from basic_styles.basic_info import basic_info
from basic_styles.basic_date import basic_date
from basic_styles.basic_white import basic_white


HEADERS_DICT = {
            1: 'Тип',
            2: 'Линейка',
            3: 'Система',
            4: 'Разм',
            5: 'Номенклатура',
            6: 'Арх ном',
            7: 'Арх кат',
            8: 'Карт кат',
            9: 'Остаток',
            10: 'Расход общий',
            19: 'ПЛАН',
            20: 'Произведено / неоприходовано',
            21: 'Непроизведено / в плане',
            22: 'Неотгружено по опт. заявкам'
        }

STYLES_DICT = {
    'header': basic_header,
    'info': basic_info,
    'date': basic_date,
    'white': basic_white,
}
