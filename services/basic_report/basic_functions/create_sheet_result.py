def create_sheet_result(sheet, start_row: int, end_row: int) -> None:
    rows_dict = {
        9: 'I',
        10: 'J',
        11: 'K',
        12: 'L',
        13: 'M',
        14: 'N',
        15: 'O',
        16: 'P',
        17: 'Q',
        18: 'R',
        20: 'T',
        21: 'U',
        22: 'V',
        23: 'W'
    }

    for col in range(1, sheet.max_column + 1):
        cell = sheet.cell(row=end_row, column=col)
        cell.style = 'header'
        if col == 5:
            cell.value = 'Итого'
        elif 9 <= col <= 18 or 20 <= col:
            cell.value = f'=SUM({rows_dict[col]}{start_row}:{rows_dict[col]}{end_row-1})'
            cell.number_format = '# ##0'
