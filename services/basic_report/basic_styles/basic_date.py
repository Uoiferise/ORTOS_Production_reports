from openpyxl.styles import NamedStyle, Side, Border, Font, Alignment


basic_date = NamedStyle(name='date')
thin = Side(border_style="thin", color="000000")
basic_date.border = Border(top=thin, left=thin, right=thin, bottom=thin)
basic_date.font = Font(name='Arial', bold=True, size=10)
basic_date.alignment = Alignment(horizontal='center', vertical='center')
