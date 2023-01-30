from openpyxl.styles import NamedStyle, Side, Border, Font


basic_info = NamedStyle(name='info')
thin = Side(border_style="thin", color="000000")
basic_info.border = Border(top=thin, left=thin, right=thin, bottom=thin)
basic_info.font = Font(name='Arial', bold=False, size=8)
