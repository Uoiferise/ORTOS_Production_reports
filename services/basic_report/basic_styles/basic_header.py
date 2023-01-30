from openpyxl.styles import NamedStyle, Font, PatternFill, Side, Border, Alignment


basic_header = NamedStyle(name='header')
basic_header.font = Font(name='Arial', bold=True, size=10)
basic_header.fill = PatternFill("solid", fgColor="D6E5CB")
thin = Side(border_style="thin", color="000000")
basic_header.border = Border(top=thin, left=thin, right=thin, bottom=thin)
basic_header.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
