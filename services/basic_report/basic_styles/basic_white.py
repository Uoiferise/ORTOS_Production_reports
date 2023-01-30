from openpyxl.styles import NamedStyle, Font, Alignment, PatternFill


basic_white = NamedStyle(name='white')
basic_white.font = Font(name='Arial', bold=False, size=8)
basic_white.alignment = Alignment(horizontal='center', vertical='center')
basic_white.fill = PatternFill("solid", fgColor="FFFFFF")
