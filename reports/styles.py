"""Estilos compartidos para los reportes Excel."""
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

FONT_TITLE = Font(name="Arial", size=14, bold=True, color="1F4E79")
FONT_HEADER = Font(name="Arial", size=10, bold=True, color="FFFFFF")
FONT_BOLD = Font(name="Arial", size=10, bold=True)
FONT_REGULAR = Font(name="Arial", size=10)


def _fill(hex_color):
    return PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")


FILL_RESUMEN = _fill("1F4E79")
FILL_1A1 = _fill("2E75B6")
FILL_LOTE = _fill("2E75B6")
FILL_PEND_EXTERNO = _fill("C65911")
FILL_PEND_MAYOR = _fill("595959")
FILL_TOTAL = _fill("F2F2F2")

_THIN = Side(border_style="thin", color="D9D9D9")
THIN_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
TOTAL_BORDER = Border(
    top=Side(border_style="thin", color="000000"),
    bottom=Side(border_style="double", color="000000"),
)

ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")

FMT_CURRENCY = '"$ "#,##0.00;("$ "#,##0.00);"-"'
