"""Tests del formateador de Mayor: usan un export sintético con la estructura nativa del ERP."""
import io

import openpyxl

from parsers.autodealer import calcular_control, dividir_concepto, formatear_mayor


def _export_nativo(saldo_final="900"):
    """Estructura del MayorPRN: SALDO INICIAL, movimientos (sin encabezado), SALDO FINAL, pie."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append([None, None, None, "SALDO INICIAL", None, None, "0", None, None])
    ws.append(["1", "05/01/2026", "TR30-50004946-0", "RC-X-0100-00100873 / TR30-50004946-0 / SANCOR SEGUROS",
               "1000.0000000001", None, "1000", "Administracion", None])
    ws.append(["2", "06/01/2026", None, "DJ IVA 02/2026", None, "100", "900", "Administracion", None])
    ws.append([None, None, None, "SALDO FINAL", None, None, saldo_final, None, None])
    ws.append(["MayorPRN.rpt", None, None, None, None, None, "Hoja", "1", "1"])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def test_dividir_concepto_tres_partes():
    assert dividir_concepto("RC-X-1 / PV 12132/0 / ARCO, SRL")[:3] == ("RC-X-1", "PV 12132", "ARCO, SRL")


def test_dividir_concepto_no_rompe_barra_sin_espacios():
    # "DJ IVA 02/2026" no se parte: el separador es " / " con espacios
    assert dividir_concepto("DJ IVA 02/2026") == ("DJ IVA 02/2026", "", "DJ IVA 02/2026", False)


def test_formatear_mayor_por_columnas_y_control_ok():
    r = formatear_mayor(_export_nativo())
    assert r.metodo == "columnas" and len(r.df) == 2
    fila = r.df.iloc[0]
    assert fila["DEBE"] == 1000.0                      # sin ruido de coma flotante
    assert fila["DETALLE"] == "RC-X-0100-00100873" and fila["ENTIDAD"] == "SANCOR SEGUROS"
    assert r.df.iloc[1]["SALDO"] == -100.0             # SALDO = Debe − Haber por movimiento
    assert r.control["ok"] is True and r.control["sin_tres_partes"] == 1


def test_control_detecta_saldo_que_no_cierra():
    r = formatear_mayor(_export_nativo(saldo_final="999"))
    assert r.control["ok"] is False and r.control["diferencia"] == -99.0


def test_control_sin_saldo_final_no_se_puede_validar():
    r = formatear_mayor(_export_nativo())
    assert calcular_control(r.df, 0.0, None)["ok"] is None


def test_tabla_con_otro_layout_cae_al_camino_por_texto():
    # Asiento y fecha en cols 0-1, pero sin importes numéricos en las posiciones del export nativo
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["12345", "01/10/2025", "PV", "12132/0", "RC-X-0100-00104458", "ACME", "SA", "1.000,00", "x"])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    assert formatear_mayor(buf).metodo == "texto"
