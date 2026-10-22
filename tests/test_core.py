"""Tests del núcleo: no necesitan Streamlit ni archivos. Ejecutar con: pytest"""
import pandas as pd

from core.cleaning import clean_cuit, clean_entity_key, parse_monto_txt, safe_float
from core.columns import ColumnasExterno, ColumnasMayor
from core.matching import match_1a1, match_por_lote_cuit
from reports.excel_helpers import formula_suma

COLS_E = ColumnasExterno.estandar()
COLS_M = ColumnasMayor.desde_config({"asiento": "A", "fecha": "F", "detalle": "D", "entidad": "E", "cuit": "C"})


def _df(filas):
    return pd.DataFrame(filas, columns=[
        "MONTO_CALC", "CUIT_CLEAN", "ENTITY_KEY", "CERT_STR", "ASIENTO_STR",
        COLS_E.fecha, COLS_M.fecha,
    ])


def test_safe_float_formatos():
    assert safe_float("$ 1.234,56") == 1234.56
    assert safe_float(None) == 0.0
    assert safe_float("abc") == 0.0


def test_clean_cuit_y_entidad():
    assert clean_cuit("30-11111111-1") == "30111111111"
    assert clean_cuit(None) == ""
    assert clean_entity_key("ACME SOCIEDAD ANONIMA") == "ACME"


def test_parse_monto_txt_respeta_signo():
    assert parse_monto_txt("00023732,45") == 23732.45
    assert parse_monto_txt("-0001234,50") == -1234.50


def test_match_1a1_concilia_y_marca_indices():
    ext = _df([[100.0, "111", "A", "c1", "", "2025-01-01", ""]])
    may = _df([[100.0, "111", "A", "", "9", "", "2025-01-01"]])
    ok_e, ok_m = set(), set()
    res = match_1a1(ext, may, COLS_E, COLS_M, ok_e, ok_m)
    assert len(res) == 1 and ok_e == {0} and ok_m == {0}


def test_match_por_lote_cuit_suma_grupo():
    ext = _df([[60.0, "111", "A", "c1", "", "2025-01-01", ""], [40.0, "111", "A", "c2", "", "2025-01-02", ""]])
    may = _df([[100.0, "111", "A", "", "9", "", "2025-01-03"]])
    ok_e, ok_m = set(), set()
    res = match_por_lote_cuit(ext, may, COLS_E, COLS_M, ok_e, ok_m, "ARCA")
    assert len(res) == 1 and res[0]["monto_arca"] == 100.0 and ok_e == {0, 1}


def test_formula_suma_vacia_no_es_circular():
    assert formula_suma("G", 4, 4) == 0          # sin filas: evita =SUM(G4:G3)
    assert formula_suma("G", 4, 7) == "=SUM(G4:G6)"


def test_match_1a1_toma_la_primera_fila_libre_del_mayor():
    # Dos filas del Mayor con el mismo importe: cada fila externa se queda con la primera libre, en orden
    ext = _df([[100.0, "", "A", "c1", "", "2025-01-01", ""], [100.0, "", "A", "c2", "", "2025-01-02", ""]])
    may = _df([[100.0, "", "A", "", "7", "", "2025-01-03"], [100.0, "", "A", "", "8", "", "2025-01-04"]])
    res = match_1a1(ext, may, COLS_E, COLS_M, set(), set())
    assert [r["asiento_mayor"] for r in res] == ["7", "8"]


def test_match_1a1_no_cruza_cuits_distintos_y_respeta_tolerancia():
    ext = _df([[100.0, "111", "A", "c1", "", "2025-01-01", ""], [200.0, "", "A", "c2", "", "2025-01-01", ""]])
    may = _df([[100.0, "222", "A", "", "1", "", "2025-01-03"], [200.02, "", "A", "", "2", "", "2025-01-03"]])
    assert match_1a1(ext, may, COLS_E, COLS_M, set(), set()) == []   # CUIT distinto / diferencia de 2 centavos
