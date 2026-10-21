"""Orquestador: carga → preparación → cruce (por estrategias) → reporte."""
import pandas as pd

from config import ERP_CONFIGS, ERP_DEFAULT
from core.columns import ColumnasExterno, ColumnasMayor
from core.io_utils import validar_formato_archivo
from core.matching import (
    listar_pendientes_externo, listar_pendientes_mayor, match_1a1,
    match_por_lote_cuit, match_por_lote_razon_social,
)
from core.preparacion import preparar_externo, preparar_mayor
from reports.excel_report import generar_excel


def cargar_excel(archivo):
    """Valida que sea .xlsx y lo lee, descartando filas totalmente vacías."""
    validar_formato_archivo(archivo)
    return pd.read_excel(archivo).dropna(how="all")


def ejecutar_cruce(df_ext, df_mayor, cols_e, cols_m, origen_label):
    """
    Aplica las estrategias de cruce en orden de prioridad:
    1) 1 a 1  2) lote por CUIT  3) lote por Razón Social.
    Para agregar una estrategia nueva, escribir su función en core/matching.py y sumarla acá.
    """
    idx_ext_ok, idx_mayor_ok = set(), set()

    conciliadas_1a1 = match_1a1(df_ext, df_mayor, cols_e, cols_m, idx_ext_ok, idx_mayor_ok)
    conciliadas_lote = match_por_lote_cuit(
        df_ext, df_mayor, cols_e, cols_m, idx_ext_ok, idx_mayor_ok, origen_label
    )
    conciliadas_lote += match_por_lote_razon_social(
        df_ext, df_mayor, cols_e, cols_m, idx_ext_ok, idx_mayor_ok, origen_label
    )

    return {
        "conciliadas_1a1": conciliadas_1a1,
        "conciliadas_lote": conciliadas_lote,
        "pendientes_arca": listar_pendientes_externo(df_ext, idx_ext_ok, cols_e),
        "pendientes_mayor": listar_pendientes_mayor(df_mayor, idx_mayor_ok, cols_m),
    }


def procesar_archivos(
    file_m,
    file_a,
    erp_seleccionado,
    df_arca_override=None,
    origen_label="ARCA",
    concepto="RETENCIONES",
    subtitulo="",
    priorizar_total_mayor=False,
):
    """
    Concilia el Mayor contra un reporte externo.

    - Modo ARCA (por defecto): file_a es el Excel de MisRetenciones.
    - Modo IIBB: se pasa df_arca_override (ya normalizado con construir_df_motor),
      origen_label="IIBB", concepto="PERCEPCIONES IIBB" y priorizar_total_mayor=True.

    Devuelve (excel_en_BytesIO, stats).
    """
    config_erp = ERP_CONFIGS.get(erp_seleccionado, ERP_CONFIGS[ERP_DEFAULT])
    cols_m = ColumnasMayor.desde_config(config_erp)

    df_mayor = cargar_excel(file_m)
    if df_arca_override is None:
        df_ext = cargar_excel(file_a)
        cols_e = ColumnasExterno.detectar(df_ext)
    else:
        df_ext = df_arca_override.copy().dropna(how="all")
        cols_e = ColumnasExterno.estandar()

    df_mayor = preparar_mayor(df_mayor, config_erp, cols_m, priorizar_total_mayor)
    df_ext = preparar_externo(df_ext, cols_e)

    resultado = ejecutar_cruce(df_ext, df_mayor, cols_e, cols_m, origen_label)
    excel = generar_excel(resultado, erp_seleccionado, origen_label, concepto, subtitulo)

    stats = {
        "cnt_1a1": len(resultado["conciliadas_1a1"]),
        "cnt_lote": len(resultado["conciliadas_lote"]),
        "cnt_arca": len(resultado["pendientes_arca"]),
        "cnt_mayor": len(resultado["pendientes_mayor"]),
        **resultado,
    }
    return excel, stats
