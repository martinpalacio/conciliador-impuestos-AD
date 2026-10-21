"""Preparación de los DataFrames: montos numéricos y claves de cruce normalizadas."""
import pandas as pd

from core.cleaning import clean_cuit, clean_entity_key, safe_float
from core.columns import obtener_monto_mayor, serie_o_vacia


def _a_texto_sin_decimal(serie):
    return serie.fillna("").astype(str).str.replace(".0", "", regex=False)


def preparar_mayor(df_mayor, config_erp, cols, priorizar_total=False):
    """Agrega MONTO_CALC, CUIT_CLEAN, ENTITY_KEY y ASIENTO_STR; descarta montos en cero."""
    df = df_mayor.copy()
    df["MONTO_CALC"] = df.apply(lambda r: obtener_monto_mayor(r, config_erp, priorizar_total), axis=1)
    df = df[df["MONTO_CALC"] != 0].copy()

    df["CUIT_CLEAN"] = serie_o_vacia(df, cols.cuit).apply(clean_cuit)
    df["ENTITY_KEY"] = serie_o_vacia(df, cols.entidad).apply(clean_entity_key)
    df["ASIENTO_STR"] = _a_texto_sin_decimal(serie_o_vacia(df, cols.asiento))
    return df


def preparar_externo(df_externo, cols):
    """Agrega MONTO_CALC, CUIT_CLEAN, ENTITY_KEY y CERT_STR; descarta montos en cero."""
    df = df_externo.copy()
    df["MONTO_CALC"] = df.apply(lambda r: round(safe_float(r.get(cols.monto)), 2), axis=1)
    df = df[df["MONTO_CALC"] != 0].copy()

    df["CUIT_CLEAN"] = serie_o_vacia(df, cols.cuit).apply(clean_cuit)
    df["ENTITY_KEY"] = serie_o_vacia(df, cols.razon).apply(clean_entity_key)
    df["CERT_STR"] = _a_texto_sin_decimal(serie_o_vacia(df, cols.cert))
    return df
