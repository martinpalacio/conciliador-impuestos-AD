"""Mapeo de columnas del Mayor (según ERP) y del reporte externo (ARCA / IIBB)."""
from dataclasses import dataclass

import pandas as pd

from core.cleaning import safe_float


@dataclass(frozen=True)
class ColumnasMayor:
    asiento: str
    fecha: str
    detalle: str
    referencia: str
    ref_secundaria: str
    entidad: str
    cuit: str

    @classmethod
    def desde_config(cls, cfg):
        return cls(
            asiento=cfg.get("asiento", "ASIENTO"),
            fecha=cfg.get("fecha", "FECHA"),
            detalle=cfg.get("detalle", "DETALLE"),
            referencia=cfg.get("referencia", "REFERENCIA"),
            ref_secundaria=cfg.get("ref_secundaria", ""),
            entidad=cfg.get("entidad", "ENTIDAD"),
            cuit=cfg.get("cuit", "CUIT"),
        )


@dataclass(frozen=True)
class ColumnasExterno:
    cuit: str
    cert: str
    fecha: str
    comp: str
    razon: str
    monto: str
    tipo: str

    @classmethod
    def estandar(cls):
        """Nombres fijos: los usa el DataFrame que arma `construir_df_motor` (módulo IIBB)."""
        return cls(
            cuit="CUIT Agente Ret./Perc.",
            cert="Número Certificado",
            fecha="Fecha Ret./Perc.",
            comp="Número Comprobante",
            razon="Denominación o Razón Social",
            monto="Importe Ret./Perc.",
            tipo="Descripción Comprobante",
        )

    @classmethod
    def detectar(cls, df):
        """Detecta de forma flexible las columnas de un reporte de ARCA."""
        return cls(
            cuit=buscar_columna(df, ["CUIT Agente Ret./Perc.", "CUIT Agente", "CUIT"]),
            cert=buscar_columna(df, ["Número Certificado", "Nro Certificado", "Certificado"]),
            fecha=buscar_columna(df, ["Fecha Ret./Perc.", "Fecha Registro", "Fecha"]),
            comp=buscar_columna(df, ["Número Comprobante", "Nro Comprobante", "Comprobante"]),
            razon=buscar_columna(df, ["Denominación o Razón Social", "Razon Social", "Denominacion"]),
            monto=buscar_columna(df, ["Importe Ret./Perc.", "Importe", "Monto"]),
            tipo=buscar_columna(df, ["Descripción Comprobante", "Tipo Comprobante"]),
        )


def buscar_columna(df, posibles_nombres):
    """Busca de forma flexible una columna del DataFrame a partir de nombres candidatos."""
    for col in df.columns:
        col_clean = str(col).strip().lower()
        for p in posibles_nombres:
            if p.lower() in col_clean:
                return col
    return posibles_nombres[0] if posibles_nombres else ""


def obtener_monto_mayor(row, cols_config, priorizar_total=False):
    """
    Obtiene el monto del registro del Mayor evaluando DEBE, HABER y SALDO
    según la configuración de columnas del ERP activo.

    Si priorizar_total=True (módulo de percepciones IIBB) se toma primero la columna
    de total con signo (ImpTotal_Loc en Bejerman), donde las NC ya vienen en negativo.
    Si esa columna está vacía/0 se cae a la lógica habitual DEBE / HABER.
    """
    col_debe = cols_config.get("debe", "DEBE")
    col_haber = cols_config.get("haber", "HABER")
    col_saldo = cols_config.get("saldo", "SALDO")

    debe = safe_float(row.get(col_debe))
    haber = safe_float(row.get(col_haber))
    saldo = safe_float(row.get(col_saldo))

    if priorizar_total and saldo != 0:
        return round(saldo, 2)

    if debe != 0:
        return round(debe, 2)
    elif haber != 0:
        return round(-abs(haber), 2)
    elif saldo != 0:
        return round(saldo, 2)

    return 0.0


def serie_o_vacia(df, columna):
    """df[columna] si existe; si no, una Serie vacía (el resto del flujo la trata como NaN)."""
    return df.get(columna, pd.Series(dtype=object))
