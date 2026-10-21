"""Parser del .txt de percepciones IIBB y del padrón de proveedores."""
import io
import re
from datetime import datetime

import pandas as pd

from config import SIN_RAZON_SOCIAL
from core.cleaning import clean_cuit, parse_monto_txt
from core.io_utils import decodificar_texto

# Formato por defecto del sistema (ancho fijo):
# JJJ CUIT(13) FECHA(10) PV(4) NRO(8) COMPROBANTE(1) TIPO(1) MONTO
# Ej: 90130-50000173-501/10/2025700000095099O 00023732,45
RE_PERCEPCION_FIJO = re.compile(
    r"^(\d{3})(\d{2}-\d{8}-\d)(\d{2}/\d{2}/\d{4})(\d{4})(\d{8})([A-Z])([A-Z ])(\S+)$"
)
RE_CUIT_GUIONES = re.compile(r"^\d{2}-\d{8}-\d$")

COLUMNAS_PERCEPCIONES = [
    "Jurisdicción", "CUIT", "Fecha", "Nro Referencia", "Nro Comprobante",
    "Comprobante", "Tipo Comprobante", "Monto",
]


def _fecha_a_dt(fecha_str):
    return datetime.strptime(fecha_str.strip(), "%d/%m/%Y")


def _parsear_ancho_fijo(lineas):
    registros = []
    for linea in lineas:
        m = RE_PERCEPCION_FIJO.match(linea.strip())
        if not m:
            continue
        jur, cuit, fecha, ref, nro, comp, tipo, monto = m.groups()
        registros.append((jur, cuit, fecha, int(ref), int(nro), comp, tipo.strip(), parse_monto_txt(monto)))
    return registros


def _parsear_tabulado(lineas):
    registros = []
    for linea in lineas:
        partes = linea.split("\t")
        if len(partes) < 8:
            continue
        jur, cuit = partes[0].strip(), partes[1].strip()
        if not (re.fullmatch(r"\d{3}", jur) and RE_CUIT_GUIONES.match(cuit)):
            continue
        try:
            fecha = _fecha_a_dt(partes[2]).strftime("%d/%m/%Y")
            registros.append((jur, cuit, fecha, int(partes[3]), int(partes[4]),
                              partes[5].strip(), partes[6].strip(), parse_monto_txt(partes[7])))
        except ValueError:
            continue
    return registros


def parse_percepciones_txt(raw):
    """
    Interpreta el .txt de percepciones IIBB y devuelve un DataFrame con:
    Jurisdicción, CUIT, Fecha, Nro Referencia, Nro Comprobante, Comprobante,
    Tipo Comprobante, Monto.

    Lee el formato de ancho fijo (primer bloque del sistema). Si el archivo no tiene
    líneas en ese formato, intenta leer el formato ya tabulado (separado por tabs).
    Si el archivo trae ambos bloques, se usa solo el de ancho fijo para no duplicar.
    """
    lineas = decodificar_texto(raw).splitlines()
    registros = _parsear_ancho_fijo(lineas) or _parsear_tabulado(lineas)
    return pd.DataFrame(registros, columns=COLUMNAS_PERCEPCIONES)


def cargar_padron_proveedores(raw):
    """
    Lee el listado de proveedores (txt separado por tabs: Cód., Razón Social,
    Tipo Docum., Nro. Docum.) y devuelve {cuit_solo_digitos: razon_social}.
    Si un CUIT está repetido se usa el primero.
    """
    texto = decodificar_texto(raw)
    df = pd.read_csv(io.StringIO(texto), sep="\t", dtype=str)
    col_rs = next((c for c in df.columns if "raz" in str(c).lower()), df.columns[1])
    col_doc = next((c for c in df.columns if "nro" in str(c).lower()), df.columns[-1])

    padron = {}
    for rs, doc in zip(df[col_rs], df[col_doc]):
        k = clean_cuit(doc)
        if k and k not in padron and pd.notna(rs):
            padron[k] = str(rs).strip()
    return padron


def agregar_razon_social(df, padron):
    df = df.copy()
    df["Razón Social"] = df["CUIT"].apply(lambda c: padron.get(clean_cuit(c), SIN_RAZON_SOCIAL))
    return df


def construir_df_motor(df):
    """
    Adapta el DataFrame de percepciones al formato que consume el motor de conciliación
    (mismas columnas que un reporte de ARCA), para que la lógica de cruce sea idéntica.
    El orden de columnas importa: 'Número Comprobante' va antes que 'Descripción Comprobante'.
    """
    df = df.reset_index(drop=True)
    return pd.DataFrame({
        "CUIT Agente Ret./Perc.": df["CUIT"],
        "Número Certificado": df["Nro Comprobante"].astype(str),
        "Fecha Ret./Perc.": df["Fecha"].apply(lambda f: _fecha_a_dt(f).strftime("%Y-%m-%d")),
        "Número Comprobante": (
            df["Nro Referencia"].astype(str).str.zfill(4) + "-" + df["Nro Comprobante"].astype(str).str.zfill(8)
        ),
        "Descripción Comprobante": df["Comprobante"].astype(str) + df["Tipo Comprobante"].astype(str),
        "Denominación o Razón Social": df["Razón Social"],
        "Importe Ret./Perc.": df["Monto"],
    })
