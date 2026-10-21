"""Funciones puras de limpieza y normalización de datos (sin dependencias de UI)."""
import math
import re

import pandas as pd

_NOISE_WORDS = {"SA", "SAU", "SRL", "SACIF", "LIMITADA", "LTD", "INC", "CORP", "SOCIEDAD", "ANONIMA"}


def safe_float(val):
    """
    Convierte de forma segura cualquier valor (NaN, None, cadenas formateadas) a float.
    Evita los falsos positivos de evaluación booleana de NaN en Python.
    """
    if pd.isna(val) or val is None:
        return 0.0
    try:
        if isinstance(val, str):
            val = val.replace("$", "").replace(" ", "").replace(".", "").replace(",", ".")
        v = float(val)
        return 0.0 if math.isnan(v) else v
    except (ValueError, TypeError):
        return 0.0


def clean_cuit(cuit_val):
    """
    Limpia y normaliza un CUIT/CUIL dejando solo dígitos numéricos.
    Mantiene ceros a la izquierda si los hay.
    """
    if pd.isna(cuit_val) or cuit_val is None:
        return ""
    # Convertir a string sin decimales flotantes .0
    s_cuit = str(cuit_val).replace(".0", "").strip()
    digits = re.sub(r"\D", "", s_cuit)
    return digits if len(digits) >= 8 else s_cuit.strip()


def clean_entity_key(nombre):
    """Sanitiza y normaliza el nombre de la Razón Social / Entidad como criterio secundario."""
    if not nombre or pd.isna(nombre):
        return ""
    s = str(nombre).upper()
    s = re.sub(r"[^A-Z0-9]", " ", s)
    words = s.split()
    filtered = [w for w in words if w not in _NOISE_WORDS]
    return " ".join(filtered) if filtered else " ".join(words)


def parse_monto_txt(s):
    """
    Convierte '00023732,45' / '4755,6' / '-0001234,50' a float.
    Las notas de crédito vienen con signo negativo: se respeta el signo.
    """
    s = str(s).strip()
    negativo = "-" in s or "(" in s
    t = re.sub(r"[^\d,.]", "", s)
    if "," in t:
        t = t.replace(".", "").replace(",", ".")
    try:
        v = float(t)
    except ValueError:
        return 0.0
    return -abs(v) if negativo else v
