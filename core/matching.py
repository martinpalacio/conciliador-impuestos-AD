"""
Motor de cruce. Cada estrategia es una función independiente y recibe/devuelve datos
simples, para poder testearlas por separado o agregar nuevas (p. ej. cruce por fecha+monto).

Contrato común: reciben los DataFrames ya preparados y los conjuntos de índices ya
conciliados, los actualizan in-place y devuelven la lista de conciliaciones nuevas.
"""
from config import TOLERANCIA_1A1, TOLERANCIA_LOTE


def _rango_fechas(serie_fechas):
    fechas = sorted(str(f)[:10] for f in serie_fechas.dropna())
    if len(fechas) > 1:
        return f"{fechas[0]} a {fechas[-1]}"
    return fechas[0] if fechas else ""


def _detalle_comprobante_mayor(row_m, cols_m):
    detalle = str(row_m.get(cols_m.detalle, ""))
    if cols_m.ref_secundaria and row_m.get(cols_m.ref_secundaria):
        detalle += f" [{row_m.get(cols_m.ref_secundaria)}]"
    return detalle


def match_1a1(df_ext, df_mayor, cols_e, cols_m, idx_ext_ok, idx_mayor_ok):
    """Cruce exacto: mismo monto (±tolerancia) y mismo CUIT cuando ambos lo tienen."""
    conciliadas = []

    for idx_e, row_e in df_ext.iterrows():
        monto_e = row_e["MONTO_CALC"]
        cuit_e = row_e["CUIT_CLEAN"]

        for idx_m, row_m in df_mayor.iterrows():
            if idx_m in idx_mayor_ok:
                continue

            monto_m = row_m["MONTO_CALC"]
            cuit_m = row_m["CUIT_CLEAN"]

            monto_coincide = abs(monto_e - monto_m) < TOLERANCIA_1A1
            cuit_coincide = (cuit_e == cuit_m) if (cuit_e and cuit_m) else True

            if monto_coincide and cuit_coincide:
                idx_ext_ok.add(idx_e)
                idx_mayor_ok.add(idx_m)

                conciliadas.append({
                    "cert_arca": row_e["CERT_STR"],
                    "fecha_arca": str(row_e.get(cols_e.fecha, ""))[:10],
                    "comp_arca": str(row_e.get(cols_e.comp, "")),
                    "cuit_arca": row_e.get(cols_e.cuit, ""),
                    "razon_arca": row_e.get(cols_e.razon, ""),
                    "monto_arca": monto_e,
                    "asiento_mayor": row_m["ASIENTO_STR"],
                    "fecha_mayor": str(row_m.get(cols_m.fecha, ""))[:10],
                    "comp_mayor": _detalle_comprobante_mayor(row_m, cols_m),
                    "cuit_mayor": row_m.get(cols_m.cuit, ""),
                    "razon_mayor": row_m.get(cols_m.entidad, ""),
                    "monto_mayor": monto_m,
                })
                break

    return conciliadas


def match_por_lote(df_ext, df_mayor, cols_e, cols_m, idx_ext_ok, idx_mayor_ok,
                   clave, etiqueta_clave, origen_label, usar_clave_como_cuit=False):
    """
    Cruce por lote: agrupa los registros pendientes por `clave` (columna de df: CUIT_CLEAN
    o ENTITY_KEY) y concilia el grupo si las sumas coinciden (±TOLERANCIA_LOTE).
    """
    conciliadas = []

    pend_e = df_ext[~df_ext.index.isin(idx_ext_ok)]
    pend_m = df_mayor[~df_mayor.index.isin(idx_mayor_ok)]
    claves_comunes = (set(pend_e[clave]) & set(pend_m[clave])) - {""}

    for k in sorted(claves_comunes):
        grupo_e = pend_e[pend_e[clave] == k]
        grupo_m = pend_m[pend_m[clave] == k]

        suma_e = round(grupo_e["MONTO_CALC"].sum(), 2)
        suma_m = round(grupo_m["MONTO_CALC"].sum(), 2)

        if abs(suma_e - suma_m) < TOLERANCIA_LOTE and suma_e != 0:
            idx_ext_ok.update(grupo_e.index)
            idx_mayor_ok.update(grupo_m.index)

            certs = ", ".join(sorted(set(grupo_e["CERT_STR"].astype(str))))
            asientos = ", ".join(sorted(set(grupo_m["ASIENTO_STR"].astype(str))))

            nombre = grupo_e.iloc[0].get(cols_e.razon) or grupo_m.iloc[0].get(cols_m.entidad)
            cuit_fallback = k if usar_clave_como_cuit else ""
            cuit_real = (
                grupo_e.iloc[0].get(cols_e.cuit) or grupo_m.iloc[0].get(cols_m.cuit) or cuit_fallback
            )

            conciliadas.append({
                "empresa": nombre,
                "cuit": cuit_real,
                "certs_arca": certs,
                "fecha_arca": _rango_fechas(grupo_e[cols_e.fecha]),
                "monto_arca": suma_e,
                "asientos_mayor": asientos,
                "fecha_mayor": _rango_fechas(grupo_m[cols_m.fecha]),
                "monto_mayor": suma_m,
                "obs": f"Conciliación por {etiqueta_clave} por lote "
                       f"({len(grupo_e)} reg. {origen_label} vs {len(grupo_m)} reg. Mayor)",
            })

    return conciliadas


def match_por_lote_cuit(df_ext, df_mayor, cols_e, cols_m, idx_ext_ok, idx_mayor_ok, origen_label):
    return match_por_lote(
        df_ext, df_mayor, cols_e, cols_m, idx_ext_ok, idx_mayor_ok,
        clave="CUIT_CLEAN", etiqueta_clave="CUIT", origen_label=origen_label,
        usar_clave_como_cuit=True,
    )


def match_por_lote_razon_social(df_ext, df_mayor, cols_e, cols_m, idx_ext_ok, idx_mayor_ok, origen_label):
    return match_por_lote(
        df_ext, df_mayor, cols_e, cols_m, idx_ext_ok, idx_mayor_ok,
        clave="ENTITY_KEY", etiqueta_clave="Razón Social", origen_label=origen_label,
        usar_clave_como_cuit=False,
    )


def listar_pendientes_externo(df_ext, idx_ext_ok, cols_e):
    pend = df_ext[~df_ext.index.isin(idx_ext_ok)]
    return [
        {
            "cert_arca": row["CERT_STR"],
            "fecha_arca": str(row.get(cols_e.fecha, ""))[:10],
            "tipo_comp": str(row.get(cols_e.tipo, "")),
            "comp_arca": str(row.get(cols_e.comp, "")),
            "cuit_arca": str(row.get(cols_e.cuit, "")),
            "razon_arca": row.get(cols_e.razon, ""),
            "monto_arca": row["MONTO_CALC"],
        }
        for _, row in pend.iterrows()
    ]


def listar_pendientes_mayor(df_mayor, idx_mayor_ok, cols_m):
    pend = df_mayor[~df_mayor.index.isin(idx_mayor_ok)]
    resultado = []
    for _, row in pend.iterrows():
        ref_val = str(row.get(cols_m.referencia, ""))
        if cols_m.ref_secundaria and row.get(cols_m.ref_secundaria):
            ref_val += f" / {row.get(cols_m.ref_secundaria)}"

        resultado.append({
            "asiento_mayor": row["ASIENTO_STR"],
            "fecha_mayor": str(row.get(cols_m.fecha, ""))[:10],
            "detalle_mayor": str(row.get(cols_m.detalle, "")),
            "ref_mayor": ref_val,
            "cuit_mayor": str(row.get(cols_m.cuit, "")),
            "razon_mayor": row.get(cols_m.entidad, ""),
            "monto_mayor": row["MONTO_CALC"],
        })
    return resultado
