import io
import math
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Conciliación Contable", page_icon="📊", layout="wide"
)

st.title("📊 Conciliación Contable de Retenciones")
st.markdown(
    "Sube los archivos **Mayor.xlsx** y **MisRetenciones.xlsx** para generar el"
    " reporte consolidado."
)

col1, col2 = st.columns(2)

with col1:
    file_mayor = st.file_uploader(
        "Subir Mayor.xlsx", type=["xlsx", "xls"], key="mayor"
    )

with col2:
    file_arca = st.file_uploader(
        "Subir MisRetenciones.xlsx", type=["xlsx", "xls"], key="arca"
    )


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


def obtener_monto_mayor(row):
    """
    Obtiene el monto del registro del Mayor evaluando DEBE, HABER y SALDO.
    Garantiza que celdas vacías (NaN en DEBE) permitan la lectura de HABER y SALDO.
    """
    debe = safe_float(row.get("DEBE"))
    haber = safe_float(row.get("HABER"))
    saldo = safe_float(row.get("SALDO"))

    if debe != 0:
        return round(debe, 2)
    elif haber != 0:
        return round(-abs(haber), 2)
    elif saldo != 0:
        return round(saldo, 2)
    
    return 0.0


def procesar_archivos(file_m, file_a):
    df_mayor = pd.read_excel(file_m).dropna(how="all")
    df_arca = pd.read_excel(file_a).dropna(how="all")

    # Calculamos montos numéricos limpios resolviendo NaNs
    df_mayor["MONTO_CALC"] = df_mayor.apply(obtener_monto_mayor, axis=1)
    df_arca["MONTO_CALC"] = df_arca.apply(
        lambda r: round(safe_float(r.get("Importe Ret./Perc.")), 2), axis=1
    )

    # Filtramos únicamente registros cuyo monto calculado sea distinto de cero (incluye negativos)
    df_mayor = df_mayor[df_mayor["MONTO_CALC"] != 0].copy()
    df_arca = df_arca[df_arca["MONTO_CALC"] != 0].copy()

    df_mayor["ASIENTO_STR"] = (
        df_mayor.get("ASIENTO", pd.Series(dtype=object))
        .fillna("")
        .astype(str)
        .str.replace(".0", "", regex=False)
    )
    df_arca["CERT_STR"] = (
        df_arca.get("Número Certificado", pd.Series(dtype=object))
        .fillna("")
        .astype(str)
        .str.replace(".0", "", regex=False)
    )

    conciliadas_1a1 = []
    conciliadas_lote = []
    pendientes_arca = []
    pendientes_mayor = []

    arca_matched_indices = set()
    mayor_matched_indices = set()

    for idx_a, row_a in df_arca.iterrows():
        monto_a = row_a["MONTO_CALC"]
        
        for idx_m, row_m in df_mayor.iterrows():
            if idx_m in mayor_matched_indices:
                continue
            
            monto_m = row_m["MONTO_CALC"]

            # Cruce exacto contemplando el signo de ambas partidas
            if abs(monto_a - monto_m) < 0.01:
                arca_matched_indices.add(idx_a)
                mayor_matched_indices.add(idx_m)
                conciliadas_1a1.append({
                    "cert_arca": row_a["CERT_STR"],
                    "fecha_arca": str(row_a.get("Fecha Ret./Perc.", ""))[:10],
                    "comp_arca": str(row_a.get("Número Comprobante", "")),
                    "razon_arca": row_a.get("Denominación o Razón Social", ""),
                    "monto_arca": monto_a,
                    "asiento_mayor": row_m["ASIENTO_STR"],
                    "fecha_mayor": str(row_m.get("FECHA", ""))[:10],
                    "comp_mayor": str(row_m.get("DETALLE", "")),
                    "razon_mayor": row_m.get("ENTIDAD", ""),
                    "monto_mayor": monto_m,
                })
                break

    unmatched_arca = df_arca[~df_arca.index.isin(arca_matched_indices)]
    unmatched_mayor = df_mayor[~df_mayor.index.isin(mayor_matched_indices)]

    for idx_a, row_a in unmatched_arca.iterrows():
        pendientes_arca.append({
            "cert_arca": row_a["CERT_STR"],
            "fecha_arca": str(row_a.get("Fecha Ret./Perc.", ""))[:10],
            "tipo_comp": str(row_a.get("Descripción Comprobante", "")),
            "comp_arca": str(row_a.get("Número Comprobante", "")),
            "razon_arca": row_a.get("Denominación o Razón Social", ""),
            "monto_arca": row_a["MONTO_CALC"],
        })

    for idx_m, row_m in unmatched_mayor.iterrows():
        pendientes_mayor.append({
            "asiento_mayor": row_m["ASIENTO_STR"],
            "fecha_mayor": str(row_m.get("FECHA", ""))[:10],
            "tipo_comp": "RECIBO / PV",
            "ref_mayor": str(row_m.get("REFERENCIA", "")),
            "razon_mayor": row_m.get("ENTIDAD", ""),
            "monto_mayor": row_m["MONTO_CALC"],
        })

    # Crear Excel
    wb = openpyxl.Workbook()
    ws_resumen = wb.active
    ws_resumen.title = "Resumen General"
    ws_1a1 = wb.create_sheet(title="Conciliadas 1a1")
    ws_lote = wb.create_sheet(title="Conciliadas por Lote")
    ws_arca = wb.create_sheet(title="Pendientes_ARCA")
    ws_mayor = wb.create_sheet(title="Pendientes_MAYOR")

    for ws in wb.worksheets:
        ws.views.sheetView[0].showGridLines = True

    font_title = Font(name="Arial", size=14, bold=True, color="1F4E79")
    font_header = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    font_bold = Font(name="Arial", size=10, bold=True)
    font_regular = Font(name="Arial", size=10)

    fill_resumen_hdr = PatternFill(
        start_color="1F4E79", end_color="1F4E79", fill_type="solid"
    )
    fill_1a1_hdr = PatternFill(
        start_color="2E75B6", end_color="2E75B6", fill_type="solid"
    )
    fill_lote_hdr = PatternFill(
        start_color="2E75B6", end_color="2E75B6", fill_type="solid"
    )
    fill_arca_hdr = PatternFill(
        start_color="C65911", end_color="C65911", fill_type="solid"
    )
    fill_mayor_hdr = PatternFill(
        start_color="595959", end_color="595959", fill_type="solid"
    )
    fill_total = PatternFill(
        start_color="F2F2F2", end_color="F2F2F2", fill_type="solid"
    )

    thin_side = Side(border_style="thin", color="D9D9D9")
    thin_border = Border(
        left=thin_side, right=thin_side, top=thin_side, bottom=thin_side
    )
    double_bottom = Side(border_style="double", color="000000")
    top_thin = Side(border_style="thin", color="000000")
    total_border = Border(top=top_thin, bottom=double_bottom)

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    fmt_currency = '"$ "#,##0.00;("$ "#,##0.00);"-"'

    # 1a1
    ws_1a1.cell(
        row=1, column=1, value="PARTIDAS CONCILIADAS EXACTAS (1 A 1)"
    ).font = font_title
    headers_1a1 = [
        "Certif. ARCA",
        "Fecha ARCA",
        "Comprobante",
        "Razón Social ARCA",
        "Monto ARCA ($)",
        "Asiento Mayor",
        "Fecha Mayor",
        "Comprobante Mayor",
        "Razón Social Mayor",
        "Monto Mayor ($)",
        "Diferencia ($)",
    ]
    for col_idx, text in enumerate(headers_1a1, 1):
        c = ws_1a1.cell(row=3, column=col_idx, value=text)
        c.font, c.fill, c.alignment = font_header, fill_1a1_hdr, align_center

    r_idx = 4
    for item in conciliadas_1a1:
        ws_1a1.cell(row=r_idx, column=1, value=item["cert_arca"]).alignment = (
            align_center
        )
        ws_1a1.cell(row=r_idx, column=2, value=item["fecha_arca"]).alignment = (
            align_center
        )
        ws_1a1.cell(row=r_idx, column=3, value=item["comp_arca"]).alignment = (
            align_center
        )
        ws_1a1.cell(row=r_idx, column=4, value=item["razon_arca"]).alignment = (
            align_left
        )
        c_m1 = ws_1a1.cell(row=r_idx, column=5, value=item["monto_arca"])
        c_m1.number_format, c_m1.alignment = fmt_currency, align_right

        ws_1a1.cell(row=r_idx, column=6, value=item["asiento_mayor"]).alignment = (
            align_center
        )
        ws_1a1.cell(row=r_idx, column=7, value=item["fecha_mayor"]).alignment = (
            align_center
        )
        ws_1a1.cell(row=r_idx, column=8, value=item["comp_mayor"]).alignment = (
            align_center
        )
        ws_1a1.cell(row=r_idx, column=9, value=item["razon_mayor"]).alignment = (
            align_left
        )
        c_m2 = ws_1a1.cell(row=r_idx, column=10, value=item["monto_mayor"])
        c_m2.number_format, c_m2.alignment = fmt_currency, align_right

        c_diff = ws_1a1.cell(row=r_idx, column=11, value=f"=E{r_idx}-J{r_idx}")
        c_diff.number_format, c_diff.alignment = fmt_currency, align_right

        for c in range(1, 12):
            ws_1a1.cell(row=r_idx, column=c).font = font_regular
            ws_1a1.cell(row=r_idx, column=c).border = thin_border
        r_idx += 1

    ws_1a1.cell(row=r_idx, column=1, value="TOTAL CONCILIADO 1 A 1").font = (
        font_bold
    )
    ws_1a1.cell(row=r_idx, column=1).alignment = align_left
    c_tot1 = ws_1a1.cell(row=r_idx, column=5, value=f"=SUM(E4:E{r_idx-1})")
    c_tot1.font, c_tot1.number_format = font_bold, fmt_currency
    c_tot2 = ws_1a1.cell(row=r_idx, column=10, value=f"=SUM(J4:J{r_idx-1})")
    c_tot2.font, c_tot2.number_format = font_bold, fmt_currency
    c_totdiff = ws_1a1.cell(row=r_idx, column=11, value=f"=SUM(K4:K{r_idx-1})")
    c_totdiff.font, c_totdiff.number_format = font_bold, fmt_currency

    for c in range(1, 12):
        ws_1a1.cell(row=r_idx, column=c).fill = fill_total
        ws_1a1.cell(row=r_idx, column=c).border = total_border

    row_1a1_total = r_idx
    cnt_1a1 = len(conciliadas_1a1)

    # Lote
    ws_lote.cell(
        row=1, column=1, value="PARTIDAS CONCILIADAS POR LOTE (AGRUPADAS)"
    ).font = font_title
    headers_lote = [
        "Grupo / Empresa",
        "Certificados ARCA Incluidos",
        "Fecha ARCA",
        "Monto Total ARCA ($)",
        "Asiento Mayor",
        "Fecha Mayor",
        "Monto Mayor ($)",
        "Diferencia ($)",
        "Observaciones",
    ]
    for col_idx, text in enumerate(headers_lote, 1):
        c = ws_lote.cell(row=4, column=col_idx, value=text)
        c.font, c.fill, c.alignment = font_header, fill_lote_hdr, align_center

    r_lote_idx = 5
    ws_lote.cell(
        row=r_lote_idx, column=1, value="TOTAL CONCILIADO POR LOTE"
    ).font = font_bold
    c_lt1 = ws_lote.cell(
        row=r_lote_idx, column=4, value=f"=SUM(D5:D{r_lote_idx-1})"
    )
    c_lt1.font, c_lt1.number_format = font_bold, fmt_currency
    c_lt2 = ws_lote.cell(
        row=r_lote_idx, column=7, value=f"=SUM(G5:G{r_lote_idx-1})"
    )
    c_lt2.font, c_lt2.number_format = font_bold, fmt_currency
    c_ltdiff = ws_lote.cell(
        row=r_lote_idx, column=8, value=f"=SUM(H5:H{r_lote_idx-1})"
    )
    c_ltdiff.font, c_ltdiff.number_format = font_bold, fmt_currency

    for c in range(1, 10):
        ws_lote.cell(row=r_lote_idx, column=c).fill = fill_total
        ws_lote.cell(row=r_lote_idx, column=c).border = total_border

    row_lote_total = r_lote_idx
    cnt_lote = len(conciliadas_lote)

    # ARCA
    ws_arca.cell(
        row=1,
        column=1,
        value="RETENCIONES PENDIENTES EN ARCA (NO REGISTRADAS EN MAYOR)",
    ).font = font_title
    headers_arca = [
        "Nro Certificado",
        "Fecha Reg.",
        "Tipo Comprobante",
        "Nro Comprobante",
        "Razón Social Agente",
        "Monto Retención ($)",
    ]
    for col_idx, text in enumerate(headers_arca, 1):
        c = ws_arca.cell(row=3, column=col_idx, value=text)
        c.font, c.fill, c.alignment = font_header, fill_arca_hdr, align_center

    r_arca_idx = 4
    for item in pendientes_arca:
        ws_arca.cell(
            row=r_arca_idx, column=1, value=item["cert_arca"]
        ).alignment = align_center
        ws_arca.cell(
            row=r_arca_idx, column=2, value=item["fecha_arca"]
        ).alignment = align_center
        ws_arca.cell(
            row=r_arca_idx, column=3, value=item["tipo_comp"]
        ).alignment = align_center
        ws_arca.cell(
            row=r_arca_idx, column=4, value=item["comp_arca"]
        ).alignment = align_center
        ws_arca.cell(
            row=r_arca_idx, column=5, value=item["razon_arca"]
        ).alignment = align_left
        c_m = ws_arca.cell(row=r_arca_idx, column=6, value=item["monto_arca"])
        c_m.number_format, c_m.alignment = fmt_currency, align_right

        for c in range(1, 7):
            ws_arca.cell(row=r_arca_idx, column=c).font = font_regular
            ws_arca.cell(row=r_arca_idx, column=c).border = thin_border
        r_arca_idx += 1

    ws_arca.cell(row=r_arca_idx, column=1, value="TOTAL PENDIENTE ARCA").font = (
        font_bold
    )
    c_totarca = ws_arca.cell(
        row=r_arca_idx, column=6, value=f"=SUM(F4:F{r_arca_idx-1})"
    )
    c_totarca.font, c_totarca.number_format, c_totarca.alignment = (
        font_bold,
        fmt_currency,
        align_right,
    )

    for c in range(1, 7):
        ws_arca.cell(row=r_arca_idx, column=c).fill = fill_total
        ws_arca.cell(row=r_arca_idx, column=c).border = total_border

    row_arca_total = r_arca_idx
    cnt_arca = len(pendientes_arca)

    # MAYOR
    ws_mayor.cell(
        row=1,
        column=1,
        value="REGISTROS PENDIENTES EN MAYOR (SIN CERTIFICADO EN ARCA)",
    ).font = font_title
    headers_mayor = [
        "Nro Asiento",
        "Fecha Contable",
        "Tipo Comprobante",
        "Referencia",
        "Cuenta / Descripción",
        "Monto Registrado ($)",
    ]
    for col_idx, text in enumerate(headers_mayor, 1):
        c = ws_mayor.cell(row=3, column=col_idx, value=text)
        c.font, c.fill, c.alignment = font_header, fill_mayor_hdr, align_center

    r_mayor_idx = 4
    for item in pendientes_mayor:
        ws_mayor.cell(
            row=r_mayor_idx, column=1, value=item["asiento_mayor"]
        ).alignment = align_center
        ws_mayor.cell(
            row=r_mayor_idx, column=2, value=item["fecha_mayor"]
        ).alignment = align_center
        ws_mayor.cell(
            row=r_mayor_idx, column=3, value=item["tipo_comp"]
        ).alignment = align_center
        ws_mayor.cell(
            row=r_mayor_idx, column=4, value=item["ref_mayor"]
        ).alignment = align_center
        ws_mayor.cell(
            row=r_mayor_idx, column=5, value=item["razon_mayor"]
        ).alignment = align_left
        c_m = ws_mayor.cell(row=r_mayor_idx, column=6, value=item["monto_mayor"])
        c_m.number_format, c_m.alignment = fmt_currency, align_right

        for c in range(1, 7):
            ws_mayor.cell(row=r_mayor_idx, column=c).font = font_regular
            ws_mayor.cell(row=r_mayor_idx, column=c).border = thin_border
        r_mayor_idx += 1

    ws_mayor.cell(
        row=r_mayor_idx, column=1, value="TOTAL PENDIENTE MAYOR"
    ).font = font_bold
    c_totmayor = ws_mayor.cell(
        row=r_mayor_idx, column=6, value=f"=SUM(F4:F{r_mayor_idx-1})"
    )
    c_totmayor.font, c_totmayor.number_format, c_totmayor.alignment = (
        font_bold,
        fmt_currency,
        align_right,
    )

    for c in range(1, 7):
        ws_mayor.cell(row=r_mayor_idx, column=c).fill = fill_total
        ws_mayor.cell(row=r_mayor_idx, column=c).border = total_border

    row_mayor_total = r_mayor_idx
    cnt_mayor = len(pendientes_mayor)

    # Resumen
    ws_resumen.cell(
        row=1,
        column=1,
        value="CONCILIACIÓN DE RETENCIONES - RESUMEN CONSOLIDADO",
    ).font = font_title
    headers_resumen = [
        "Categoría / Pestaña",
        "Cantidad de Reg.",
        "Total ARCA ($)",
        "Total Mayor ($)",
        "Diferencia ($)",
    ]
    for col_idx, text in enumerate(headers_resumen, 1):
        c = ws_resumen.cell(row=4, column=col_idx, value=text)
        c.font, c.fill, c.alignment = font_header, fill_resumen_hdr, align_center

    resumen_rows = [
        (
            "1. Conciliaciones Exactas (1 a 1)",
            cnt_1a1,
            f"='Conciliadas 1a1'!E{row_1a1_total}",
            f"='Conciliadas 1a1'!J{row_1a1_total}",
            "=C5-D5",
        ),
        (
            "2. Conciliaciones por Lote (N a 1)",
            cnt_lote,
            f"='Conciliadas por Lote'!D{row_lote_total}",
            f"='Conciliadas por Lote'!G{row_lote_total}",
            "=C6-D6",
        ),
        (
            "3. Pendientes en ARCA (Sin Mayor)",
            cnt_arca,
            f"=Pendientes_ARCA!F{row_arca_total}",
            0,
            "=C7-D7",
        ),
        (
            "4. Pendientes en MAYOR (Sin ARCA)",
            cnt_mayor,
            0,
            f"=Pendientes_MAYOR!F{row_mayor_total}",
            "=C8-D8",
        ),
    ]

    for idx, rdata in enumerate(resumen_rows, 5):
        ws_resumen.cell(row=idx, column=1, value=rdata[0]).alignment = align_left
        ws_resumen.cell(row=idx, column=2, value=rdata[1]).alignment = align_center

        c_arca = ws_resumen.cell(row=idx, column=3, value=rdata[2])
        c_arca.number_format, c_arca.alignment = fmt_currency, align_right

        c_myr = ws_resumen.cell(row=idx, column=4, value=rdata[3])
        c_myr.number_format, c_myr.alignment = fmt_currency, align_right

        c_diff = ws_resumen.cell(row=idx, column=5, value=rdata[4])
        c_diff.number_format, c_diff.alignment = fmt_currency, align_right

        for c in range(1, 6):
            ws_resumen.cell(row=idx, column=c).font = font_regular
            ws_resumen.cell(row=idx, column=c).border = thin_border

    # TOTAL GENERAL
    ws_resumen.cell(row=9, column=1, value="TOTAL GENERAL EMANADO").font = font_bold
    ws_resumen.cell(row=9, column=2, value="=SUM(B5:B8)").font = font_bold
    ws_resumen.cell(row=9, column=2).alignment = align_center

    c_tot_arca = ws_resumen.cell(row=9, column=3, value="=SUM(C5:C8)")
    c_tot_arca.font, c_tot_arca.number_format, c_tot_arca.alignment = (
        font_bold,
        fmt_currency,
        align_right,
    )

    c_tot_myr = ws_resumen.cell(row=9, column=4, value="=SUM(D5:D8)")
    c_tot_myr.font, c_tot_myr.number_format, c_tot_myr.alignment = (
        font_bold,
        fmt_currency,
        align_right,
    )

    c_tot_diff = ws_resumen.cell(row=9, column=5, value="=C9-D9")
    c_tot_diff.font, c_tot_diff.number_format, c_tot_diff.alignment = (
        font_bold,
        fmt_currency,
        align_right,
    )

    for c in range(1, 6):
        ws_resumen.cell(row=9, column=c).fill = fill_total
        ws_resumen.cell(row=9, column=c).border = total_border

    for ws in wb.worksheets:
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.row in [1, 2]:
                    continue
                val_str = str(cell.value or "")
                if cell.number_format and "$" in cell.number_format:
                    val_str += "    "
                max_len = max(max_len, len(val_str))
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


if file_mayor and file_arca:
    if st.button("⚙️ Procesar Conciliación", type="primary"):
        with st.spinner("Procesando datos y armando informe..."):
            excel_bytes = procesar_archivos(file_mayor, file_arca)
            st.success("✅ ¡Conciliación completada!")
            st.download_button(
                label="📥 Descargar Archivo Excel Conciliado",
                data=excel_bytes,
                file_name="Conciliacion_Retenciones.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
```

### Principales cambios implementados:
1. **Sanitización de valores nulos con `safe_float`:** Filtra los `NaN` de Pandas y asegura que celdas de `DEBE` vacías no bloqueen el cálculo de `HABER` o `SALDO`.
2. **Precálculo de columna `MONTO_CALC`:** Todos los registros del Mayor y ARCA son normalizados antes del filtrado y cruce de datos.
3. **Manejo explícito de importes negativos:** Los registros con valor en `HABER` o `SALDO` negativo se incluyen y escriben correctamente en las pestañas correspondientes (`Pendientes_MAYOR`, `Conciliadas 1a1`, etc.).
