import io

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import pandas as pd
import streamlit as st


# ============================================================
# CONFIGURACIÓN DE LA APP
# ============================================================

st.set_page_config(
    page_title="Conciliación Contable",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Conciliación Contable de Retenciones")

st.markdown(
    "Sube los archivos **Mayor.xlsx** y **MisRetenciones.xlsx** para generar el "
    "reporte consolidado."
)


col1, col2 = st.columns(2)

with col1:
    file_mayor = st.file_uploader(
        "Subir Mayor.xlsx",
        type=["xlsx", "xls"],
        key="mayor"
    )

with col2:
    file_arca = st.file_uploader(
        "Subir MisRetenciones.xlsx",
        type=["xlsx", "xls"],
        key="arca"
    )


# ============================================================
# FUNCIÓN PRINCIPAL
# ============================================================

def procesar_archivos(file_m, file_a):

    # --------------------------------------------------------
    # LEER ARCHIVOS
    # --------------------------------------------------------

    df_mayor = pd.read_excel(file_m).dropna(how="all")
    df_arca = pd.read_excel(file_a).dropna(how="all")


    # --------------------------------------------------------
    # VALIDAR COLUMNAS NECESARIAS
    # --------------------------------------------------------

    columnas_mayor_requeridas = [
        "SALDO",
        "ASIENTO",
        "FECHA",
    ]

    columnas_arca_requeridas = [
        "Importe Ret./Perc.",
        "Número Certificado",
        "Fecha Ret./Perc.",
        "Denominación o Razón Social",
    ]

    faltantes_mayor = [
        col for col in columnas_mayor_requeridas
        if col not in df_mayor.columns
    ]

    faltantes_arca = [
        col for col in columnas_arca_requeridas
        if col not in df_arca.columns
    ]

    if faltantes_mayor:
        raise ValueError(
            "Faltan estas columnas en Mayor.xlsx: "
            + ", ".join(faltantes_mayor)
        )

    if faltantes_arca:
        raise ValueError(
            "Faltan estas columnas en MisRetenciones.xlsx: "
            + ", ".join(faltantes_arca)
        )


    # --------------------------------------------------------
    # CONVERTIR IMPORTES A NÚMEROS
    # --------------------------------------------------------

    df_mayor["SALDO"] = pd.to_numeric(
        df_mayor["SALDO"],
        errors="coerce"
    )

    df_arca["Importe Ret./Perc."] = pd.to_numeric(
        df_arca["Importe Ret./Perc."],
        errors="coerce"
    )


    # --------------------------------------------------------
    # FILTRAR MAYOR
    #
    # IMPORTANTE:
    # Antes se utilizaba:
    #
    # SALDO > 0
    #
    # Eso eliminaba todas las partidas negativas.
    #
    # Ahora se conservan tanto los SALDOS positivos como
    # los negativos.
    #
    # Solo se excluyen:
    # - valores vacíos
    # - SALDO = 0
    # --------------------------------------------------------

    df_mayor = df_mayor[
        df_mayor["SALDO"].notna()
        & (df_mayor["SALDO"] != 0)
    ].copy()


    # --------------------------------------------------------
    # FILTRAR ARCA
    #
    # Se mantienen también importes negativos.
    # Solo se excluyen:
    # - valores vacíos
    # - importes = 0
    # --------------------------------------------------------

    df_arca = df_arca[
        df_arca["Importe Ret./Perc."].notna()
        & (df_arca["Importe Ret./Perc."] != 0)
    ].copy()


    # --------------------------------------------------------
    # NORMALIZAR IDENTIFICADORES
    # --------------------------------------------------------

    df_mayor["ASIENTO_STR"] = (
        df_mayor["ASIENTO"]
        .astype(str)
        .str.replace(".0", "", regex=False)
    )

    df_arca["CERT_STR"] = (
        df_arca["Número Certificado"]
        .astype(str)
        .str.replace(".0", "", regex=False)
    )


    # --------------------------------------------------------
    # LISTAS DE RESULTADOS
    # --------------------------------------------------------

    conciliadas_1a1 = []
    conciliadas_lote = []
    pendientes_arca = []
    pendientes_mayor = []


    # --------------------------------------------------------
    # ÍNDICES YA CONCILIADOS
    # --------------------------------------------------------

    arca_matched_indices = set()
    mayor_matched_indices = set()


    # ========================================================
    # CONCILIACIÓN 1 A 1
    #
    # ARCA:
    #   Importe Ret./Perc.
    #
    # MAYOR:
    #   SALDO
    #
    # IMPORTANTE:
    # NO se utiliza abs() sobre los importes.
    #
    # Por lo tanto:
    #
    #     ARCA 1000  = MAYOR 1000   -> CONCILIA
    #     ARCA -1000 = MAYOR -1000  -> CONCILIA
    #
    #     ARCA 1000  != MAYOR -1000 -> NO CONCILIA
    #     ARCA -1000 != MAYOR 1000  -> NO CONCILIA
    #
    # El abs() solamente se utiliza sobre la diferencia para
    # permitir una tolerancia de centavos.
    # ========================================================

    for idx_a, row_a in df_arca.iterrows():

        monto_a = round(
            float(row_a["Importe Ret./Perc."]),
            2
        )

        for idx_m, row_m in df_mayor.iterrows():

            # Si esta partida del Mayor ya fue utilizada,
            # no puede volver a utilizarse.
            if idx_m in mayor_matched_indices:
                continue


            # ------------------------------------------------
            # AHORA SE UTILIZA SALDO Y NO DEBE
            # ------------------------------------------------

            monto_m = round(
                float(row_m["SALDO"]),
                2
            )


            # ------------------------------------------------
            # COMPARACIÓN CONSERVANDO EL SIGNO
            # ------------------------------------------------

            if abs(monto_a - monto_m) < 0.01:

                arca_matched_indices.add(idx_a)
                mayor_matched_indices.add(idx_m)


                conciliadas_1a1.append({

                    "cert_arca": row_a["CERT_STR"],

                    "fecha_arca": str(
                        row_a["Fecha Ret./Perc."]
                    )[:10],

                    "comp_arca": str(
                        row_a.get(
                            "Número Comprobante",
                            ""
                        )
                    ),

                    "razon_arca": row_a[
                        "Denominación o Razón Social"
                    ],

                    "monto_arca": monto_a,

                    "asiento_mayor": row_m[
                        "ASIENTO_STR"
                    ],

                    "fecha_mayor": str(
                        row_m["FECHA"]
                    )[:10],

                    "comp_mayor": str(
                        row_m.get(
                            "DETALLE",
                            ""
                        )
                    ),

                    "razon_mayor": row_m.get(
                        "ENTIDAD",
                        ""
                    ),

                    "monto_mayor": monto_m,
                })

                break


    # ========================================================
    # OBTENER REGISTROS NO CONCILIADOS
    # ========================================================

    unmatched_arca = df_arca[
        ~df_arca.index.isin(arca_matched_indices)
    ]

    unmatched_mayor = df_mayor[
        ~df_mayor.index.isin(mayor_matched_indices)
    ]


    # ========================================================
    # PENDIENTES ARCA
    # ========================================================

    for idx_a, row_a in unmatched_arca.iterrows():

        pendientes_arca.append({

            "cert_arca": row_a["CERT_STR"],

            "fecha_arca": str(
                row_a["Fecha Ret./Perc."]
            )[:10],

            "tipo_comp": str(
                row_a.get(
                    "Descripción Comprobante",
                    ""
                )
            ),

            "comp_arca": str(
                row_a.get(
                    "Número Comprobante",
                    ""
                )
            ),

            "razon_arca": row_a[
                "Denominación o Razón Social"
            ],

            "monto_arca": round(
                float(
                    row_a["Importe Ret./Perc."]
                ),
                2
            ),
        })


    # ========================================================
    # PENDIENTES MAYOR
    #
    # IMPORTANTE:
    # Se muestra SALDO y no DEBE.
    # ========================================================

    for idx_m, row_m in unmatched_mayor.iterrows():

        pendientes_mayor.append({

            "asiento_mayor": row_m[
                "ASIENTO_STR"
            ],

            "fecha_mayor": str(
                row_m["FECHA"]
            )[:10],

            "tipo_comp": "RECIBO / PV",

            "ref_mayor": str(
                row_m.get(
                    "REFERENCIA",
                    ""
                )
            ),

            "razon_mayor": row_m.get(
                "ENTIDAD",
                ""
            ),

            "monto_mayor": round(
                float(
                    row_m["SALDO"]
                ),
                2
            ),
        })


    # ========================================================
    # CREAR ARCHIVO EXCEL
    # ========================================================

    wb = openpyxl.Workbook()

    ws_resumen = wb.active
    ws_resumen.title = "Resumen General"

    ws_1a1 = wb.create_sheet(
        title="Conciliadas 1a1"
    )

    ws_lote = wb.create_sheet(
        title="Conciliadas por Lote"
    )

    ws_arca = wb.create_sheet(
        title="Pendientes_ARCA"
    )

    ws_mayor = wb.create_sheet(
        title="Pendientes_MAYOR"
    )


    # --------------------------------------------------------
    # CONFIGURACIÓN DE HOJAS
    # --------------------------------------------------------

    for ws in wb.worksheets:
        ws.views.sheetView[0].showGridLines = True


    # ========================================================
    # ESTILOS
    # ========================================================

    font_title = Font(
        name="Arial",
        size=14,
        bold=True,
        color="1F4E79"
    )

    font_subtitle = Font(
        name="Arial",
        size=10,
        italic=True,
        color="595959"
    )

    font_header = Font(
        name="Arial",
        size=10,
        bold=True,
        color="FFFFFF"
    )

    font_bold = Font(
        name="Arial",
        size=10,
        bold=True
    )

    font_regular = Font(
        name="Arial",
        size=10
    )


    fill_resumen_hdr = PatternFill(
        start_color="1F4E79",
        end_color="1F4E79",
        fill_type="solid"
    )

    fill_1a1_hdr = PatternFill(
        start_color="2E75B6",
        end_color="2E75B6",
        fill_type="solid"
    )

    fill_lote_hdr = PatternFill(
        start_color="2E75B6",
        end_color="2E75B6",
        fill_type="solid"
    )

    fill_arca_hdr = PatternFill(
        start_color="C65911",
        end_color="C65911",
        fill_type="solid"
    )

    fill_mayor_hdr = PatternFill(
        start_color="595959",
        end_color="595959",
        fill_type="solid"
    )

    fill_total = PatternFill(
        start_color="F2F2F2",
        end_color="F2F2F2",
        fill_type="solid"
    )


    thin_side = Side(
        border_style="thin",
        color="D9D9D9"
    )

    thin_border = Border(
        left=thin_side,
        right=thin_side,
        top=thin_side,
        bottom=thin_side
    )


    double_bottom = Side(
        border_style="double",
        color="000000"
    )

    top_thin = Side(
        border_style="thin",
        color="000000"
    )

    total_border = Border(
        top=top_thin,
        bottom=double_bottom
    )


    align_center = Alignment(
        horizontal="center",
        vertical="center"
    )

    align_left = Alignment(
        horizontal="left",
        vertical="center"
    )

    align_right = Alignment(
        horizontal="right",
        vertical="center"
    )


    fmt_currency = '"$ "#,##0.00'


    # ========================================================
    # HOJA: CONCILIADAS 1 A 1
    # ========================================================

    ws_1a1.cell(
        row=1,
        column=1,
        value="PARTIDAS CONCILIADAS EXACTAS (1 A 1)"
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
        "Saldo Mayor ($)",
        "Diferencia ($)",
    ]


    for col_idx, text in enumerate(
        headers_1a1,
        1
    ):

        c = ws_1a1.cell(
            row=3,
            column=col_idx,
            value=text
        )

        c.font = font_header
        c.fill = fill_1a1_hdr
        c.alignment = align_center


    r_idx = 4


    for item in conciliadas_1a1:

        ws_1a1.cell(
            row=r_idx,
            column=1,
            value=item["cert_arca"]
        ).alignment = align_center

        ws_1a1.cell(
            row=r_idx,
            column=2,
            value=item["fecha_arca"]
        ).alignment = align_center

        ws_1a1.cell(
            row=r_idx,
            column=3,
            value=item["comp_arca"]
        ).alignment = align_center

        ws_1a1.cell(
            row=r_idx,
            column=4,
            value=item["razon_arca"]
        ).alignment = align_left


        c_m1 = ws_1a1.cell(
            row=r_idx,
            column=5,
            value=item["monto_arca"]
        )

        c_m1.number_format = fmt_currency
        c_m1.alignment = align_right


        ws_1a1.cell(
            row=r_idx,
            column=6,
            value=item["asiento_mayor"]
        ).alignment = align_center

        ws_1a1.cell(
            row=r_idx,
            column=7,
            value=item["fecha_mayor"]
        ).alignment = align_center

        ws_1a1.cell(
            row=r_idx,
            column=8,
            value=item["comp_mayor"]
        ).alignment = align_center

        ws_1a1.cell(
            row=r_idx,
            column=9,
            value=item["razon_mayor"]
        ).alignment = align_left


        c_m2 = ws_1a1.cell(
            row=r_idx,
            column=10,
            value=item["monto_mayor"]
        )

        c_m2.number_format = fmt_currency
        c_m2.alignment = align_right


        c_diff = ws_1a1.cell(
            row=r_idx,
            column=11,
            value=f"=E{r_idx}-J{r_idx}"
        )

        c_diff.number_format = fmt_currency
        c_diff.alignment = align_right


        for c in range(1, 12):

            ws_1a1.cell(
                row=r_idx,
                column=c
            ).font = font_regular

            ws_1a1.cell(
                row=r_idx,
                column=c
            ).border = thin_border


        r_idx += 1


    # --------------------------------------------------------
    # TOTAL 1 A 1
    # --------------------------------------------------------

    ws_1a1.cell(
        row=r_idx,
        column=1,
        value="TOTAL CONCILIADO 1 A 1"
    ).font = font_bold

    ws_1a1.cell(
        row=r_idx,
        column=1
    ).alignment = align_left


    # Evitar fórmulas con rangos invertidos cuando no hay datos
    if r_idx > 4:

        rango_inicio = 4
        rango_fin = r_idx - 1

        formula_arca = (
            f"=SUM(E{rango_inicio}:E{rango_fin})"
        )

        formula_mayor = (
            f"=SUM(J{rango_inicio}:J{rango_fin})"
        )

        formula_diff = (
            f"=SUM(K{rango_inicio}:K{rango_fin})"
        )

    else:

        formula_arca = "=0"
        formula_mayor = "=0"
        formula_diff = "=0"


    c_tot1 = ws_1a1.cell(
        row=r_idx,
        column=5,
        value=formula_arca
    )

    c_tot1.font = font_bold
    c_tot1.number_format = fmt_currency


    c_tot2 = ws_1a1.cell(
        row=r_idx,
        column=10,
        value=formula_mayor
    )

    c_tot2.font = font_bold
    c_tot2.number_format = fmt_currency


    c_totdiff = ws_1a1.cell(
        row=r_idx,
        column=11,
        value=formula_diff
    )

    c_totdiff.font = font_bold
    c_totdiff.number_format = fmt_currency


    for c in range(1, 12):

        ws_1a1.cell(
            row=r_idx,
            column=c
        ).fill = fill_total

        ws_1a1.cell(
            row=r_idx,
            column=c
        ).border = total_border


    row_1a1_total = r_idx
    cnt_1a1 = len(conciliadas_1a1)


    # ========================================================
    # HOJA: CONCILIADAS POR LOTE
    #
    # Se conserva la hoja original.
    # La lógica de lote no estaba implementada en el código
    # original, por lo que queda preparada sin inventar
    # agrupaciones.
    # ========================================================

    ws_lote.cell(
        row=1,
        column=1,
        value="PARTIDAS CONCILIADAS POR LOTE (AGRUPADAS)"
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


    for col_idx, text in enumerate(
        headers_lote,
        1
    ):

        c = ws_lote.cell(
            row=4,
            column=col_idx,
            value=text
        )

        c.font = font_header
        c.fill = fill_lote_hdr
        c.alignment = align_center


    # No se generan lotes porque la lógica de agrupación
    # no estaba definida en el código original.

    r_lote_idx = 5


    ws_lote.cell(
        row=r_lote_idx,
        column=1,
        value="TOTAL CONCILIADO POR LOTE"
    ).font = font_bold


    c_lt1 = ws_lote.cell(
        row=r_lote_idx,
        column=4,
        value="=0"
    )

    c_lt1.font = font_bold
    c_lt1.number_format = fmt_currency


    c_lt2 = ws_lote.cell(
        row=r_lote_idx,
        column=7,
        value="=0"
    )

    c_lt2.font = font_bold
    c_lt2.number_format = fmt_currency


    c_ltdiff = ws_lote.cell(
        row=r_lote_idx,
        column=8,
        value="=0"
    )

    c_ltdiff.font = font_bold
    c_ltdiff.number_format = fmt_currency


    for c in range(1, 10):

        ws_lote.cell(
            row=r_lote_idx,
            column=c
        ).fill = fill_total

        ws_lote.cell(
            row=r_lote_idx,
            column=c
        ).border = total_border


    row_lote_total = r_lote_idx
    cnt_lote = len(conciliadas_lote)


    # ========================================================
    # HOJA: PENDIENTES ARCA
    # ========================================================

    ws_arca.cell(
        row=1,
        column=1,
        value="RETENCIONES PENDIENTES EN ARCA (NO REGISTRADAS EN MAYOR)"
    ).font = font_title


    headers_arca = [
        "Nro Certificado",
        "Fecha Reg.",
        "Tipo Comprobante",
        "Nro Comprobante",
        "Razón Social Agente",
        "Monto Retención ($)",
    ]


    for col_idx, text in enumerate(
        headers_arca,
        1
    ):

        c = ws_arca.cell(
            row=3,
            column=col_idx,
            value=text
        )

        c.font = font_header
        c.fill = fill_arca_hdr
        c.alignment = align_center


    r_arca_idx = 4


    for item in pendientes_arca:

        ws_arca.cell(
            row=r_arca_idx,
            column=1,
            value=item["cert_arca"]
        ).alignment = align_center

        ws_arca.cell(
            row=r_arca_idx,
            column=2,
            value=item["fecha_arca"]
        ).alignment = align_center

        ws_arca.cell(
            row=r_arca_idx,
            column=3,
            value=item["tipo_comp"]
        ).alignment = align_center

        ws_arca.cell(
            row=r_arca_idx,
            column=4,
            value=item["comp_arca"]
        ).alignment = align_center

        ws_arca.cell(
            row=r_arca_idx,
            column=5,
            value=item["razon_arca"]
        ).alignment = align_left


        c_m = ws_arca.cell(
            row=r_arca_idx,
            column=6,
            value=item["monto_arca"]
        )

        c_m.number_format = fmt_currency
        c_m.alignment = align_right


        for c in range(1, 7):

            ws_arca.cell(
                row=r_arca_idx,
                column=c
            ).font = font_regular

            ws_arca.cell(
                row=r_arca_idx,
                column=c
            ).border = thin_border


        r_arca_idx += 1


    ws_arca.cell(
        row=r_arca_idx,
        column=1,
        value="TOTAL PENDIENTE ARCA"
    ).font = font_bold


    if r_arca_idx > 4:

        formula_tot_arca = (
            f"=SUM(F4:F{r_arca_idx - 1})"
        )

    else:

        formula_tot_arca = "=0"


    c_totarca = ws_arca.cell(
        row=r_arca_idx,
        column=6,
        value=formula_tot_arca
    )

    c_totarca.font = font_bold
    c_totarca.number_format = fmt_currency
    c_totarca.alignment = align_right


    for c in range(1, 7):

        ws_arca.cell(
            row=r_arca_idx,
            column=c
        ).fill = fill_total

        ws_arca.cell(
            row=r_arca_idx,
            column=c
        ).border = total_border


    row_arca_total = r_arca_idx
    cnt_arca = len(pendientes_arca)


    # ========================================================
    # HOJA: PENDIENTES MAYOR
    # ========================================================

    ws_mayor.cell(
        row=1,
        column=1,
        value="REGISTROS PENDIENTES EN MAYOR (SIN CERTIFICADO EN ARCA)"
    ).font = font_title


    headers_mayor = [
        "Nro Asiento",
        "Fecha Contable",
        "Tipo Comprobante",
        "Referencia",
        "Cuenta / Descripción",
        "Saldo ($)",
    ]


    for col_idx, text in enumerate(
        headers_mayor,
        1
    ):

        c = ws_mayor.cell(
            row=3,
            column=col_idx,
            value=text
        )

        c.font = font_header
        c.fill = fill_mayor_hdr
        c.alignment = align_center


    r_mayor_idx = 4


    for item in pendientes_mayor:

        ws_mayor.cell(
            row=r_mayor_idx,
            column=1,
            value=item["asiento_mayor"]
        ).alignment = align_center

        ws_mayor.cell(
            row=r_mayor_idx,
            column=2,
            value=item["fecha_mayor"]
        ).alignment = align_center

        ws_mayor.cell(
            row=r_mayor_idx,
            column=3,
            value=item["tipo_comp"]
        ).alignment = align_center

        ws_mayor.cell(
            row=r_mayor_idx,
            column=4,
            value=item["ref_mayor"]
        ).alignment = align_center

        ws_mayor.cell(
            row=r_mayor_idx,
            column=5,
            value=item["razon_mayor"]
        ).alignment = align_left


        c_m = ws_mayor.cell(
            row=r_mayor_idx,
            column=6,
            value=item["monto_mayor"]
        )

        c_m.number_format = fmt_currency
        c_m.alignment = align_right


        for c in range(1, 7):

            ws_mayor.cell(
                row=r_mayor_idx,
                column=c
            ).font = font_regular

            ws_mayor.cell(
                row=r_mayor_idx,
                column=c
            ).border = thin_border


        r_mayor_idx += 1


    ws_mayor.cell(
        row=r_mayor_idx,
        column=1,
        value="TOTAL PENDIENTE MAYOR"
    ).font = font_bold


    if r_mayor_idx > 4:

        formula_tot_mayor = (
            f"=SUM(F4:F{r_mayor_idx - 1})"
        )

    else:

        formula_tot_mayor = "=0"


    c_totmayor = ws_mayor.cell(
        row=r_mayor_idx,
        column=6,
        value=formula_tot_mayor
    )

    c_totmayor.font = font_bold
    c_totmayor.number_format = fmt_currency
    c_totmayor.alignment = align_right


    for c in range(1, 7):

        ws_mayor.cell(
            row=r_mayor_idx,
            column=c
        ).fill = fill_total

        ws_mayor.cell(
            row=r_mayor_idx,
            column=c
        ).border = total_border


    row_mayor_total = r_mayor_idx
    cnt_mayor = len(pendientes_mayor)


    # ========================================================
    # HOJA: RESUMEN GENERAL
    # ========================================================

    ws_resumen.cell(
        row=1,
        column=1,
        value="CONCILIACIÓN DE RETENCIONES - RESUMEN CONSOLIDADO"
    ).font = font_title


    headers_resumen = [
        "Categoría / Pestaña",
        "Cantidad de Reg.",
        "Total ARCA ($)",
        "Total Mayor ($)",
        "Diferencia ($)",
    ]


    for col_idx, text in enumerate(
        headers_resumen,
        1
    ):

        c = ws_resumen.cell(
            row=4,
            column=col_idx,
            value=text
        )

        c.font = font_header
        c.fill = fill_resumen_hdr
        c.alignment = align_center


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


    for idx, rdata in enumerate(
        resumen_rows,
        5
    ):

        ws_resumen.cell(
            row=idx,
            column=1,
            value=rdata[0]
        ).alignment = align_left

        ws_resumen.cell(
            row=idx,
            column=2,
            value=rdata[1]
        ).alignment = align_center


        c_arca = ws_resumen.cell(
            row=idx,
            column=3,
            value=rdata[2]
        )

        c_arca.number_format = fmt_currency
        c_arca.alignment = align_right


        c_myr = ws_resumen.cell(
            row=idx,
            column=4,
            value=rdata[3]
        )

        c_myr.number_format = fmt_currency
        c_myr.alignment = align_right


        c_diff = ws_resumen.cell(
            row=idx,
            column=5,
            value=rdata[4]
        )

        c_diff.number_format = fmt_currency
        c_diff.alignment = align_right


        for c in range(1, 6):

            ws_resumen.cell(
                row=idx,
                column=c
            ).font = font_regular

            ws_resumen.cell(
                row=idx,
                column=c
            ).border = thin_border


    # ========================================================
    # TOTAL GENERAL
    # ========================================================

    ws_resumen.cell(
        row=9,
        column=1,
        value="TOTAL GENERAL EMANADO"
    ).font = font_bold


    ws_resumen.cell(
        row=9,
        column=2,
        value="=SUM(B5:B8)"
    ).font = font_bold

    ws_resumen.cell(
        row=9,
        column=2
    ).alignment = align_center


    c_tot_arca = ws_resumen.cell(
        row=9,
        column=3,
        value="=SUM(C5:C8)"
    )

    c_tot_arca.font = font_bold
    c_tot_arca.number_format = fmt_currency
    c_tot_arca.alignment = align_right


    c_tot_myr = ws_resumen.cell(
        row=9,
        column=4,
        value="=SUM(D5:D8)"
    )

    c_tot_myr.font = font_bold
    c_tot_myr.number_format = fmt_currency
    c_tot_myr.alignment = align_right


    c_tot_diff = ws_resumen.cell(
        row=9,
        column=5,
        value="=C9-D9"
    )

    c_tot_diff.font = font_bold
    c_tot_diff.number_format = fmt_currency
    c_tot_diff.alignment = align_right


    for c in range(1, 6):

        ws_resumen.cell(
            row=9,
            column=c
        ).fill = fill_total

        ws_resumen.cell(
            row=9,
            column=c
        ).border = total_border


    # ========================================================
    # AJUSTAR ANCHO DE COLUMNAS
    # ========================================================

    for ws in wb.worksheets:

        for col in ws.columns:

            max_len = 0

            col_letter = get_column_letter(
                col[0].column
            )

            for cell in col:

                if cell.row in [1, 2]:
                    continue

                val_str = str(
                    cell.value or ""
                )

                if (
                    cell.number_format
                    and "$" in cell.number_format
                ):
                    val_str += "   "

                max_len = max(
                    max_len,
                    len(val_str)
                )

            ws.column_dimensions[
                col_letter
            ].width = max(
                max_len + 4,
                12
            )


    # ========================================================
    # GENERAR ARCHIVO EN MEMORIA
    # ========================================================

    output = io.BytesIO()

    wb.save(output)

    output.seek(0)

    return output


# ============================================================
# BOTÓN DE PROCESAMIENTO
# ============================================================

if file_mayor and file_arca:

    if st.button(
        "⚙️ Procesar Conciliación",
        type="primary"
    ):

        try:

            with st.spinner(
                "Procesando datos y armando informe..."
            ):

                excel_bytes = procesar_archivos(
                    file_mayor,
                    file_arca
                )


            st.success(
                "✅ ¡Conciliación completada!"
            )


            st.download_button(
                label="📥 Descargar Archivo Excel Conciliado",
                data=excel_bytes,
                file_name="Conciliacion_Retenciones.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
            )


        except Exception as e:

            st.error(
                f"❌ Se produjo un error al procesar los archivos: {e}"
            )
