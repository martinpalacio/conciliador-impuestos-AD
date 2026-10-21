"""Configuración central: mapeo de columnas por ERP y constantes globales."""

ERP_CONFIGS = {
    "AutoDealer": {
        "asiento": "ASIENTO",
        "fecha": "FECHA",
        "detalle": "DETALLE",
        "referencia": "REFERENCIA",
        "entidad": "ENTIDAD",
        "cuit": "CUIT",
        "debe": "DEBE",
        "haber": "HABER",
        "saldo": "SALDO",
    },
    "Bejerman": {
        "asiento": "AsientoNumero",
        "fecha": "Fecha",
        "detalle": "Concepto",
        "referencia": "CodigoCliProv",
        "ref_secundaria": "NroCompFlex",
        "entidad": "RazonSocial",
        "cuit": "CliProvNroDocumento",
        "debe": "ImpDebe_Loc",
        "haber": "ImpHaber_Loc",
        "saldo": "ImpTotal_Loc",
    },
}

ERP_DEFAULT = "AutoDealer"

SIN_RAZON_SOCIAL = "NO ENCONTRADO EN PADRÓN"

# Tolerancias del motor de cruce
TOLERANCIA_1A1 = 0.01   # diferencia máxima para considerar dos montos iguales
TOLERANCIA_LOTE = 0.50  # diferencia máxima entre sumas de un lote

MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
