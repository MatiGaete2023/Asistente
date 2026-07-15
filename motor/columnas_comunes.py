# -*- coding: utf-8 -*-
"""Alias y utilidades compartidas entre motor y comunicaciones."""
from .utilidades import normalizar, detectar_tribunal, titulo_programa, es_dce

ALIAS_PROGRAMA = ["DERIVACION", "DERIVACIÓN", "PROGRAMA", "NOMBRE CENTRO"]
ALIAS_TRIBUNAL = ["TRIBUNAL"]
ALIAS_NOMBRE = ["NOMBRE", "NOMBRE COMPLETO", "NOMBRE MENOR"]
ALIAS_RUT = ["RUT", "RUT MENOR", "RUT NNA", "RUT LITIGANTE"]
ALIAS_RIT = ["RIT"]
ALIAS_VENCIMIENTO = ["FECHA VENCIMIENTO", "FEC.VENCIMIENTO", "FEC. VENCIMIENTO"]
ALIAS_ESPERA = ["T ESPERA", "T_ESPERA", "DIAS_ESPERA", "TESPERA", "DÍAS DE ESPERA", "DIAS DE ESPERA"]
TRIBUNAL_DISPLAY = {"LAJA": "Laja", "MULCHEN": "Mulchén", "TOME": "Tomé"}

def obtener_col(df, aliases):
    for c in df.columns:
        for a in aliases:
            if normalizar(c) == normalizar(a):
                return c
    return None
