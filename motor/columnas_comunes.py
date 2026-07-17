# -*- coding: utf-8 -*-
"""Alias y utilidades compartidas entre motor y comunicaciones."""
from .utilidades import normalizar, detectar_tribunal, titulo_programa, es_dce

# Re-exports deliberados (Mejoras §1.9): comunicaciones/ importa las funciones
# de formato/detección desde este módulo, no desde utilidades directamente.
__all__ = [
    "normalizar", "detectar_tribunal", "titulo_programa", "es_dce",
    "obtener_col", "TRIBUNAL_DISPLAY",
    "ALIAS_PROGRAMA", "ALIAS_TRIBUNAL", "ALIAS_NOMBRE", "ALIAS_RUT",
    "ALIAS_RIT", "ALIAS_VENCIMIENTO", "ALIAS_ESPERA", "ALIAS_NACIMIENTO",
    "ALIAS_RESOLUCION", "ALIAS_DIAS_CUMPLIMIENTO", "ALIAS_DIAS_EGRESO",
    "ALIAS_INGRESO_EFECTIVO", "ALIAS_EGRESO_PROYECTADO",
    "ALIAS_FICHA_FAE", "ALIAS_FICHA_INDIVIDUAL",
]

ALIAS_PROGRAMA = ["DERIVACION", "DERIVACIÓN", "PROGRAMA", "NOMBRE CENTRO"]
ALIAS_TRIBUNAL = ["TRIBUNAL"]
ALIAS_NOMBRE = ["NOMBRE", "NOMBRE COMPLETO", "NOMBRE MENOR"]
ALIAS_RUT = ["RUT", "RUT MENOR", "RUT NNA", "RUT LITIGANTE"]
# Solo "RIT": ampliar alias en el motor fue evaluado y RECHAZADO (Mejoras §2.5)
ALIAS_RIT = ["RIT"]
ALIAS_VENCIMIENTO = ["FECHA VENCIMIENTO", "FEC.VENCIMIENTO", "FEC. VENCIMIENTO"]
ALIAS_ESPERA = ["T ESPERA", "T_ESPERA", "DIAS_ESPERA", "TESPERA",
                "DÍAS DE ESPERA", "DIAS DE ESPERA"]
ALIAS_NACIMIENTO = ["FEC. NACIMIENTO", "FEC.NACIMIENTO", "FECHA NACIMIENTO", "FEC NACIMIENTO"]
ALIAS_RESOLUCION = ["FEC. RESOLUCIÓN", "FEC.RESOLUCIÓN", "FEC. RESOLUCION",
                    "FEC.RESOLUCION", "FECHA RESOLUCION", "FEC RESOLUCION"]
ALIAS_DIAS_CUMPLIMIENTO = ["DIAS DE CUMPLIMIENTO", "DÍAS DE CUMPLIMIENTO",
                           "DIAS CUMPLIMIENTO"]
ALIAS_DIAS_EGRESO = ["DIAS PARA EGRESAR", "DÍAS PARA EGRESAR", "DIAS EGRESAR"]
ALIAS_INGRESO_EFECTIVO = ["FEC.INGRESO EFECTIVO", "FEC. INGRESO EFECTIVO",
                          "FEC INGRESO EFECTIVO"]
ALIAS_EGRESO_PROYECTADO = ["FEC.EGRESO PROYECTADO", "FEC. EGRESO PROYECTADO",
                           "FEC EGRESO PROYECTADO"]
ALIAS_FICHA_FAE = ["FEC.ACT.F.FAE", "FEC. ACT. F. FAE", "FEC ACT F FAE",
                   "FEC.ACT.F.FAE/FAS", "FEC. ACT. F. FAE/FAS"]
ALIAS_FICHA_INDIVIDUAL = ["FEC.ACT.F.INDIVIDUAL", "FEC. ACT. F. INDIVIDUAL",
                          "FEC ACT F INDIVIDUAL"]
# Denominaciones institucionales para comunicaciones dirigidas a programas.
# No usar las claves abreviadas de RUS en textos que salen del sistema.
TRIBUNAL_DISPLAY = {
    "LAJA": "Jgdo. L. y G. de Laja",
    "MULCHEN": "Jgdo. L. y G. de Mulchen",
    "TOME": "Juzgado de Familia Tomé",
}

def obtener_col(df, aliases):
    for c in df.columns:
        for a in aliases:
            if normalizar(c) == normalizar(a):
                return c
    return None
