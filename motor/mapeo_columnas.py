#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
motor/mapeo_columnas.py — v8.1
Cambios: agrega alias prox_aud para PROXS. AUDS.
         agrega alias para columnas Hoja2 (cruce informes)
"""

from .columnas_comunes import (
    obtener_col, ALIAS_PROGRAMA, ALIAS_TRIBUNAL, ALIAS_NOMBRE,
    ALIAS_RUT, ALIAS_RIT, ALIAS_VENCIMIENTO, ALIAS_ESPERA,
)


def mapear_columnas(df, modo):
    cols = {}

    # Comunes a todos los modos
    cols["programa"]    = obtener_col(df, ALIAS_PROGRAMA)
    cols["tribunal"]    = obtener_col(df, ALIAS_TRIBUNAL)
    cols["nombre"]      = obtener_col(df, ALIAS_NOMBRE)
    cols["nacimiento"]  = obtener_col(df, ["FEC. NACIMIENTO", "FEC.NACIMIENTO", "FECHA NACIMIENTO",
                                            "FEC NACIMIENTO"])
    cols["edad"]        = obtener_col(df, ["EDAD"])
    cols["curador"]     = obtener_col(df, ["CURADOR", "CURADOR AD LITEM", "CURADOR AD-LITEM", "CURADOR AD LITEM.", "CUR. AD LITEM", "CURADOR/A AD LITEM", "CURADOR/A AD-LITEM"])
    cols["oido"]        = obtener_col(df, ["FEC. OIDO", "FEC.OIDO", "FEC OIDO", "FECHAOIDO",
                                            "FEC. OÍDO", "FEC.OÍDO"])
    # FEC. RESOLUCIÓN — requerida por INGRESO_ORDENADO_CON_RESOLUCION (ESPERA)
    cols["resolucion"]  = obtener_col(df, ["FEC. RESOLUCIÓN", "FEC.RESOLUCIÓN", "FEC. RESOLUCION",
                                            "FEC.RESOLUCION", "FECHA RESOLUCION", "FEC RESOLUCION"])
    # PROXS. AUDS. — opcional, presente en algunas hojas
    cols["prox_aud"]    = obtener_col(df, ["PROXS. AUDS.", "PROXS AUDS", "PROX AUD",
                                            "PROXIMA AUDIENCIA", "PRÓXIMA AUDIENCIA",
                                            "PROX. AUD.", "PROX. AUDS."])

    # Columna RIT (para cruce)
    cols["rit"]         = obtener_col(df, ALIAS_RIT)
    # Columna RUT (para cruce — incluye alias Hoja2)
    cols["rut"]         = obtener_col(df, ALIAS_RUT)

    if modo == "ESPERA":
        cols["espera"] = obtener_col(df, ALIAS_ESPERA)

    elif modo == "CUMPLIMIENTO":
        cols["dias_cumpl"]  = obtener_col(df, ["DIAS DE CUMPLIMIENTO", "DÍAS DE CUMPLIMIENTO",
                                                "DIAS CUMPLIMIENTO"])
        cols["dias_egresar"] = obtener_col(df, ["DIAS PARA EGRESAR", "DÍAS PARA EGRESAR",
                                                 "DIAS EGRESAR"])
        cols["ingreso"]      = obtener_col(df, ["FEC.INGRESO EFECTIVO", "FEC. INGRESO EFECTIVO",
                                                 "FEC INGRESO EFECTIVO"])
        cols["egreso_proy"]  = obtener_col(df, ["FEC.EGRESO PROYECTADO", "FEC. EGRESO PROYECTADO",
                                                 "FEC EGRESO PROYECTADO"])
        cols["ficha_fae"]    = obtener_col(df, ["FEC.ACT.F.FAE", "FEC. ACT. F. FAE",
                                                 "FEC ACT F FAE", "FEC.ACT.F.FAE/FAS",
                                                 "FEC. ACT. F. FAE/FAS"])

    elif modo == "INFORMES":
        cols["vencimiento"] = obtener_col(df, ALIAS_VENCIMIENTO)
        cols["ingreso"]     = obtener_col(df, ["FEC.INGRESO EFECTIVO", "FEC. INGRESO EFECTIVO",
                                                "FECHA INGRESO"])

    return cols


def mapear_columnas_hoja2(df):
    """
    Mapea columnas de Hoja2 (informes por vencer).
    Retorna dict con claves: rit, rut, nombre, tribunal, programa, vencimiento
    """
    cols = {}
    cols["rit"]         = obtener_col(df, ALIAS_RIT)
    cols["rut"]         = obtener_col(df, ["RUT MENOR", "RUT"])
    cols["nombre"]      = obtener_col(df, ["NOMBRE MENOR", "NOMBRE"])
    cols["tribunal"]    = obtener_col(df, ALIAS_TRIBUNAL)
    cols["programa"]    = obtener_col(df, ["NOMBRE CENTRO", "DERIVACION", "DERIVACIÓN"])
    cols["vencimiento"] = obtener_col(df, ["FECHA VENCIMIENTO", "FEC.VENCIMIENTO",
                                            "FEC. VENCIMIENTO"])
    return cols
