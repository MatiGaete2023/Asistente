#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
motor/mapeo_columnas.py — v8.1
Cambios: agrega alias prox_aud para PROXS. AUDS.
         agrega alias para columnas Hoja2 (cruce informes)
"""

from .utilidades import normalizar


def obtener_col(df, aliases):
    for c in df.columns:
        for a in aliases:
            if normalizar(c) == normalizar(a):
                return c
    return None


def mapear_columnas(df, modo):
    cols = {}

    # Comunes a todos los modos
    cols["programa"]    = obtener_col(df, ["DERIVACION", "DERIVACIÓN", "PROGRAMA", "NOMBRE CENTRO"])
    cols["tribunal"]    = obtener_col(df, ["TRIBUNAL"])
    cols["nombre"]      = obtener_col(df, ["NOMBRE", "NOMBRE COMPLETO", "NOMBRE MENOR"])
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
    cols["rit"]         = obtener_col(df, ["RIT"])
    # Columna RUT (para cruce — incluye alias Hoja2)
    cols["rut"]         = obtener_col(df, ["RUT", "RUT MENOR"])

    if modo == "ESPERA":
        cols["espera"] = obtener_col(df, ["T ESPERA", "T_ESPERA", "DIAS_ESPERA", "TESPERA",
                                           "DÍAS DE ESPERA", "DIAS DE ESPERA"])

    elif modo == "CUMPLIMIENTO":
        cols["dias_cumpl"]  = obtener_col(df, ["DIAS DE CUMPLIMIENTO", "DÍAS DE CUMPLIMIENTO",
                                                "DIAS CUMPLIMIENTO"])
        cols["dias_egresar"] = obtener_col(df, ["DIAS PARA EGRESAR", "DÍAS PARA EGRESAR",
                                                 "DIAS EGRESAR"])
        cols["ingreso"]      = obtener_col(df, ["FEC.INGRESO EFECTIVO", "FEC. INGRESO EFECTIVO",
                                                 "FEC INGRESO EFECTIVO"])
        cols["egreso_proy"]  = obtener_col(df, ["FEC.EGRESO PROYECTADO", "FEC. EGRESO PROYECTADO",
                                                 "FEC EGRESO PROYECTADO"])
        cols["ficha_res"]    = obtener_col(df, ["FEC.ACT.F.RESIDENCIAL", "FEC. ACT. F. RESIDENCIAL",
                                                 "FEC ACT F RESIDENCIAL"])
        cols["ficha_ind"]    = obtener_col(df, ["FEC.ACT.F.INDIVIDUAL", "FEC. ACT. F. INDIVIDUAL",
                                                 "FEC ACT F INDIVIDUAL"])
        cols["ficha_amb"]    = obtener_col(df, ["FEC.ACT.F.AMBULATORIA", "FEC. ACT. F. AMBULATORIA",
                                                 "FEC ACT F AMBULATORIA"])
        cols["ficha_fae"]    = obtener_col(df, ["FEC.ACT.F.FAE", "FEC. ACT. F. FAE",
                                                 "FEC ACT F FAE", "FEC.ACT.F.FAE/FAS",
                                                 "FEC. ACT. F. FAE/FAS"])

    elif modo == "INFORMES":
        cols["vencimiento"] = obtener_col(df, ["FECHA VENCIMIENTO", "FEC.VENCIMIENTO",
                                                "FEC. VENCIMIENTO"])
        cols["ingreso"]     = obtener_col(df, ["FEC.INGRESO EFECTIVO", "FEC. INGRESO EFECTIVO",
                                                "FECHA INGRESO"])

    return cols


def mapear_columnas_hoja2(df):
    """
    Mapea columnas de Hoja2 (informes por vencer).
    Retorna dict con claves: rit, rut, nombre, tribunal, programa, vencimiento
    """
    cols = {}
    cols["rit"]         = obtener_col(df, ["RIT"])
    cols["rut"]         = obtener_col(df, ["RUT MENOR", "RUT"])
    cols["nombre"]      = obtener_col(df, ["NOMBRE MENOR", "NOMBRE"])
    cols["tribunal"]    = obtener_col(df, ["TRIBUNAL"])
    cols["programa"]    = obtener_col(df, ["NOMBRE CENTRO", "DERIVACION", "DERIVACIÓN"])
    cols["vencimiento"] = obtener_col(df, ["FECHA VENCIMIENTO", "FEC.VENCIMIENTO",
                                            "FEC. VENCIMIENTO"])
    return cols
