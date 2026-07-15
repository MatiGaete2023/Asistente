#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
motor/reglas_espera.py — v8.13
Textos migrados a motor/textos_observaciones.json (fuente unica de verdad).
Logica de disparo IDENTICA a v8.12.1, salvo:
  - NUEVO: si FEC. RESOLUCION esta presente, el caso DCE<30d y el fallback
    generico usan INGRESO_ORDENADO_CON_RESOLUCION (texto literal entregado
    por el usuario 2026-06-11) en vez del texto v7.1 de dias transcurridos.
  - Sin FEC. RESOLUCION disponible: se preserva el texto v7.1 como residual.
Esquema de union: fragmentos SIN punto final, unidos con ". ".
"""

from datetime import datetime
from .utilidades import (
    prefijo_observacion, fecha_es, limpiar_nombre, audiencia_suffix,
    es_dce, es_derivacion_sin_seg, tiene_curador_real,
    get_int, get_date, calcular_edad_exacta, dias_para_mayoria, fecha_mayoria
)
from .textos import render


def generar_observacion_espera(row, tribunal, cols):
    obs = []

    programa    = str(row.get(cols.get("programa"), "")).strip()
    nombre      = str(row.get(cols.get("nombre"),   "")).strip()
    curador     = str(row.get(cols.get("curador"),  "")).strip()
    dias_espera = get_int(row.get(cols.get("espera"), 0)) or 0
    oido        = get_date(row.get(cols.get("oido"), None))
    fec_nacim   = get_date(row.get(cols.get("nacimiento"), None))
    fec_resol   = get_date(row.get(cols.get("resolucion"), None))

    edad_real = calcular_edad_exacta(fec_nacim) or get_int(row.get(cols.get("edad"), 0)) or 0
    dias_oido = (datetime.now() - oido).days if oido else None
    pfx       = prefijo_observacion(nombre, programa)
    aud       = audiencia_suffix(row, cols)

    _nombre_limpio = limpiar_nombre(nombre)
    pnombre = _nombre_limpio.split()[0] if _nombre_limpio else ""

    # R0
    if es_derivacion_sin_seg(programa):
        return f"{pfx}{render('ESPERA', 'R0')}{aud}"

    # R1: mayor de edad (corte)
    if edad_real >= 18:
        fec_may = fecha_mayoria(fec_nacim) if fec_nacim else None
        if pnombre and fec_may:
            texto = render("ESPERA", "R1_CON_FECHA", PNOMBRE=pnombre, FECHA_MAYORIA=fecha_es(fec_may))
        else:
            texto = render("ESPERA", "R1_FALLBACK")
        if aud:
            texto = texto[:-1]
        return f"{pfx}{texto}{aud}"

    # R2: proximo a mayoria <=60 dias — acumulable
    dias_may = dias_para_mayoria(fec_nacim) if fec_nacim else None
    if dias_may is not None and 0 < dias_may <= 60:
        obs.append(render("ESPERA", "R2_PROXIMA_MAYORIA",
                           PNOMBRE=pnombre or "[NOMBRE]",
                           FECHA_MAYORIA=fecha_es(fecha_mayoria(fec_nacim))))

    # R3/R4: DCE
    if es_dce(programa):
        if 0 <= dias_espera < 30:
            if fec_resol:
                obs.append(render("ESPERA", "INGRESO_ORDENADO_CON_RESOLUCION",
                                   PROGRAMA=programa, FECHA_RESOLUCION=fecha_es(fec_resol)))
            else:
                obs.append(render("ESPERA", "R3_DCE_CORTO_SIN_RESOLUCION", DIAS_ESPERA=dias_espera))
        elif dias_espera >= 30:
            obs.append(render("ESPERA", "R4_DCE_LARGO"))
    else:
        # R5: Mulchen >=30 dias
        if tribunal == "MULCHEN" and dias_espera >= 30:
            obs.append(render("ESPERA", "R5_MULCHEN"))
        # R6: Laja/Tome >=60 dias
        elif tribunal in ("LAJA", "TOME") and dias_espera >= 60:
            obs.append(render("ESPERA", "R6_LAJA_TOME"))

    # R7: curador — columna existe Y sin RUT real
    if cols.get("curador") and not tiene_curador_real(curador):
        obs.append(render("ESPERA", "R7_CURADOR"))

    # R8: oido reciente <=45 dias
    if dias_oido is not None and 0 < dias_oido <= 45:
        obs.append(render("ESPERA", "R8_OIDO"))

    # Fallback — usa FEC. RESOLUCION si esta disponible, si no texto v7.1 residual
    if not obs:
        if fec_resol:
            obs.append(render("ESPERA", "INGRESO_ORDENADO_CON_RESOLUCION",
                               PROGRAMA=programa, FECHA_RESOLUCION=fecha_es(fec_resol)))
        else:
            obs.append(render("ESPERA", "FALLBACK_SIN_RESOLUCION", PROGRAMA=programa))

    return pfx + ". ".join(obs) + aud
