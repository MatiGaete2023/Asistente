#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
motor/reglas_cumplimiento.py — v8.13
Textos migrados a motor/textos_observaciones.json (fuente unica de verdad).
Logica de disparo IDENTICA a v8.12.1 — cero cambios de comportamiento,
solo el mecanismo de almacenamiento del texto.
"""

from datetime import datetime
from .utilidades import (
    prefijo_observacion, fecha_es, limpiar_nombre, audiencia_suffix,
    normalizar, es_derivacion_sin_seg, tiene_curador_real,
    get_int, get_date, proximo_informe,
    calcular_edad_exacta, dias_para_mayoria, fecha_mayoria
)
from .textos import render

# Prefijos de programas residenciales — usados por R8 (ficha residencial) y R9 (ficha individual).
# Excluye AFT, PIE, PAS, PRM, FAE, DCE (no residenciales).
_PREFIJOS_RESIDENCIAL = ("rta", "rtt", "res", "pee", "rfa", "rppm")


def _es_residencial(programa_norm: str) -> bool:
    """True si el programa normalizado empieza con un prefijo residencial."""
    return programa_norm.startswith(_PREFIJOS_RESIDENCIAL)


def generar_observacion_cumplimiento(row, tribunal, cols):
    obs = []

    programa          = str(row.get(cols.get("programa"),    "")).strip()
    nombre            = str(row.get(cols.get("nombre"),      "")).strip()
    dias_para_egresar = get_int(row.get(cols.get("dias_egresar"), None))
    curador           = str(row.get(cols.get("curador"),     "")).strip()
    oido              = get_date(row.get(cols.get("oido"),   None))
    fec_ingreso       = get_date(row.get(cols.get("ingreso"),None))
    fec_egrso_proy    = get_date(row.get(cols.get("egreso_proy"), None))
    fec_ficha_res     = get_date(row.get(cols.get("ficha_res"),   None))
    fec_ficha_ind     = get_date(row.get(cols.get("ficha_ind"),   None))
    fec_ficha_fae     = get_date(row.get(cols.get("ficha_fae"),   None))
    fec_nacimiento    = get_date(row.get(cols.get("nacimiento"),  None))

    edad_real     = calcular_edad_exacta(fec_nacimiento) or get_int(row.get(cols.get("edad"), 0)) or 0
    _ahora        = datetime.now()
    dias_oido     = (_ahora - oido).days        if oido        else None
    dias_ingreso  = (_ahora - fec_ingreso).days if fec_ingreso else None
    dias_ficha_res = (_ahora - fec_ficha_res).days if fec_ficha_res else None
    dias_ficha_ind = (_ahora - fec_ficha_ind).days if fec_ficha_ind else None

    programa_norm = normalizar(programa)
    pfx = prefijo_observacion(nombre, programa)
    aud = audiencia_suffix(row, cols)

    _nombre_limpio = limpiar_nombre(nombre)
    pnombre = _nombre_limpio.split()[0] if _nombre_limpio else ""

    # R0
    if es_derivacion_sin_seg(programa):
        return f"{pfx}{render('CUMPLIMIENTO', 'R0')}{aud}"

    # R1: Mayor de edad
    if edad_real >= 18:
        fec_may = fecha_mayoria(fec_nacimiento) if fec_nacimiento else None
        if pnombre and fec_may:
            texto = render("CUMPLIMIENTO", "R1_CON_FECHA", PNOMBRE=pnombre, FECHA_MAYORIA=fecha_es(fec_may))
        else:
            texto = render("CUMPLIMIENTO", "R1_FALLBACK")
        if aud:
            texto = texto[:-1]
        return f"{pfx}{texto}{aud}"

    # R2 — acumulable
    dias_may = dias_para_mayoria(fec_nacimiento) if fec_nacimiento else None
    if dias_may is not None and 0 < dias_may <= 60:
        obs.append(render("CUMPLIMIENTO", "R2_PROXIMA_MAYORIA",
                           PNOMBRE=pnombre or "[NOMBRE]",
                           FECHA_MAYORIA=fecha_es(fecha_mayoria(fec_nacimiento))))

    # R3
    if dias_ingreso is not None and 0 < dias_ingreso <= 30:
        obs.append(render("CUMPLIMIENTO", "R3_INGRESO_RECIENTE", FECHA_INGRESO=fecha_es(fec_ingreso)))

    # R4
    vencida = (dias_para_egresar is not None and dias_para_egresar <= 0)
    if vencida:
        obs.append(render("CUMPLIMIENTO", "R4_MEDIDA_VENCIDA", FECHA_EGRESO_PROY=fecha_es(fec_egrso_proy)))

    # R5
    r5 = (dias_para_egresar is not None and 0 < dias_para_egresar <= 45)
    if r5:
        obs.append(render("CUMPLIMIENTO", "R5_POR_VENCER", FECHA_EGRESO_PROY=fecha_es(fec_egrso_proy)))

    # R5b
    if not r5 and not vencida:
        fi = proximo_informe(
            fec_ingreso.date() if fec_ingreso and hasattr(fec_ingreso, "date") else None,
            fec_egrso_proy, tribunal
        )
        if fi:
            obs.append(render("CUMPLIMIENTO", "R5B_PROXIMO_INFORME", FECHA_PROX_INFORME=fecha_es(fi)))

    # R6: columna debe existir Y no contener un RUT real
    if cols.get("curador") and not tiene_curador_real(curador):
        obs.append(render("CUMPLIMIENTO", "R6_CURADOR"))

    # R7
    if dias_oido is not None and 0 < dias_oido <= 45:
        obs.append(render("CUMPLIMIENTO", "R7_OIDO", FECHA_OIDO=fecha_es(oido)))

    # R8: solo residencias (RTA, RTT, RES, PEE, RFA, RPPM)
    if (_es_residencial(programa_norm) and cols.get("ficha_res")
            and dias_ficha_res is not None and dias_ficha_res > 180):
        obs.append(render("CUMPLIMIENTO", "R8_FICHA_RESIDENCIAL"))

    # R9: solo aplica a residencias (mismos prefijos que R8)
    if _es_residencial(programa_norm) and cols.get("ficha_ind"):
        if dias_ficha_ind is not None and dias_ficha_ind > 180:
            obs.append(render("CUMPLIMIENTO", "R9_FICHA_INDIVIDUAL_VIEJA"))
        elif dias_ficha_ind is not None and 0 < dias_ficha_ind <= 30:
            obs.append(render("CUMPLIMIENTO", "R9_FICHA_INDIVIDUAL_RECIENTE", FECHA_FICHA_IND=fecha_es(fec_ficha_ind)))

    # R10
    if "fae" in programa_norm or "fas" in programa_norm:
        if cols.get("ficha_fae") and dias_ingreso is not None and dias_ingreso > 120 and not fec_ficha_fae:
            obs.append(render("CUMPLIMIENTO", "R10_FICHA_FAE", PNOMBRE=pnombre or "[NOMBRE]"))

    if not obs:
        obs.append(render("CUMPLIMIENTO", "FALLBACK"))

    # BUG-02: cada fragmento termina en '.', audiencia_suffix empieza con '. '
    base = " ".join(obs)
    if aud and base.endswith("."):
        base = base[:-1]
    return pfx + base + aud


def r5_aplica_para_fila(row, cols):
    dias = get_int(row.get(cols.get("dias_egresar"), None))
    return dias is not None and 0 < dias <= 45


def medida_vencida_para_fila(row, cols):
    dias = get_int(row.get(cols.get("dias_egresar"), None))
    return dias is not None and dias <= 0
