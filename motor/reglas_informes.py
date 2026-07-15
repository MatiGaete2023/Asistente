#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
motor/reglas_informes.py — v8.13
Textos migrados a motor/textos_observaciones.json. Logica identica a v8.4.
"""

from datetime import datetime
from .utilidades import (
    prefijo_observacion, fecha_es, titulo_programa, audiencia_suffix, get_date
)
from .textos import render


def generar_observacion_informes(row, tribunal, cols):
    programa        = str(row.get(cols.get("programa"), "")).strip()
    nombre          = str(row.get(cols.get("nombre"),   "")).strip()
    fec_vencimiento = get_date(row.get(cols.get("vencimiento"), None))

    if not fec_vencimiento:
        return ""

    pfx      = prefijo_observacion(nombre, programa)
    prog_fmt = titulo_programa(programa)
    aud      = audiencia_suffix(row, cols)
    hoy      = datetime.now().date()
    fv       = fec_vencimiento.date() if hasattr(fec_vencimiento, "date") else fec_vencimiento
    dias     = (fv - hoy).days

    if dias <= 0:
        obs = render("INFORMES", "VENCIDO", PROGRAMA=prog_fmt, FECHA_VENCIMIENTO=fecha_es(fec_vencimiento))
    elif 0 < dias <= 45:
        if "DCE" in str(programa).upper():
            obs = render("INFORMES", "POR_VENCER_DCE", PROGRAMA=prog_fmt, FECHA_VENCIMIENTO=fecha_es(fec_vencimiento))
        else:
            obs = render("INFORMES", "POR_VENCER_GENERAL", PROGRAMA=prog_fmt, FECHA_VENCIMIENTO=fecha_es(fec_vencimiento))
    else:
        return None

    base = obs[:-1] if (aud and obs.endswith(".")) else obs
    return pfx + base + aud
