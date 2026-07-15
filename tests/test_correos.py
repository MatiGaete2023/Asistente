# -*- coding: utf-8 -*-
"""Bordes de correos (Mejoras §1.2/§1.6/§1.7) — sin Outlook, con despachador fake."""
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

from comunicaciones.contactos_programas import CatastroContactos
from comunicaciones.generador_correos import GeneradorCorreos, PATRON_ESPERA
from motor.reglas_espera import generar_observacion_espera
from motor.utilidades import normalizar

CATASTRO = Path(__file__).resolve().parents[1] / "comunicaciones" / "catastro_programas.json"
HOY = datetime.now()


def _programa_real():
    return CatastroContactos(str(CATASTRO)).todos()[0]["nombre"]


def _gen():
    capturados = []
    gen = GeneradorCorreos({}, str(CATASTRO),
                           despachador=lambda b: capturados.append(b))
    return gen, capturados


def _fila_informes(programa, dias):
    return {
        "RIT": "R-1", "TRIBUNAL": "LAJA", "RUT": "1-9", "NOMBRE": "Caso",
        "DERIVACION": programa, "FECHA VENCIMIENTO": HOY + timedelta(days=dias),
    }


# ── Clasificación por días: −1 vencido, 0 y 30 por vencer, 31 nada (1.6/1.7) ──

def test_dia_menos_uno_es_vencido():
    gen, caps = _gen()
    gen.procesar(pd.DataFrame([_fila_informes(_programa_real(), -1)]))
    assert len(caps) == 1 and "vencidos" in caps[0]["asunto"]

def test_dia_cero_es_por_vencer_no_vencido():
    gen, caps = _gen()
    gen.procesar(pd.DataFrame([_fila_informes(_programa_real(), 0)]))
    assert len(caps) == 1 and "por vencer" in caps[0]["asunto"]

def test_dia_30_por_vencer_y_31_sin_correo():
    gen, caps = _gen()
    gen.procesar(pd.DataFrame([_fila_informes(_programa_real(), 30)]))
    assert len(caps) == 1 and "por vencer" in caps[0]["asunto"]
    gen2, caps2 = _gen()
    resultado = gen2.procesar(pd.DataFrame([_fila_informes(_programa_real(), 31)]))
    assert resultado["borradores_creados"] == 0 and not caps2


# ── PATRON_ESPERA calza con las TRES ramas de E-05 (Mejoras §1.2) ─────────────

def test_patron_espera_matchea_las_dos_variantes_e05():
    cols = {k: k for k in ("programa", "tribunal", "nombre", "espera")}
    solo_correo = generar_observacion_espera(
        {"programa": "PIE X", "tribunal": "TOME", "nombre": "ANA SOTO", "espera": 35},
        "TOME", cols)
    proyecto = generar_observacion_espera(
        {"programa": "PIE X", "tribunal": "LAJA", "nombre": "ANA SOTO", "espera": 35},
        "LAJA", cols)
    assert PATRON_ESPERA in normalizar(solo_correo)
    assert PATRON_ESPERA in normalizar(proyecto)


def test_procesar_espera_genera_borrador_desde_observacion_motor():
    programa = _programa_real()
    cols = {k: k for k in ("programa", "tribunal", "nombre", "espera")}
    observacion = generar_observacion_espera(
        {"programa": programa, "tribunal": "LAJA", "nombre": "ANA SOTO", "espera": 40},
        "LAJA", cols)
    df = pd.DataFrame([{
        "RIT": "R-1", "TRIBUNAL": "LAJA", "RUT": "1-9", "NOMBRE": "Ana Soto",
        "DERIVACION": programa, "T ESPERA": 40, "OBSERVACION": observacion,
    }])
    gen, caps = _gen()
    resultado = gen.procesar_espera(df)
    assert resultado["borradores_creados"] == 1
    assert "Lista de espera" in caps[0]["asunto"]


def test_espera_sin_patron_no_genera():
    programa = _programa_real()
    df = pd.DataFrame([{
        "RIT": "R-1", "TRIBUNAL": "LAJA", "RUT": "1-9", "NOMBRE": "Ana Soto",
        "DERIVACION": programa, "T ESPERA": 10,
        "OBSERVACION": "Ana PIE: Medida revisada, a la espera de ingreso "
                       "efectivo al programa PIE X.",
    }])
    gen, caps = _gen()
    resultado = gen.procesar_espera(df)
    assert resultado["borradores_creados"] == 0 and not caps


# ── Ya no existe el correo al tribunal (Mejoras §1.1) ─────────────────────────

def test_no_existe_borrador_a_tribunal():
    gen, caps = _gen()
    resultado = gen.procesar(pd.DataFrame(
        [_fila_informes(_programa_real(), -10) for _ in range(3)]))
    assert resultado["borradores_creados"] == len(caps) == 1  # solo al programa
    assert not any("TRIBUNAL" in d for d in resultado["detalle"])
