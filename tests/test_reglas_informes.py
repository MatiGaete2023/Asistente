# -*- coding: utf-8 -*-
"""Matriz de bordes de INFORMES — Catálogo v2 §6/§13 y Plan Apéndice C.3."""
from datetime import datetime, timedelta

from motor.composicion import Incidencias
from motor.reglas_informes import generar_observacion_informes
from motor.utilidades import fecha_es, titulo_programa

HOY = datetime.now()

COLS = {
    "programa": "programa", "tribunal": "tribunal", "nombre": "nombre",
    "vencimiento": "vencimiento", "prox_aud": "prox_aud",
    "curador": "curador", "oido": "oido", "rit": "rit",
}


def d(dias_atras):
    return (HOY - timedelta(days=dias_atras)).strftime("%d/%m/%Y")


def f(dias_adelante):
    return (HOY + timedelta(days=dias_adelante)).strftime("%d/%m/%Y")


def fila(**kw):
    base = {
        "programa": "PIE Sintético", "tribunal": "LAJA", "nombre": "MARTIN ROJAS",
        "vencimiento": f(10), "prox_aud": "", "curador": "", "oido": "",
        "rit": "I-1-2026",
    }
    base.update(kw)
    return base


def obs(row, incidencias=None):
    return generar_observacion_informes(row, "LAJA", COLS,
                                        incidencias=incidencias, fila_excel=2)


# ── I-01 / I-02: bordes −1 / 0 / 30 / 31 ─────────────────────────────────────

def test_bordes_menos1_0_30_31():
    assert "se encuentra vencido en RUS desde el" in obs(fila(vencimiento=d(1)))
    o0 = obs(fila(vencimiento=f(0)))
    assert "vence el" in o0 and "vencido" not in o0  # día 0 = por vencer (I-02)
    assert "vence el" in obs(fila(vencimiento=f(30)))
    assert obs(fila(vencimiento=f(31))) == ""


def test_dce_informe_diagnostico_en_ambas_ramas():
    o_venc = obs(fila(programa="DCE Diagnóstico", vencimiento=d(5)))
    assert "informe diagnóstico" in o_venc and "informe de avance" not in o_venc
    o_prox = obs(fila(programa="DCE Diagnóstico", vencimiento=f(5)))
    assert "informe diagnóstico" in o_prox
    assert "ordenado en autos debe ser remitido a más tardar" in o_prox


def test_e01_corta_en_informes():
    o = obs(fila(programa="OPD Concepción", vencimiento=d(10)))
    assert "no se encuentra sujeta a seguimiento" in o
    assert "correo electrónico" not in o


def test_sin_fecha_vencimiento_vacio_mas_incidencia():
    inc = Incidencias()
    assert obs(fila(vencimiento=""), incidencias=inc) == ""
    assert len(inc) == 1


def test_audiencia_se_agrega_pero_no_curador_ni_oido():
    o = obs(fila(vencimiento=d(3), prox_aud=f(7), curador="NO POSEE", oido=d(2)))
    assert "Se cita a audiencia" in o
    assert "curador" not in o and "Oído" not in o


def test_ejemplo_10_6_dce_vencido():
    venc = HOY - timedelta(days=15)
    row = fila(programa="DCE Tomé", vencimiento=d(15))
    o = obs(row)
    prog_fmt = titulo_programa("DCE Tomé")
    assert o == (f"Martin DCE: Se remite correo electrónico al programa {prog_fmt} "
                 f"a fin de requerir el informe diagnóstico que se encuentra "
                 f"vencido en RUS desde el {fecha_es(venc)}.")


def test_sin_marcadores_ni_dobles_puntos():
    for row in (fila(vencimiento=d(1)), fila(vencimiento=f(0), prox_aud=f(2)),
                fila(programa="DCE X", vencimiento=f(20))):
        o = obs(row)
        assert "{" not in o and ".." not in o
        assert o.endswith(".")
