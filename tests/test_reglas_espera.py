# -*- coding: utf-8 -*-
"""Matriz de bordes de ESPERA — Catálogo v2 §3/§9.1/§13 y Plan Apéndice C.1."""
from datetime import datetime, timedelta

from motor.composicion import Incidencias
from motor.reglas_espera import generar_observacion_espera
from motor.utilidades import fecha_es

HOY = datetime.now()

COLS = {
    "programa": "programa", "tribunal": "tribunal", "nombre": "nombre",
    "nacimiento": "nacimiento", "curador": "curador", "oido": "oido",
    "resolucion": "resolucion", "prox_aud": "prox_aud", "espera": "espera",
    "rit": "rit",
}

# Sin la "S" inicial: dentro de E05_PROYECTO_Y_CORREO la frase va en
# minúscula ("Igualmente, se remite correo…").
TXT_CORREO = ("e remite correo electrónico al programa consultando respecto "
              "de la fecha estimada de ingreso efectivo.")
TXT_PROYECTO = ("Se remite proyecto de resolución pidiendo cuenta al programa "
                "respecto del ingreso efectivo.")


def d(dias_atras):
    return (HOY - timedelta(days=dias_atras)).strftime("%d/%m/%Y")


def f(dias_adelante):
    return (HOY + timedelta(days=dias_adelante)).strftime("%d/%m/%Y")


def fila(**kw):
    base = {
        "programa": "PIE Sintético", "tribunal": "LAJA", "nombre": "CAMILA PRUEBA",
        "nacimiento": d(10 * 365), "curador": "11111111-1", "oido": "",
        "resolucion": "", "prox_aud": "", "espera": 10, "rit": "P-1-2026",
    }
    base.update(kw)
    return base


def obs(row, tribunal="LAJA", incidencias=None):
    return generar_observacion_espera(row, tribunal, COLS,
                                      incidencias=incidencias, fila_excel=2)


# ── E-05 tramos por tribunal (§13: 29/30 Laja-Mulchén; 29/30/59/60 Tomé) ──────

def test_e05_laja_29_no_dispara():
    assert TXT_PROYECTO not in obs(fila(espera=29), "LAJA")

def test_e05_laja_30_proyecto_y_correo():
    o = obs(fila(espera=30), "LAJA")
    assert TXT_PROYECTO in o and TXT_CORREO in o

def test_e05_mulchen_29_no_30_si():
    assert TXT_PROYECTO not in obs(fila(espera=29), "MULCHEN")
    o = obs(fila(espera=30), "MULCHEN")
    assert TXT_PROYECTO in o and TXT_CORREO in o

def test_e05_tome_tramos_30_59_60():
    assert TXT_CORREO not in obs(fila(espera=29), "TOME")
    o30 = obs(fila(espera=30), "TOME")
    assert TXT_CORREO in o30 and TXT_PROYECTO not in o30
    o59 = obs(fila(espera=59), "TOME")
    assert TXT_CORREO in o59 and TXT_PROYECTO not in o59
    o60 = obs(fila(espera=60), "TOME")
    assert TXT_PROYECTO in o60 and TXT_CORREO in o60


# ── E-05 rama DCE: nunca proyecto, cualquier tribunal (Manual/§3) ─────────────

def test_e05_dce_nunca_proyecto():
    for trib in ("LAJA", "MULCHEN", "TOME"):
        o = obs(fila(programa="DCE Diagnóstico", espera=40), trib)
        assert TXT_CORREO in o, trib
        assert "proyecto de resolución" not in o, trib
    assert TXT_CORREO not in obs(fila(programa="DCE Diagnóstico", espera=29), "LAJA")


# ── E-04 (0/29/30) y co-ocurrencia E-04+E-05 ─────────────────────────────────

def test_e04_bordes():
    assert "el Tribunal ordenó el ingreso efectivo" in obs(fila(resolucion=d(0)))
    assert "el Tribunal ordenó el ingreso efectivo" in obs(fila(resolucion=d(29)))
    assert "el Tribunal ordenó el ingreso efectivo" not in obs(fila(resolucion=d(30)))

def test_e04_y_e05_se_agregan_ambas_en_orden():
    o = obs(fila(resolucion=d(10), espera=35), "LAJA")
    i_e04 = o.index("el Tribunal ordenó el ingreso efectivo")
    i_e05 = o.index(TXT_PROYECTO)
    assert i_e04 < i_e05
    # E-06 no debe aparecer cuando hay principal
    assert "a la espera de ingreso efectivo al programa" not in o


# ── E-06 fallback (§9.1 / D11) ────────────────────────────────────────────────

def test_e06_espera_corta_sin_resolucion():
    assert "Medida revisada, a la espera de ingreso efectivo al programa" in obs(fila(espera=29))

def test_e06_fallback_con_tribunal_desconocido_y_espera_larga():
    inc = Incidencias()
    o = obs(fila(espera=40), tribunal=None, incidencias=inc)
    assert "a la espera de ingreso efectivo al programa" in o
    assert "proyecto" not in o
    assert len(inc) == 1  # G-06: regla E-05 omitida queda registrada


# ── E-01 corte anclado (§3 corrección D1) ─────────────────────────────────────

def test_e01_corta_todo_incluida_audiencia():
    o = obs(fila(programa="OPD Concepción", prox_aud=f(10), curador=""))
    assert "no se encuentra sujeta a seguimiento" in o
    assert "audiencia" not in o and "curador" not in o

def test_e01_match_anclado_no_subcadena():
    o = obs(fila(programa="PRODAM Sintético", espera=5))
    assert "no se encuentra sujeta a seguimiento" not in o

def test_e01_lista_completa_catalogo():
    for prog in ("OPD Laja", "DAM Centro", "Salud Privada X", "Hospital Regional",
                 "Unidad de Salud Mental Sur", "CESFAM Norte", "Red Salud Bío Bío",
                 "Consulta Externa A", "Colegio San Juan", "Chile Crece Contigo B"):
        assert "no se encuentra sujeta a seguimiento" in obs(fila(programa=prog)), prog


# ── E-02 / E-03: mayoría de edad ──────────────────────────────────────────────

def test_e02_mayor_de_edad_acumula_curador_oido_audiencia():
    o = obs(fila(nacimiento=d(18 * 366), curador="Sin designar",
                 oido=d(5), prox_aud=f(3)))
    assert "alcanzó la mayoría de edad" in o
    assert "curador ad litem" in o
    assert "Oído con fecha" in o
    assert "Se cita a audiencia" in o
    # No se agregan las principales de espera
    assert "a la espera de ingreso efectivo" not in o

def test_e03_bordes_1_60():
    def nac_para(dias_hasta_18):
        return (HOY + timedelta(days=dias_hasta_18)
                - timedelta(days=int(18 * 365.25) + 1)).strftime("%d/%m/%Y")
    assert "alcanzará la mayoría de edad" in obs(fila(nacimiento=nac_para(30)))
    assert "alcanzará la mayoría de edad" not in obs(fila(nacimiento=nac_para(90)))


# ── T-01 / T-02 / T-03 ────────────────────────────────────────────────────────

def test_t01_curador():
    assert "curador ad litem" in obs(fila(curador="NO POSEE"))
    assert "curador ad litem" not in obs(fila(curador="12345678-9 Juan Pérez"))
    assert "curador ad litem" not in obs(fila(curador="(0-0) Institución: CAJ Biobío"))

def test_t02_oido_bordes_0_45_46():
    assert f"Oído con fecha {fecha_es(HOY)}" in obs(fila(oido=d(0)))
    assert "Oído con fecha" in obs(fila(oido=d(45)))
    assert "Oído con fecha" not in obs(fila(oido=d(46)))
    assert "Oído con fecha" not in obs(fila(oido=f(3)))  # nunca futura

def test_t03_audiencia_hoy_si_ayer_no():
    assert "Se cita a audiencia" in obs(fila(prox_aud=f(0)))
    assert "Se cita a audiencia" in obs(fila(prox_aud=f(300)))
    assert "Se cita a audiencia" not in obs(fila(prox_aud=d(1)))


# ── Ejemplos de composición del catálogo (§10) — byte a byte ─────────────────

def test_ejemplo_10_1_tome_35_dias():
    o = obs(fila(nombre="CAMILA SOTO", programa="PIE Renacer", espera=35), "TOME")
    assert o == ("Camila PIE: Medida revisada, a la espera de ingreso efectivo. "
                 "Se remite correo electrónico al programa consultando respecto "
                 "de la fecha estimada de ingreso efectivo.")

def test_ejemplo_10_2_tome_65_dias_con_curador_y_oido():
    o = obs(fila(nombre="CAMILA SOTO", programa="PIE Renacer", espera=65,
                 curador="---", oido=d(7)), "TOME")
    assert o == ("Camila PIE: Medida revisada, a la espera de ingreso efectivo. "
                 + TXT_PROYECTO + " Igualmente, se remite correo electrónico al "
                 "programa consultando respecto de la fecha estimada de ingreso "
                 "efectivo. No registra curador asociado en RUS, se sugiere "
                 "asociar curador ad litem informáticamente. "
                 f"Oído con fecha {fecha_es(HOY - timedelta(days=7))}.")

def test_ejemplo_10_3_dce_40_dias():
    o = obs(fila(nombre="MARTIN ROJAS", programa="DCE Diagnóstico", espera=40), "LAJA")
    assert o == ("Martin DCE: Medida revisada, a la espera de ingreso efectivo. "
                 "Se remite correo electrónico al programa consultando respecto "
                 "de la fecha estimada de ingreso efectivo.")


# ── Invariantes G-04/G-05 ─────────────────────────────────────────────────────

def test_sin_marcadores_ni_dobles_puntos():
    for row in (fila(), fila(espera=40), fila(programa="OPD X"),
                fila(nacimiento=d(19 * 366)), fila(resolucion=d(5), espera=50)):
        o = obs(row)
        assert "{" not in o and "}" not in o
        assert ".." not in o
        assert o.endswith(".")
