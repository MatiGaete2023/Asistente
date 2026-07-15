# -*- coding: utf-8 -*-
"""Matriz de bordes de CUMPLIMIENTO — Catálogo v2 §5/§9.2/§13 y Plan Apéndice C.2."""
from datetime import datetime, timedelta

from motor.composicion import Incidencias
from motor.reglas_cumplimiento import generar_observacion_cumplimiento
from motor.utilidades import fecha_es

HOY = datetime.now()

COLS = {
    "programa": "programa", "tribunal": "tribunal", "nombre": "nombre",
    "nacimiento": "nacimiento", "curador": "curador", "oido": "oido",
    "prox_aud": "prox_aud", "rit": "rit", "dias_cumpl": "dias_cumpl",
    "dias_egresar": "dias_egresar", "ingreso": "ingreso",
    "egreso_proy": "egreso_proy", "ficha_ind": "ficha_ind",
    "ficha_fae": "ficha_fae",
}


def d(dias_atras):
    return (HOY - timedelta(days=dias_atras)).strftime("%d/%m/%Y")


def f(dias_adelante):
    return (HOY + timedelta(days=dias_adelante)).strftime("%d/%m/%Y")


def fila(**kw):
    base = {
        "programa": "PIE Sintético", "tribunal": "LAJA", "nombre": "DANIELA PAZ",
        "nacimiento": d(10 * 365), "curador": "11111111-1", "oido": "",
        "prox_aud": "", "rit": "C-1-2026", "dias_cumpl": 200,
        "dias_egresar": 200, "ingreso": d(200), "egreso_proy": f(200),
        "ficha_ind": "", "ficha_fae": "",
    }
    base.update(kw)
    return base


def obs(row, fecha_hoja2=None, incidencias=None):
    return generar_observacion_cumplimiento(
        row, "LAJA", COLS, fecha_hoja2=fecha_hoja2,
        incidencias=incidencias, fila_excel=2)


# ── C-03: condición por columna DIAS DE CUMPLIMIENTO (0/30/31) ───────────────

def test_c03_bordes_0_30_31():
    assert "el ingreso efectivo al programa" in obs(fila(dias_cumpl=0))
    assert "el ingreso efectivo al programa" in obs(fila(dias_cumpl=30))
    assert "el ingreso efectivo al programa" not in obs(fila(dias_cumpl=31))

def test_c03_manda_la_columna_no_el_calculo_de_fechas():
    # días=15 pero la fecha de ingreso es de hace 90 días: dispara igual
    o = obs(fila(dias_cumpl=15, ingreso=d(90)))
    assert "el ingreso efectivo al programa" in o
    assert fecha_es(HOY - timedelta(days=90)) in o

def test_c03_sin_fecha_ingreso_se_omite_con_incidencia():
    inc = Incidencias()
    o = obs(fila(dias_cumpl=10, ingreso=""), incidencias=inc)
    assert "el ingreso efectivo al programa" not in o
    assert any(i["REGLA"] == "C-03" for i in inc.como_dataframe().to_dict("records"))


# ── C-04 / C-05: vencida, vence hoy, por vencer (−1/0/1/45/46) ───────────────

def test_c04_vencida_por_dias_egresar_negativo():
    o = obs(fila(dias_egresar=-1, egreso_proy=d(1)))
    assert "se visualiza vencida en RUS desde el" in o

def test_c04_vencida_por_dias_cumplimiento_negativo():
    o = obs(fila(dias_cumpl=-1, egreso_proy=d(1)))
    assert "se visualiza vencida en RUS desde el" in o

def test_c04_sin_fecha_egreso_no_inventa_fecha():
    inc = Incidencias()
    o = obs(fila(dias_egresar=-5, egreso_proy=""), incidencias=inc)
    assert "vencida" not in o
    assert fecha_es(HOY) not in o  # jamás la fecha de hoy inventada (H-04)
    assert any(i["REGLA"] == "C-04" for i in inc.como_dataframe().to_dict("records"))

def test_c05_dia_cero_texto_hoy():
    o = obs(fila(dias_egresar=0, egreso_proy=f(0)))
    assert "con vencimiento para el día de hoy" in o
    assert "vencida" not in o

def test_c05_bordes_1_45_46():
    assert "próxima a vencer en RUS" in obs(fila(dias_egresar=1, egreso_proy=f(1)))
    assert "próxima a vencer en RUS" in obs(fila(dias_egresar=45, egreso_proy=f(45)))
    assert "próxima a vencer en RUS" not in obs(fila(dias_egresar=46, egreso_proy=f(46)))


# ── C-10: Hoja2 — supresión con C-04/C-05, convivencia con C-03 ──────────────

FH2 = (HOY + timedelta(days=15)).date()

def test_c10_sola_dispara():
    o = obs(fila(), fecha_hoja2=FH2)
    assert "deberá remitir informe de avance a más tardar el" in o

def test_c10_suprimida_con_c04():
    o = obs(fila(dias_egresar=-1, egreso_proy=d(1)), fecha_hoja2=FH2)
    assert "deberá remitir informe" not in o
    assert "vencida" in o

def test_c10_suprimida_con_c05():
    o = obs(fila(dias_egresar=10, egreso_proy=f(10)), fecha_hoja2=FH2)
    assert "deberá remitir informe" not in o
    assert "próxima a vencer" in o

def test_c10_convive_con_c03():
    o = obs(fila(dias_cumpl=5, ingreso=d(5)), fecha_hoja2=FH2)
    assert "el ingreso efectivo al programa" in o
    assert "deberá remitir informe de avance" in o


# ── C-07: ficha individual fusionada, prefijos residenciales ─────────────────

def test_c07_tres_ramas_y_bordes():
    assert "No registra ficha individual" in obs(fila(programa="RTA Hogar", ficha_ind=""))
    assert "superando los 180 días" in obs(fila(programa="RTA Hogar", ficha_ind=d(181)))
    o180 = obs(fila(programa="RTA Hogar", ficha_ind=d(180)))
    assert "ficha individual" not in o180
    o31 = obs(fila(programa="RTA Hogar", ficha_ind=d(31)))
    assert "ficha individual" not in o31
    assert "fue actualizada con fecha" in obs(fila(programa="RTA Hogar", ficha_ind=d(30)))
    assert "fue actualizada con fecha" in obs(fila(programa="RTA Hogar", ficha_ind=d(0)))

def test_c07_prefijos_exclusivos_catalogo():
    # RVA entra; PEE y RPPM salen (catálogo §13); PIE nunca; RESIDENCIA cuenta como RES
    assert "ficha individual" in obs(fila(programa="RVA Amanecer", ficha_ind=""))
    assert "ficha individual" in obs(fila(programa="RESIDENCIA Hogar San Pablo", ficha_ind=""))
    assert "ficha individual" not in obs(fila(programa="PEE Sintético", ficha_ind=""))
    assert "ficha individual" not in obs(fila(programa="RPPM Renacer", ficha_ind=""))
    assert "ficha individual" not in obs(fila(programa="PIE Sintético", ficha_ind=""))


# ── C-08: ficha FAE por token (falso positivo "Rafael" — H-03) ───────────────

def test_c08_fae_bordes_120_121():
    assert "no tiene ficha FAE" in obs(fila(programa="FAE Familia Sur", ingreso=d(121)))
    assert "no tiene ficha FAE" not in obs(fila(programa="FAE Familia Sur", ingreso=d(120)))
    assert "no tiene ficha FAE" not in obs(
        fila(programa="FAE Familia Sur", ingreso=d(121), ficha_fae=d(10)))

def test_c08_fas_tambien_y_rafael_no():
    assert "no tiene ficha FAE" in obs(fila(programa="FAS Norte", ingreso=d(150)))
    assert "no tiene ficha FAE" not in obs(
        fila(programa="Residencia San Rafael", ingreso=d(150)))


# ── C-01 / C-09 y composición ─────────────────────────────────────────────────

def test_c01_mayoria_acumula_curador_oido_audiencia():
    o = obs(fila(nacimiento=d(19 * 366), curador="---", oido=d(3), prox_aud=f(5)))
    assert "alcanzó la mayoría de edad" in o
    assert "curador ad litem" in o and "Oído con fecha" in o
    assert "Se cita a audiencia" in o
    assert "Medida revisada" not in o

def test_c09_sin_reglas_cierre_completo():
    assert obs(fila()) == "Daniela PIE: Medida revisada, sin observaciones."

def test_c09_solo_complementarias_base_breve():
    o = obs(fila(curador="NO POSEE"))
    assert o == ("Daniela PIE: Medida revisada. No registra curador asociado "
                 "en RUS, se sugiere asociar curador ad litem informáticamente.")

def test_c09_solo_c02_usa_base_breve():
    nac = (HOY + timedelta(days=30) - timedelta(days=int(18 * 365.25) + 1)).strftime("%d/%m/%Y")
    o = obs(fila(nacimiento=nac))
    assert o.startswith("Daniela PIE: Medida revisada. Se hace presente que "
                        "Daniela alcanzará la mayoría de edad")

def test_orden_9_2_audiencia_despues_de_fichas():
    o = obs(fila(programa="RTA Hogar", ficha_ind=d(200), prox_aud=f(4),
                 curador="NO POSEE", oido=d(2)))
    i_curador = o.index("curador ad litem")
    i_oido = o.index("Oído con fecha")
    i_ficha = o.index("superando los 180 días")
    i_aud = o.index("Se cita a audiencia")
    assert i_curador < i_oido < i_ficha < i_aud

def test_ejemplo_10_5_solo_ficha_fae():
    o = obs(fila(nombre="IGNACIO PEREZ", programa="FAE Familia Sur",
                 ingreso=d(130), dias_cumpl=130))
    assert o == ("Ignacio FAE: Medida revisada. Se hace presente que Ignacio "
                 "no tiene ficha FAE, se sugiere confeccionar.")


# ── E-01 corte y transversales ────────────────────────────────────────────────

def test_e01_corta_en_cumplimiento():
    o = obs(fila(programa="OPD Concepción", prox_aud=f(3), curador=""))
    assert "no se encuentra sujeta a seguimiento" in o
    assert "audiencia" not in o

def test_sin_marcadores_ni_dobles_puntos():
    for row in (fila(), fila(dias_cumpl=5, ingreso=d(5)),
                fila(programa="RTA X", ficha_ind=d(300)),
                fila(dias_egresar=0, egreso_proy=f(0), curador="---")):
        o = obs(row, fecha_hoja2=FH2)
        assert "{" not in o and "}" not in o and ".." not in o
        assert o.endswith(".")
