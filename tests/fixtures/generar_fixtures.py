#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tests/fixtures/generar_fixtures.py

Genera 3 Excel SINTETICOS (excel_espera.xlsx, excel_cumplimiento.xlsx,
excel_informes.xlsx) para poder correr tests/test_regresion.py en un
entorno donde NO estan disponibles los 4 Excel reales de produccion
(ese es el caso de este sandbox de desarrollo).

IMPORTANTE (N1 FACT): estos datos son inventados, NO son NNA reales,
NO reflejan las cifras reales de produccion (16/100/333 filas). Sirven
solo para verificar que el pipeline no crashea y que las reglas
disparan sin errores estructurales. La regresion cuantitativa real
(conteos exactos 16/100/333, diffs=0 vs baseline v8.13) SOLO puede
correrse en la maquina de Matias con sus 4 Excel reales — ver
tests/test_regresion.py para instrucciones (variables de entorno).

Uso:
    python tests/fixtures/generar_fixtures.py
"""

from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

HOY = datetime.now()
DIR = Path(__file__).resolve().parent


def d(dias_atras):
    """Fecha `dias_atras` dias antes de hoy, formato DD/MM/YYYY."""
    return (HOY - timedelta(days=dias_atras)).strftime("%d/%m/%Y")


def d_fwd(dias_adelante):
    return (HOY + timedelta(days=dias_adelante)).strftime("%d/%m/%Y")


def nacimiento_edad(anios, dias_extra=0):
    """Fecha de nacimiento para que la persona tenga `anios` años hoy
    (+ dias_extra para afinar proximidad a un corte)."""
    return (HOY - timedelta(days=anios * 365 + dias_extra)).strftime("%d/%m/%Y")


def nacimiento_dias_hasta_18(dias_hasta_18):
    """Fecha de nacimiento tal que el 18º cumpleaños cae en `dias_hasta_18`
    días a partir de hoy (aprox., usa 365.25 d/año — suficiente para smoke test)."""
    return (HOY + timedelta(days=dias_hasta_18) - timedelta(days=18 * 365.25)).strftime("%d/%m/%Y")


# ═══════════════════════════════════════════════════════════════════════
# ESPERA — 16 filas sintéticas cubriendo las 9 ramas de reglas
# ═══════════════════════════════════════════════════════════════════════
ESPERA_ROWS = [
    # RIT, RUT, NOMBRE, TRIBUNAL, DERIVACION, FEC.NACIMIENTO, EDAD, CURADOR, FEC.OIDO, FEC.RESOLUCION, T ESPERA
    ("P-1-2026", "11111111-1", "Ana Torres",    "LAJA",    "OPD Laja",         "", "10", "", "", "", 5),
    ("P-2-2026", "12222222-2", "Luis Soto",     "TOME",    "RTA Esperanza",    nacimiento_edad(18, 10), "18", "33333333-3", "", "", 40),
    ("P-3-2026", "13333333-3", "Marta Diaz",    "MULCHEN", "PIE Renacer",      "", "17", "44444444-4", "", "", 10),
    ("P-4-2026", "14444444-4", "Pedro Rojas",   "MULCHEN", "PIE Renacer",      nacimiento_dias_hasta_18(30), "17", "55555555-5", "", "", 10),
    ("P-5-2026", "15555555-5", "Sofia Vega",    "LAJA",    "DCE Diagnostico",  "", "8", "66666666-6", "", d(5), 10),
    ("P-6-2026", "16666666-6", "Jose Paredes",  "LAJA",    "DCE Diagnostico",  "", "9", "77777777-7", "", "", 10),
    ("P-7-2026", "17777777-7", "Elena Munoz",   "TOME",    "DCE Diagnostico",  "", "11", "88888888-8", "", "", 45),
    ("P-8-2026", "18888888-8", "Carlos Reyes",  "MULCHEN", "RTT Nueva Vida",   "", "12", "99999999-9", "", "", 35),
    ("P-9-2026", "19999999-9", "Valentina Paz", "LAJA",    "RES Los Alamos",   "", "13", "12121212-1", "", "", 65),
    ("P-10-2026","10101010-1", "Diego Fuentes", "TOME",    "PAS Ambulatorio",  "", "14", "", "", "", 20),
    ("P-11-2026","10202020-2", "Camila Ortiz",  "LAJA",    "FAE Familia Sur",  "", "15", "13131313-1", "", "", 15),
    ("P-12-2026","10303030-3", "Ignacio Rios",  "MULCHEN", "PRM Esperanza",    "", "9", "14141414-1", d(20), "", 8),
    ("P-13-2026","10404040-4", "Fernanda Cruz", "TOME",    "RFA Amanecer",     "", "10", "(0-0) Institución: CAJ Biobío", "", "", 12),
    ("P-14-2026","10505050-5", "Matias Leon",   "LAJA",    "RPPM Renacer",     "", "16", "15151515-1", "", d(2), 3),
    ("P-15-2026","10606060-6", "Isidora Silva", "MULCHEN", "PIE Renacer",      "", "7", "16161616-1", "", "", 2),
    ("P-16-2026","10707070-7", "Tomas Bravo",   "TOME",    "AFT Familia",      "", "6", "17171717-1", "", "", 61),
]

# ═══════════════════════════════════════════════════════════════════════
# CUMPLIMIENTO — Hoja1 (25 filas) + Hoja2 informes por vencer (10 filas)
# ═══════════════════════════════════════════════════════════════════════
CUMPL_ROWS = [
    # RIT,RUT,NOMBRE,TRIBUNAL,DERIVACION,FEC.NACIM,EDAD,CURADOR,FEC.OIDO,
    # DIAS_EGRESAR(no usado, se calcula desde FEC.EGRESO_PROY-ficticia -> usamos directo el string),
    # FEC.INGRESO_EFECTIVO, FEC.EGRESO_PROYECTADO, FEC.ACT.F.RESIDENCIAL, FEC.ACT.F.INDIVIDUAL, FEC.ACT.F.FAE, DIAS_PARA_EGRESAR
    ("C-1-2026","20101010-1","Rosa Campos",  "LAJA",   "OPD Laja",        "", "11", "",           "", d(10),  d(20),  "", "", "", 400),
    ("C-2-2026","20202020-2","Hugo Nunez",   "TOME",   "RTA Esperanza",   nacimiento_edad(18,5), "18", "21212121-2", "", d(100), d(50), d(200), d(200), "", 100),
    ("C-3-2026","20303030-3","Paula Ibanez", "MULCHEN","PIE Renacer",     "", "17", "22222222-2", "", d(200), d(400), "", "", "", 55),
    ("C-4-2026","20404040-4","Sergio Diaz",  "MULCHEN","PIE Renacer",     nacimiento_edad(17,-40), "17", "23232323-2", "", d(15), d(300), "", "", "", 200),
    ("C-5-2026","20505050-5","Karen Torres", "LAJA",   "RES Los Alamos",  "", "10", "24242424-2", "", d(300), d(-5), d(200), d(200), "", -5),
    ("C-6-2026","20606060-6","Andres Silva", "LAJA",   "RES Los Alamos",  "", "9",  "25252525-2", "", d(50),  d(30),  d(200), d(200), "", 30),
    ("C-7-2026","20707070-7","Monica Reyes", "TOME",   "RTT Nueva Vida",  "", "13", "26262626-2", "", d(500), d(300), d(200), d(50),  "", 300),
    ("C-8-2026","20808080-8","Felipe Rojas", "MULCHEN","RFA Amanecer",    "", "14", "27272727-2", "", d(60),  d(400), d(10),  d(10),  "", 400),
    ("C-9-2026","20909090-9","Antonia Vera", "TOME",   "RPPM Renacer",    "", "12", "",           "", d(90),  d(500), d(190), d(190), "", 500),
    ("C-10-2026","21010101-0","Ricardo Paz", "LAJA",   "FAE Familia Sur", "", "8",  "28282828-2", "", d(150), d(600), "", "", "", 600),
    ("C-11-2026","21111111-1","Constanza Loyola","TOME","PRM Esperanza",  "", "9",  "29292929-2", "", d(25),  d(700), "", "", "", 700),
    ("C-12-2026","21212121-2","Benjamin Duran","MULCHEN","PAS Ambulatorio","", "10", "30303030-3", "", d(400), d(800), "", "", "", 800),
    ("C-13-2026","21313131-3","Javiera Cid",  "LAJA",   "DCE Diagnostico", "", "11", "31313131-3", "", d(10),  d(900), "", "", "", 900),
    ("C-14-2026","21414141-4","Emilio Gatica","TOME",   "AFT Familia",     "", "15", "32323232-3", "", d(200), d(1000),"", "", "", 1000),
    ("C-15-2026","21515151-5","Trinidad Salas","LAJA",  "RES Los Alamos",  "", "16", "33333433-3", "", d(35),  d(1100),d(1200),d(1200),"",1100),
    ("C-16-2026","21616161-6","Maximiliano Rojo","MULCHEN","RTA Esperanza","", "6",  "34343434-3", "", d(400), d(1200),d(5),   d(5),   "",1200),
    ("C-17-2026","21717171-7","Florencia Mora","TOME",  "RTT Nueva Vida",  "", "7",  "",           "", d(500), d(1300),d(210), d(210), "",1300),
    ("C-18-2026","21818181-8","Cristobal Leiva","LAJA", "RFA Amanecer",    "", "8",  "35353535-3", "", d(600), d(1400),d(220), d(220), "",1400),
    ("C-19-2026","21919191-9","Renata Vidal", "MULCHEN","PIE Renacer",     "", "9",  "(0-0) Institución: CAJ Biobío","", d(700), d(1500),"", "", "",1500),
    ("C-20-2026","22020202-0","Vicente Aguilera","TOME","RPPM Renacer",    "", "10", "36363636-3", "", d(800), d(1600),d(230),d(230),"",1600),
    ("C-21-2026","22121212-1","Amanda Bustos","LAJA",   "FAE Familia Sur", "", "11", "37373737-3", "", d(125), d(1700),"", "", d(125),1700),
    ("C-22-2026","22222323-2","Agustin Perez","MULCHEN","FAE Familia Sur", "", "12", "38383838-3", "", d(130), d(1800),"", "", "",1800),
    ("C-23-2026","22323232-3","Isabella Nunez","TOME",  "PRM Esperanza",   "", "13", "39393939-3", "", d(900), d(1900),"", "", "",1900),
    ("C-24-2026","22424242-4","Bastian Molina","LAJA",  "PAS Ambulatorio", "", "14", "40404040-4", "", d(1000),d(2000),"", "", "",2000),
    ("C-25-2026","22525252-5","Josefa Aravena","MULCHEN","DCE Diagnostico","", "15", "41414141-4", d(20), d(1100),d(2100),"", "","",2100),
]

CUMPL_HOJA2_ROWS = [
    # RIT, RUT MENOR, NOMBRE MENOR, TRIBUNAL, NOMBRE CENTRO, FECHA VENCIMIENTO (futura)
    ("C-6-2026", "20606060-6", "Andres Silva",  "LAJA",   "RES Los Alamos", d_fwd(20)),
    ("C-7-2026", "20707070-7", "Monica Reyes",  "TOME",   "RTT Nueva Vida", d_fwd(15)),
    ("C-8-2026", "20808080-8", "Felipe Rojas",  "MULCHEN","RFA Amanecer",   d_fwd(30)),
    ("C-13-2026","21313131-3","Javiera Cid",    "LAJA",   "DCE Diagnostico",d_fwd(10)),
    ("C-14-2026","21414141-4","Emilio Gatica",  "TOME",   "AFT Familia",    d_fwd(25)),
]

# ═══════════════════════════════════════════════════════════════════════
# INFORMES — 20 filas
# ═══════════════════════════════════════════════════════════════════════
INFORMES_ROWS = [
    ("I-1-2026","30101010-1","Nicolas Paredes","LAJA",   "OPD Laja",        d(5),   ""),
    ("I-2-2026","30202020-2","Daniela Fuentes","TOME",   "RTA Esperanza",   d_fwd(10),  d(300)),
    ("I-3-2026","30303030-3","Gonzalo Herrera","MULCHEN","PIE Renacer",     d_fwd(40),  d(400)),
    ("I-4-2026","30404040-4","Yasna Cortes",   "LAJA",   "RES Los Alamos",  d_fwd(60),  d(200)),
    ("I-5-2026","30505050-5","Ivan Espinoza",  "TOME",   "DCE Diagnostico", d_fwd(20),  d(100)),
    ("I-6-2026","30606060-6","Katherine Munoz","MULCHEN","DCE Diagnostico", d_fwd(5),   d(90)),
    ("I-7-2026","30707070-7","Esteban Vargas", "LAJA",   "RTT Nueva Vida",  "",     d(150)),
    ("I-8-2026","30808080-8","Barbara Sepulveda","TOME", "RFA Amanecer",    d(1),   d(250)),
    ("I-9-2026","30909090-9","Alvaro Contreras","MULCHEN","RPPM Renacer",   d(0),   d(180)),
    ("I-10-2026","31010101-0","Natalia Cabrera","LAJA",  "FAE Familia Sur", d_fwd(44),  d(220)),
]


def _to_df(rows, columnas):
    return pd.DataFrame(rows, columns=columnas)


def generar():
    # ESPERA
    cols_espera = ["RIT", "RUT", "NOMBRE", "TRIBUNAL", "DERIVACION",
                   "FEC. NACIMIENTO", "EDAD", "CURADOR", "FEC. OIDO",
                   "FEC. RESOLUCIÓN", "T ESPERA"]
    df_espera = _to_df(ESPERA_ROWS, cols_espera)
    df_espera.to_excel(DIR / "excel_espera.xlsx", index=False)

    # CUMPLIMIENTO (2 hojas)
    cols_cumpl = ["RIT", "RUT", "NOMBRE", "TRIBUNAL", "DERIVACION",
                  "FEC. NACIMIENTO", "EDAD", "CURADOR", "FEC. OIDO",
                  "FEC.INGRESO EFECTIVO", "FEC.EGRESO PROYECTADO",
                  "FEC.ACT.F.RESIDENCIAL", "FEC.ACT.F.INDIVIDUAL",
                  "FEC.ACT.F.FAE", "DIAS PARA EGRESAR"]
    df_cumpl = _to_df(CUMPL_ROWS, cols_cumpl)
    cols_h2 = ["RIT", "RUT MENOR", "NOMBRE MENOR", "TRIBUNAL",
               "NOMBRE CENTRO", "FECHA VENCIMIENTO"]
    df_h2 = _to_df(CUMPL_HOJA2_ROWS, cols_h2)
    with pd.ExcelWriter(DIR / "excel_cumplimiento.xlsx", engine="openpyxl") as xw:
        df_cumpl.to_excel(xw, sheet_name="CUMPLIMIENTO", index=False)
        df_h2.to_excel(xw, sheet_name="INFORMES POR VENCER", index=False)

    # INFORMES
    cols_inf = ["RIT", "RUT", "NOMBRE", "TRIBUNAL", "DERIVACION",
                "FECHA VENCIMIENTO", "FEC.INGRESO EFECTIVO"]
    df_inf = _to_df(INFORMES_ROWS, cols_inf)
    df_inf.to_excel(DIR / "excel_informes.xlsx", index=False)

    print(f"OK: fixtures generados en {DIR}")
    print(f"  excel_espera.xlsx        : {len(df_espera)} filas")
    print(f"  excel_cumplimiento.xlsx  : {len(df_cumpl)} filas (Hoja1) + {len(df_h2)} (Hoja2)")
    print(f"  excel_informes.xlsx      : {len(df_inf)} filas")


if __name__ == "__main__":
    generar()
