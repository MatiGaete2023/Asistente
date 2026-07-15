# -*- coding: utf-8 -*-
"""
Reglas de validación Fase 1 — CSMP Assistant v8.0

Cada función recibe (df, cols, modo) y retorna dict:
{
    "id": "A1",
    "severidad": "BLOQUEA" | "ADVIERTE",
    "descripcion": "...",
    "detalle": [...],       # filas o columnas afectadas
    "count": int
}
O None si la regla no se dispara.
"""

from datetime import datetime
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from motor.utilidades import normalizar, detectar_tribunal, get_date, get_int


# ── Columnas mínimas por modo ─────────────────────────────────────────────────

COLS_MINIMAS = {
    "ESPERA":        ["DERIVACION", "TRIBUNAL", "NOMBRE"],
    "CUMPLIMIENTO":  ["DERIVACION", "TRIBUNAL", "NOMBRE"],
    "INFORMES":      ["DERIVACION", "TRIBUNAL", "NOMBRE"],
}

COLS_CRITICAS_VACIAS = {
    "ESPERA":       ["DERIVACION", "TRIBUNAL", "NOMBRE"],
    "CUMPLIMIENTO": ["DERIVACION", "TRIBUNAL", "NOMBRE", "FEC.INGRESO EFECTIVO"],
    "INFORMES":     ["DERIVACION", "TRIBUNAL", "NOMBRE", "FECHA VENCIMIENTO",
                     "FEC.VENCIMIENTO", "FEC. VENCIMIENTO"],
}

ALIASES_TRIBUNAL   = ["TRIBUNAL"]
ALIASES_DERIVACION = ["DERIVACION", "DERIVACIÓN", "PROGRAMA"]
ALIASES_NOMBRE     = ["NOMBRE", "NOMBRE COMPLETO"]
ALIASES_NACIMIENTO = ["FEC. NACIMIENTO", "FEC.NACIMIENTO", "FECHA NACIMIENTO", "FEC NACIMIENTO"]
ALIASES_EDAD       = ["EDAD"]


def _col_match(df, aliases):
    """Retorna nombre real de columna en df que matchea algún alias, o None."""
    cols_norm = {normalizar(c): c for c in df.columns}
    for a in aliases:
        key = normalizar(a)
        if key in cols_norm:
            return cols_norm[key]
    return None


# ─────────────────────────────────────────────────────────────────────────────
# A1 — BLOQUEA: Columnas mínimas ausentes
# ─────────────────────────────────────────────────────────────────────────────

def a1_columnas_minimas(df, modo):
    requeridas = {
        "derivacion": ALIASES_DERIVACION,
        "tribunal":   ALIASES_TRIBUNAL,
        "nombre":     ALIASES_NOMBRE,
    }
    faltantes = []
    for clave, aliases in requeridas.items():
        if _col_match(df, aliases) is None:
            faltantes.append(f"{clave.upper()} (buscado: {aliases[0]})")

    if not faltantes:
        return None
    return {
        "id": "A1",
        "severidad": "BLOQUEA",
        "descripcion": "Columnas mínimas ausentes",
        "detalle": faltantes,
        "count": len(faltantes),
    }


# ─────────────────────────────────────────────────────────────────────────────
# A2 — BLOQUEA: RIT duplicados (mismo RIT + NNA + programa)
# ─────────────────────────────────────────────────────────────────────────────

def a2_rit_duplicados(df, modo):
    # RIT puede estar en columna "RIT", "N° RIT" o similar
    col_rit  = _col_match(df, ["RIT", "N° RIT", "N RIT", "NÚMERO RIT", "NUMERO RIT"])
    col_prog = _col_match(df, ALIASES_DERIVACION)
    col_nna  = _col_match(df, ALIASES_NOMBRE)

    if col_rit is None:
        # Sin columna RIT, no se puede validar — solo advierte con count=0
        return None

    claves = [col_rit]
    if col_prog:
        claves.append(col_prog)
    if col_nna:
        claves.append(col_nna)

    sub = df[claves].copy()
    sub = sub.dropna(subset=[col_rit])
    sub[col_rit] = sub[col_rit].astype(str).str.strip()
    duplicados = sub[sub.duplicated(subset=claves, keep=False)]

    if duplicados.empty:
        return None

    detalle = []
    for _, row in duplicados.head(20).iterrows():
        detalle.append(str(row[col_rit]))

    return {
        "id": "A2",
        "severidad": "BLOQUEA",
        "descripcion": "RIT duplicados (mismo RIT + NNA + programa)",
        "detalle": list(dict.fromkeys(detalle)),  # dedup conservando orden
        "count": len(duplicados),
    }


# ─────────────────────────────────────────────────────────────────────────────
# A3 — BLOQUEA: Derivación vacía (programa sin texto)
# ─────────────────────────────────────────────────────────────────────────────

def a3_derivacion_vacia(df, modo):
    col = _col_match(df, ALIASES_DERIVACION)
    if col is None:
        return None  # A1 ya lo captura

    vacias = df[df[col].astype(str).str.strip().isin(["", "nan", "None", "NaT"])]
    if vacias.empty:
        return None

    # Intentar incluir RIT en detalle
    col_rit = _col_match(df, ["RIT", "N° RIT", "N RIT"])
    detalle = []
    for idx, row in vacias.head(20).iterrows():
        rit = str(row[col_rit]).strip() if col_rit else f"fila {idx + 2}"
        detalle.append(rit)

    return {
        "id": "A3",
        "severidad": "BLOQUEA",
        "descripcion": "Derivación vacía (programa sin texto)",
        "detalle": detalle,
        "count": len(vacias),
    }


# ─────────────────────────────────────────────────────────────────────────────
# A4 — ADVIERTE: Fechas futuras en columnas de ingreso/vencimiento
# ─────────────────────────────────────────────────────────────────────────────

def a4_fechas_futuras(df, modo):
    COLS_FECHA = [
        "FEC.INGRESO EFECTIVO", "FEC. INGRESO EFECTIVO", "FEC INGRESO EFECTIVO",
        "FECHA VENCIMIENTO",    "FEC.VENCIMIENTO",        "FEC. VENCIMIENTO",
        "FEC.EGRESO PROYECTADO","FEC. EGRESO PROYECTADO",
        "FEC. NACIMIENTO",      "FEC.NACIMIENTO",
    ]
    hoy = datetime.now()
    afectadas = []

    for alias in COLS_FECHA:
        col = _col_match(df, [alias])
        if col is None:
            continue
        for idx, val in df[col].items():
            fecha = get_date(val)
            if fecha is None:
                continue
            # Solo checar ingreso efectivo y nacimiento como "no deben ser futuras"
            if alias in ["FEC.INGRESO EFECTIVO", "FEC. INGRESO EFECTIVO",
                         "FEC INGRESO EFECTIVO", "FEC. NACIMIENTO", "FEC.NACIMIENTO"]:
                if fecha > hoy:
                    col_rit = _col_match(df, ["RIT", "N° RIT"])
                    rit = str(df.at[idx, col_rit]).strip() if col_rit else f"fila {idx + 2}"
                    afectadas.append(f"{rit} — {col}: {fecha.strftime('%d/%m/%Y')}")

    if not afectadas:
        return None

    return {
        "id": "A4",
        "severidad": "ADVIERTE",
        "descripcion": "Fechas futuras en columnas de ingreso o nacimiento",
        "detalle": afectadas[:20],
        "count": len(afectadas),
    }


# ─────────────────────────────────────────────────────────────────────────────
# A5 — ADVIERTE: Edades inconsistentes (EDAD ≠ FEC.NACIMIENTO ± 1 año)
# ─────────────────────────────────────────────────────────────────────────────

def a5_edades_inconsistentes(df, modo):
    col_edad = _col_match(df, ALIASES_EDAD)
    col_nac  = _col_match(df, ALIASES_NACIMIENTO)
    if col_edad is None or col_nac is None:
        return None

    hoy = datetime.now()
    afectadas = []
    col_rit = _col_match(df, ["RIT", "N° RIT"])

    for idx, row in df.iterrows():
        edad_reg = get_int(row[col_edad])
        fec_nac  = get_date(row[col_nac])
        if edad_reg is None or fec_nac is None:
            continue
        edad_calc = (hoy - fec_nac).days // 365
        if abs(edad_calc - edad_reg) > 1:
            rit = str(row[col_rit]).strip() if col_rit else f"fila {idx + 2}"
            afectadas.append(
                f"{rit} — edad registrada: {edad_reg}, calculada: {edad_calc}"
            )

    if not afectadas:
        return None

    return {
        "id": "A5",
        "severidad": "ADVIERTE",
        "descripcion": "Edad registrada no coincide con fecha de nacimiento (±1 año)",
        "detalle": afectadas[:20],
        "count": len(afectadas),
    }


# ─────────────────────────────────────────────────────────────────────────────
# A6 — ADVIERTE: Edad fuera de rango (< 0 o > 25)
# ─────────────────────────────────────────────────────────────────────────────

def a6_edad_fuera_rango(df, modo):
    col_edad = _col_match(df, ALIASES_EDAD)
    if col_edad is None:
        return None

    col_rit = _col_match(df, ["RIT", "N° RIT"])
    afectadas = []

    for idx, row in df.iterrows():
        edad = get_int(row[col_edad])
        if edad is None:
            continue
        if edad < 0 or edad > 25:
            rit = str(row[col_rit]).strip() if col_rit else f"fila {idx + 2}"
            afectadas.append(f"{rit} — EDAD: {edad}")

    if not afectadas:
        return None

    return {
        "id": "A6",
        "severidad": "ADVIERTE",
        "descripcion": "Edad fuera de rango esperado (< 0 o > 25)",
        "detalle": afectadas[:20],
        "count": len(afectadas),
    }


# ─────────────────────────────────────────────────────────────────────────────
# A7 — ADVIERTE: Campos vacíos masivos (> 20% en columnas críticas)
# ─────────────────────────────────────────────────────────────────────────────

def a7_campos_vacios_masivos(df, modo):
    CRITICAS_POR_MODO = {
        "ESPERA":       ALIASES_DERIVACION + ALIASES_TRIBUNAL + ALIASES_NOMBRE + ALIASES_EDAD,
        "CUMPLIMIENTO": ALIASES_DERIVACION + ALIASES_TRIBUNAL + ALIASES_NOMBRE + ALIASES_EDAD
                        + ["FEC.INGRESO EFECTIVO", "FEC. INGRESO EFECTIVO"],
        "INFORMES":     ALIASES_DERIVACION + ALIASES_TRIBUNAL + ALIASES_NOMBRE
                        + ["FECHA VENCIMIENTO", "FEC.VENCIMIENTO"],
    }
    umbrales = CRITICAS_POR_MODO.get(modo, ALIASES_DERIVACION)
    total = len(df)
    if total == 0:
        return None

    afectadas = []
    ya_revisadas = set()

    for alias in umbrales:
        col = _col_match(df, [alias])
        if col is None or col in ya_revisadas:
            continue
        ya_revisadas.add(col)
        vacias = df[col].astype(str).str.strip().isin(["", "nan", "None", "NaT"]).sum()
        pct = vacias / total
        if pct > 0.20:
            afectadas.append(f"{col}: {vacias}/{total} vacíos ({pct:.0%})")

    if not afectadas:
        return None

    return {
        "id": "A7",
        "severidad": "ADVIERTE",
        "descripcion": "Campos críticos con >20% de valores vacíos",
        "detalle": afectadas,
        "count": len(afectadas),
    }


# ─────────────────────────────────────────────────────────────────────────────
# A8 — ADVIERTE: Tribunal no reconocido (no matchea LAJA/MULCHEN/TOME)
# ─────────────────────────────────────────────────────────────────────────────

def a8_tribunal_no_reconocido(df, modo):
    col = _col_match(df, ALIASES_TRIBUNAL)
    if col is None:
        return None

    col_rit = _col_match(df, ["RIT", "N° RIT"])
    afectadas = []

    for idx, row in df.iterrows():
        val = str(row[col]).strip()
        if val in ("", "nan", "None"):
            continue
        if detectar_tribunal(val) is None:
            rit = str(row[col_rit]).strip() if col_rit else f"fila {idx + 2}"
            afectadas.append(f"{rit} — TRIBUNAL: '{val}'")

    if not afectadas:
        return None

    return {
        "id": "A8",
        "severidad": "ADVIERTE",
        "descripcion": "Valor de tribunal no reconocido (no es LAJA/MULCHEN/TOME)",
        "detalle": afectadas[:20],
        "count": len(afectadas),
    }
