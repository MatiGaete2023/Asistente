# -*- coding: utf-8 -*-
"""
Reglas de validación — CSMP Assistant v9.0.1

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
from motor.utilidades import normalizar, detectar_tribunal, get_date, get_int
from motor.columnas_comunes import (
    ALIAS_DIAS_CUMPLIMIENTO,
    ALIAS_DIAS_EGRESO,
    ALIAS_EGRESO_PROYECTADO,
    ALIAS_ESPERA,
    ALIAS_FICHA_FAE,
    ALIAS_FICHA_INDIVIDUAL,
    ALIAS_INGRESO_EFECTIVO,
    ALIAS_NACIMIENTO,
    ALIAS_NOMBRE as ALIASES_NOMBRE,
    ALIAS_PROGRAMA as ALIASES_DERIVACION,
    ALIAS_RESOLUCION,
    ALIAS_RIT,
    ALIAS_TRIBUNAL as ALIASES_TRIBUNAL,
    ALIAS_VENCIMIENTO,
)


# ── Columnas mínimas por modo ─────────────────────────────────────────────────

ALIASES_NACIMIENTO = ALIAS_NACIMIENTO
ALIASES_EDAD       = ["EDAD"]
ALIASES_RIT        = ALIAS_RIT

REQUERIDAS_POR_MODO = {
    "ESPERA": {
        "tiempo_espera": ALIAS_ESPERA,
        "fecha_resolucion": ALIAS_RESOLUCION,
    },
    "CUMPLIMIENTO": {
        "dias_cumplimiento": ALIAS_DIAS_CUMPLIMIENTO,
        "dias_egreso": ALIAS_DIAS_EGRESO,
        "fecha_ingreso": ALIAS_INGRESO_EFECTIVO,
        "fecha_egreso": ALIAS_EGRESO_PROYECTADO,
        "ficha_individual": ALIAS_FICHA_INDIVIDUAL,
        "ficha_fae": ALIAS_FICHA_FAE,
    },
    "INFORMES": {
        "fecha_vencimiento": ALIAS_VENCIMIENTO,
    },
}


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
        "rit":        ALIASES_RIT,
        "derivacion": ALIASES_DERIVACION,
        "tribunal":   ALIASES_TRIBUNAL,
        "nombre":     ALIASES_NOMBRE,
    }
    requeridas.update(REQUERIDAS_POR_MODO.get(modo.upper(), {}))
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
    GRUPOS_NO_FUTUROS = [
        ["FEC.INGRESO EFECTIVO", "FEC. INGRESO EFECTIVO", "FEC INGRESO EFECTIVO"],
        ["FEC. NACIMIENTO", "FEC.NACIMIENTO", "FECHA NACIMIENTO", "FEC NACIMIENTO"],
    ]
    hoy = datetime.now()
    afectadas = []
    revisadas = set()

    for aliases in GRUPOS_NO_FUTUROS:
        col = _col_match(df, aliases)
        if col is None or col in revisadas:
            continue
        revisadas.add(col)
        for idx, val in df[col].items():
            fecha = get_date(val)
            if fecha is None:
                continue
            if fecha > hoy:
                col_rit = _col_match(df, ALIASES_RIT)
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
        if val.lower() in ("", "nan", "none", "nat"):
            continue  # A9 bloquea vacíos con un mensaje específico.
        if detectar_tribunal(val) is None:
            rit = str(row[col_rit]).strip() if col_rit else f"fila {idx + 2}"
            afectadas.append(f"{rit} — TRIBUNAL: '{val}'")

    if not afectadas:
        return None

    return {
        "id": "A8",
        "severidad": "BLOQUEA",
        "descripcion": "Valor de tribunal no reconocido (no es LAJA/MULCHEN/TOME)",
        "detalle": afectadas[:20],
        "count": len(afectadas),
    }


def a9_campos_identificacion_vacios(df, modo):
    """BLOQUEA filas que no pueden identificarse o enrutarse con seguridad."""
    campos = {
        "RIT": ALIASES_RIT,
        "NOMBRE": ALIASES_NOMBRE,
        "TRIBUNAL": ALIASES_TRIBUNAL,
    }
    afectados = []
    total_afectados = 0
    for etiqueta, aliases in campos.items():
        col = _col_match(df, aliases)
        if col is None:
            continue  # A1 ya informa la columna ausente.
        vacios = df[col].isna() | df[col].astype(str).str.strip().str.lower().isin(
            ["", "nan", "none", "nat"])
        total_afectados += int(vacios.sum())
        for idx in df.index[vacios][:20]:
            afectados.append(f"fila {idx + 2} — {etiqueta} vacío")

    if total_afectados == 0:
        return None
    return {
        "id": "A9",
        "severidad": "BLOQUEA",
        "descripcion": "Campos de identificación o enrutamiento vacíos",
        "detalle": afectados[:20],
        "count": total_afectados,
    }


def a10_valores_operativos_invalidos(df, modo):
    """BLOQUEA valores sin los cuales una regla principal no puede decidir."""
    modo = modo.upper()
    afectados = []

    if modo == "ESPERA":
        col = _col_match(df, REQUERIDAS_POR_MODO[modo]["tiempo_espera"])
        if col:
            for idx, valor in df[col].items():
                dias = get_int(valor)
                if dias is None or dias < 0:
                    afectados.append(f"fila {idx + 2} — {col}: valor inválido")

    elif modo == "CUMPLIMIENTO":
        for clave in ("dias_cumplimiento", "dias_egreso"):
            col = _col_match(df, REQUERIDAS_POR_MODO[modo][clave])
            if not col:
                continue
            for idx, valor in df[col].items():
                if get_int(valor) is None:
                    afectados.append(f"fila {idx + 2} — {col}: valor inválido")

    elif modo == "INFORMES":
        col = _col_match(df, REQUERIDAS_POR_MODO[modo]["fecha_vencimiento"])
        if col:
            for idx, valor in df[col].items():
                if get_date(valor) is None:
                    afectados.append(f"fila {idx + 2} — {col}: fecha inválida")

    if not afectados:
        return None
    return {
        "id": "A10",
        "severidad": "BLOQUEA",
        "descripcion": "Valores operativos ausentes o inválidos",
        "detalle": afectados[:20],
        "count": len(afectados),
    }
