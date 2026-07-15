#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
motor/procesador.py — v9.0.1
Optimizaciones:
  CLAIM-3: _guardar_excel usa dataframe_to_rows (3x más rápido)
  CLAIM-4: _construir_indice_hoja2 usa to_dict('records') (8.6x más rápido)
"""

import logging
import os
import uuid
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows   # CLAIM-3

from .utilidades import (
    normalizar_match, detectar_tribunal, get_date, validar_archivo_excel,
    es_texto_formula,
)
from .mapeo_columnas import mapear_columnas, mapear_columnas_hoja2
from .composicion import Incidencias
from .reglas_espera import generar_observacion_espera
from .reglas_cumplimiento import generar_observacion_cumplimiento
from .reglas_informes import generar_observacion_informes
from validador.precheck import validar_excel

logger = logging.getLogger(__name__)


# ── cruce Hoja2 ───────────────────────────────────────────────────────────────

def _construir_indice_hoja2(df_h2, cols_h2):
    """
    CLAIM-4: to_dict('records') en vez de iterrows — 8.6x más rápido.
    """
    indice = {}
    campos = ["rit", "rut", "nombre", "tribunal", "programa", "vencimiento"]
    cols_v = [cols_h2.get(c) for c in campos]
    if any(c is None for c in cols_v):
        # Hoja2 malformada degrada a "sin cruce" con aviso fuerte (Mejoras §2.2);
        # nunca aborta el procesamiento completo de CUMPLIMIENTO.
        faltantes = [campo for campo, columna in zip(campos, cols_v) if columna is None]
        logger.warning("Hoja2: faltan columnas (%s) — cruce desactivado", ", ".join(faltantes))
        return {}
    col_rit, col_rut, col_nom, col_tri, col_prog, col_venc = cols_v
    hoy = datetime.now().date()

    for row in df_h2.to_dict("records"):          # CLAIM-4
        venc = get_date(row.get(col_venc))
        if venc is None:
            continue
        vd = venc.date() if hasattr(venc, "date") else venc
        if vd <= hoy:  # descartar vencidas incluyendo hoy
            continue
        rit = normalizar_match(row.get(col_rit))
        if not rit:
            continue
        clave = (
            rit,
            normalizar_match(row.get(col_rut)),
            normalizar_match(row.get(col_nom)),
            normalizar_match(row.get(col_tri)),
            normalizar_match(row.get(col_prog)),
        )
        # Catálogo C-10: con duplicados en Hoja2 "se conserva la última fila".
        indice[clave] = vd
    return indice


def _clave_fila_h1(row, cols):
    keys = ("rit", "rut", "nombre", "tribunal", "programa")
    if not all(cols.get(k) for k in keys):
        return None
    return tuple(normalizar_match(row.get(cols[k], "")) for k in keys)



# ── guardado Excel ─────────────────────────────────────────────────────────────

def _neutralizar_formulas(ws):
    """Fuerza tipo string en celdas que openpyxl escribió como fórmula.

    Conserva el texto EXACTO (sin prefijos visibles) y evita que un valor
    '=...' proveniente de la planilla se ejecute al abrir el archivo.
    """
    for fila in ws.iter_rows():
        for celda in fila:
            if celda.data_type == "f" or es_texto_formula(celda.value):
                celda.data_type = "s"


def _guardar_excel(df, modo, ruta, q, incidencias=None):
    """
    CLAIM-3: dataframe_to_rows en vez de iterrows — 3x más rápido.
    """
    try:
        Path(ruta).mkdir(parents=True, exist_ok=True)
        nombre = f"RUS_{modo}_{datetime.now():%Y%m%d_%H%M%S_%f}.xlsx"
        arch   = Path(ruta) / nombre
        temporal = arch.with_name(f".{arch.stem}.{uuid.uuid4().hex}.tmp.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.title = modo

        # CLAIM-3: escritura en bloque
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)
        _neutralizar_formulas(ws)

        if incidencias is not None and len(incidencias):
            ws_val = wb.create_sheet("VALIDACION")
            for r in dataframe_to_rows(incidencias.como_dataframe(), index=False, header=True):
                ws_val.append(r)
            _neutralizar_formulas(ws_val)
            q.put(("log", f"⚠️ {len(incidencias)} incidencias de validación — ver hoja VALIDACION"))

        # Estilos del encabezado (fila 1)
        for c in ws[1]:
            c.font      = Font(bold=True, color="FFFFFF")
            c.fill      = PatternFill("solid", fgColor="1F4E78")
            c.alignment = Alignment(horizontal="center")

        # Ancho de columnas
        for col in ws.columns:
            ancho = max((len(str(c.value or "")) for c in col), default=0)
            ws.column_dimensions[col[0].column_letter].width = min(ancho + 2, 60)

        # Guardado atómico. Sin os.fsync: sobre un handle de solo lectura
        # falla en Windows (FlushFileBuffers exige acceso de escritura) y
        # habría abortado TODOS los guardados en producción.
        wb.save(temporal)
        os.replace(temporal, arch)
        q.put(("log", f"✅ Guardado: {nombre}"))
        q.put(("log", f"📁 {arch}"))
        return nombre
    except Exception as e:
        if "temporal" in locals():
            temporal.unlink(missing_ok=True)
        q.put(("log", f"❌ Error al guardar: {e}"))
        return None

def _prevalidar(df, modo, config, q):
    resultado = validar_excel(df, modo, ruta_salida_reportes=config.get("ruta_salida_excel"))
    for a in resultado.get("anomalias_bloqueantes", []):
        q.put(("log", f"❌ {a['id']}: {a['count']} — {a['descripcion']}"))
    for a in resultado.get("anomalias_advertencia", []):
        q.put(("log", f"⚠️ {a['id']}: {a['count']} — {a['descripcion']}"))
    if resultado.get("ruta_reporte_html"):
        q.put(("log", f"📋 Reporte validación: {resultado['ruta_reporte_html']}"))
    return resultado.get("puede_procesar", False)


# ── procesadores ──────────────────────────────────────────────────────────────

def _insertar_columnas_salida(df):
    df = df.drop(columns=[c for c in ["OBSERVACION", "FECHA_OBS", "TT", "CC", "RES"] if c in df.columns])
    df["OBSERVACION"] = ""
    pos = df.columns.get_loc("OBSERVACION")
    df.insert(pos, "FECHA_OBS", datetime.now().strftime("%d/%m/%Y"))
    pos = df.columns.get_loc("OBSERVACION") + 1
    df.insert(pos, "TT", "")
    df.insert(pos + 1, "CC", "")
    df.insert(pos + 2, "RES", "")
    return df


def _calcular_simple(df, cols, fn_regla):
    incidencias = Incidencias()
    col_trib = cols.get("tribunal")

    def aplicar(row):
        try:
            trib = detectar_tribunal(str(row.get(col_trib, ""))) if col_trib else None
            if trib is None:
                incidencias.agregar(row.name + 2, row.get(cols.get("rit"), "") if cols.get("rit") else "", "G-06", "tribunal no reconocido")
            return fn_regla(row, trib, cols, incidencias=incidencias, fila_excel=row.name + 2)
        except Exception as e:
            raise RuntimeError(f"Fila Excel {row.name + 2}: {e}") from e

    df_r = _insertar_columnas_salida(df.copy())
    df_r["OBSERVACION"] = df.apply(aplicar, axis=1)
    return df_r, incidencias

def _procesar_simple(df, modo, fn_regla, config, q):
    """Procesador genérico para ESPERA e INFORMES."""
    cols     = mapear_columnas(df, modo)
    col_prog = cols.get("programa")
    col_trib = cols.get("tribunal")
    if not col_prog or not col_trib:
        q.put(("log", "❌ Columnas DERIVACION o TRIBUNAL no encontradas"))
        q.put(("done", f"❌ {modo} cancelado: columnas no encontradas."))
        return

    df_r, incidencias = _calcular_simple(df, cols, fn_regla)
    nombre = _guardar_excel(df_r, modo, config["ruta_salida_excel"], q, incidencias)
    if nombre:
        q.put(("done", f"{modo} completado.\n{len(df_r)} casos.\nArchivo: {nombre}"))
    else:
        q.put(("done", f"❌ {modo}: error al guardar el archivo."))


def _calcular_cumplimiento(df_h1, cols, indice):
    incidencias = Incidencias()
    col_trib = cols.get("tribunal")

    def aplicar(row):
        try:
            trib = detectar_tribunal(str(row.get(col_trib, ""))) if col_trib else None
            if trib is None:
                incidencias.agregar(row.name + 2, row.get(cols.get("rit"), "") if cols.get("rit") else "", "G-06", "tribunal no reconocido")
            clave = _clave_fila_h1(row, cols)
            fecha_h2 = indice.get(clave) if clave and indice else None
            return generar_observacion_cumplimiento(row, trib, cols, fecha_hoja2=fecha_h2, incidencias=incidencias, fila_excel=row.name + 2)
        except Exception as e:
            raise RuntimeError(f"Fila Excel {row.name + 2}: {e}") from e

    df_r = _insertar_columnas_salida(df_h1.copy())
    df_r["OBSERVACION"] = df_h1.apply(aplicar, axis=1)
    return df_r, incidencias

def _procesar_cumplimiento(df_h1, df_h2, config, q):
    cols     = mapear_columnas(df_h1, "CUMPLIMIENTO")
    col_prog = cols.get("programa")
    col_trib = cols.get("tribunal")
    if not col_prog or not col_trib:
        q.put(("log", "❌ Columnas DERIVACION o TRIBUNAL no encontradas"))
        q.put(("done", "❌ CUMPLIMIENTO cancelado: columnas no encontradas."))
        return

    indice = {}
    if df_h2 is not None:
        indice = _construir_indice_hoja2(df_h2, mapear_columnas_hoja2(df_h2))
        q.put(("log", f"📋 Hoja2: {len(indice)} registros con fecha futura"))
    else:
        q.put(("log", "ℹ️  Sin Hoja2 — cruce desactivado"))
    if not indice:
        q.put(("warn_hoja2", "Sin Hoja2 utilizable: HOY no se generará NINGUNA observación de próximo informe (C-10). Verifica el archivo."))

    df_r, incidencias = _calcular_cumplimiento(df_h1, cols, indice)
    nombre = _guardar_excel(df_r, "CUMPLIMIENTO", config["ruta_salida_excel"], q, incidencias)
    if nombre:
        q.put(("done", f"CUMPLIMIENTO completado.\n{len(df_r)} casos.\nArchivo: {nombre}"))
    else:
        q.put(("done", "❌ CUMPLIMIENTO: error al guardar el archivo."))


# ── dispatcher ────────────────────────────────────────────────────────────────

def _leer_excel(path):
    path = validar_archivo_excel(path)
    engine = "xlrd" if Path(path).suffix.lower() == ".xls" else "openpyxl"
    df = pd.read_excel(path, engine=engine)
    df.columns = [str(x).strip() for x in df.columns]
    return df.dropna(how="all").fillna("")


def _leer_cumplimiento(path, q=None):
    """Carga ambas hojas una sola vez y libera siempre el archivo origen."""
    path = validar_archivo_excel(path)
    engine = "xlrd" if path.suffix.lower() == ".xls" else "openpyxl"
    with pd.ExcelFile(path, engine=engine) as xl:
        names = {s.strip().lower(): s for s in xl.sheet_names}
        hoja_key = "cumplimiento" if "cumplimiento" in names else next(iter(names))
        df_h1 = pd.read_excel(xl, sheet_name=names[hoja_key])
        df_h1.columns = [str(x).strip() for x in df_h1.columns]
        df_h1 = df_h1.dropna(how="all").fillna("")

        otras = [s for s in xl.sheet_names if s.strip().lower() != hoja_key]
        df_h2 = None
        if otras:
            try:
                df_h2 = pd.read_excel(xl, sheet_name=otras[0])
                df_h2.columns = [str(x).strip() for x in df_h2.columns]
                df_h2 = df_h2.dropna(how="all").fillna("")
                if q is not None:
                    q.put(("log", f"📄 Hoja2: '{otras[0]}' ({len(df_h2)} filas)"))
            except Exception as exc:
                if q is not None:
                    q.put(("log", f"⚠️ No se pudo cargar Hoja2: {exc}"))
        return df_h1, df_h2


def procesar(df_o_path, modo, config, q):
    if modo == "CUMPLIMIENTO":
        if isinstance(df_o_path, (str, Path)):
            path = str(df_o_path)
            try:
                df_h1, df_h2 = _leer_cumplimiento(path, q=q)

                q.put(("log", f"✓ {len(df_h1)} filas CUMPLIMIENTO"))
                if not _prevalidar(df_h1, "CUMPLIMIENTO", config, q):
                    q.put(("done", "❌ CUMPLIMIENTO cancelado por validación.")); return
                _procesar_cumplimiento(df_h1, df_h2, config, q)
            except Exception as e:
                q.put(("log", f"❌ Error cargando archivo: {e}"))
                q.put(("done", f"❌ CUMPLIMIENTO cancelado: {e}"))
        else:
            df = df_o_path
            df.columns = [str(x).strip() for x in df.columns]
            df = df.dropna(how="all").fillna("")
            if not _prevalidar(df, "CUMPLIMIENTO", config, q):
                q.put(("done", "❌ CUMPLIMIENTO cancelado por validación.")); return
            _procesar_cumplimiento(df, None, config, q)

    elif modo in ("ESPERA", "INFORMES"):
        fn = generar_observacion_espera if modo == "ESPERA" else generar_observacion_informes
        try:
            df = (_leer_excel(df_o_path)
                  if isinstance(df_o_path, (str, Path))
                  else df_o_path.dropna(how="all").fillna(""))
            q.put(("log", f"✓ {len(df)} filas {modo}"))
            if not _prevalidar(df, modo, config, q):
                q.put(("done", f"❌ {modo} cancelado por validación.")); return
            _procesar_simple(df, modo, fn, config, q)
        except Exception as e:
            q.put(("log", f"❌ Error cargando archivo: {e}"))
            q.put(("done", f"❌ {modo} cancelado: {e}"))
    else:
        q.put(("log", f"❌ Modo desconocido: {modo}"))
        q.put(("done", f"❌ Modo desconocido: {modo}"))


# ── vista previa (S3 — v8.14) ──────────────────────────────────────────────────

def calcular_preview(df_o_path, modo):
    """
    Calcula el DataFrame con columna OBSERVACION para `modo`, IDENTICO
    byte-a-byte al que produce procesar() (misma logica, mismos textos,
    mismo cruce Hoja2 en CUMPLIMIENTO), pero SIN escribir ningun archivo
    a disco. Usado exclusivamente por la vista previa de gui/app.py (S3).

    Retorna (df_resultado, None) en exito, o (None, mensaje_error) si
    faltan columnas obligatorias o el modo es desconocido.
    """
    try:
        if modo == "CUMPLIMIENTO":
            if isinstance(df_o_path, (str, Path)):
                df_h1, df_h2 = _leer_cumplimiento(df_o_path)
            else:
                df_h1 = df_o_path.copy()
                df_h1.columns = [str(x).strip() for x in df_h1.columns]
                df_h1 = df_h1.dropna(how="all").fillna("")
                df_h2 = None

            cols = mapear_columnas(df_h1, "CUMPLIMIENTO")
            if not cols.get("programa") or not cols.get("tribunal"):
                return None, "Columnas DERIVACION o TRIBUNAL no encontradas"
            validacion = validar_excel(df_h1, "CUMPLIMIENTO")
            if not validacion.get("puede_procesar", False):
                ids = ", ".join(
                    a["id"] for a in validacion.get("anomalias_bloqueantes", [])
                )
                return None, f"Validación bloqueante: {ids}"

            indice = {}
            if df_h2 is not None:
                indice = _construir_indice_hoja2(df_h2, mapear_columnas_hoja2(df_h2))

            df_r, _inc = _calcular_cumplimiento(df_h1, cols, indice)
            return df_r, None

        elif modo in ("ESPERA", "INFORMES"):
            fn = generar_observacion_espera if modo == "ESPERA" else generar_observacion_informes
            df = (_leer_excel(df_o_path)
                  if isinstance(df_o_path, (str, Path))
                  else df_o_path.dropna(how="all").fillna(""))
            cols = mapear_columnas(df, modo)
            if not cols.get("programa") or not cols.get("tribunal"):
                return None, "Columnas DERIVACION o TRIBUNAL no encontradas"
            validacion = validar_excel(df, modo)
            if not validacion.get("puede_procesar", False):
                ids = ", ".join(
                    a["id"] for a in validacion.get("anomalias_bloqueantes", [])
                )
                return None, f"Validación bloqueante: {ids}"

            df_r, _inc = _calcular_simple(df, cols, fn)
            return df_r, None

        else:
            return None, f"Modo desconocido: {modo}"

    except Exception as e:
        return None, f"Error al calcular vista previa: {e}"
