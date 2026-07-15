#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
motor/utilidades.py — v9.0.1
Optimizaciones:
  CLAIM-1: relativedelta importado a nivel de módulo (no dentro de función)
  CLAIM-2: regex pre-compiladas a nivel de módulo
"""

import re
import unicodedata
import math
import zipfile
from datetime import datetime
from pathlib import Path
import pandas as pd

_PREPOSICIONES  = {"de", "del", "los", "las", "y", "en", "el", "la", "por", "con", "a"}
_VALORES_VACIOS = {"", "---", "-", "nan", "nat", "none"}
_MESES = {
    1:"enero", 2:"febrero", 3:"marzo", 4:"abril", 5:"mayo", 6:"junio",
    7:"julio", 8:"agosto", 9:"septiembre", 10:"octubre",
    11:"noviembre", 12:"diciembre"
}

# CLAIM-2: expresiones regulares pre-compiladas
_ISO_RE          = re.compile(r'^\d{4}-\d{2}-\d{2}$')
_LATAM_RE        = re.compile(r'^\d{1,2}/\d{1,2}/\d{4}$')
_RUT_RE          = re.compile(r'\d{6,8}-[\dkK]')  # RUT chileno — curador real asignado
_PAREN_RE        = re.compile(r'\(.*?\)')
_PAREN_SUELTO_RE = re.compile(r'[()]')
_ESPACIOS_RE     = re.compile(r'\s+')
_GUIONES_RE      = re.compile(r'[-().]+')
_NO_LETRAS_RE    = re.compile(r'[^a-zA-ZáéíóúÁÉÍÓÚüÜñÑ]')
_SEP_PROG_RE     = re.compile(r'[\s\-]+')
_EXTENSIONES_EXCEL = {".xls", ".xlsx", ".xlsm"}
_MAX_FILAS_EXCEL = 50_000
_MAX_COLUMNAS_EXCEL = 100


# ── fechas ────────────────────────────────────────────────────────────────────

def fecha_es(fecha=None):
    """Retorna fecha en formato '13 de mayo de 2026'."""
    if fecha is None:
        h = datetime.now()
    else:
        try:
            h = pd.to_datetime(fecha)
        except (OverflowError, TypeError, ValueError):
            return str(fecha)
    return f"{h.day} de {_MESES[h.month]} de {h.year}"


# ── limpieza de texto ─────────────────────────────────────────────────────────

def limpiar_nombre(txt):
    """Elimina paréntesis y residuos, colapsa espacios, Title Case."""
    if not txt:
        return ""
    s = _PAREN_RE.sub('', str(txt))
    s = _PAREN_SUELTO_RE.sub('', s)
    s = _ESPACIOS_RE.sub(' ', s).strip()
    return " ".join(t.capitalize() for t in s.split())


def normalizar(txt):
    """Sin acentos, lowercase, strip. Para detecciones internas."""
    if not txt:
        return ""
    txt = unicodedata.normalize("NFKD", str(txt)).encode("ASCII", "ignore").decode()
    txt = txt.lower().strip()
    txt = _GUIONES_RE.sub(" ", txt)
    return _ESPACIOS_RE.sub(" ", txt).strip()


def normalizar_match(txt):
    """
    Normalización para cruce entre hojas.
    Limpia (), guard pd.isna() con try/except (protege contra listas/arrays).
    """
    try:
        if pd.isna(txt):
            return ""
    except (TypeError, ValueError):
        pass
    # G-08: conservar EXACTAMENTE la normalización histórica del cruce Hoja2
    # (sin quitar tildes ni guiones — cambiarla altera qué filas matchean).
    s = _PAREN_RE.sub('', str(txt))
    s = _PAREN_SUELTO_RE.sub('', s)
    return _ESPACIOS_RE.sub(' ', s.strip().lower())


def es_texto_formula(valor) -> bool:
    """True si openpyxl interpretaría el valor como fórmula al escribirlo.

    Solo el prefijo '=' se convierte en fórmula en un .xlsx generado por
    openpyxl. No se altera el valor (el viejo prefijo apóstrofe corrompía
    placeholders legítimos como '---'): el guardado fuerza data_type='s'.
    """
    return isinstance(valor, str) and valor.lstrip().startswith("=")


def _validar_dimensiones_xlsx(path: Path, *, max_filas: int, max_columnas: int) -> None:
    """Rechaza libros XLSX/XLSM con dimensiones incompatibles con el flujo RUS."""
    from openpyxl import load_workbook

    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        for hoja in workbook.worksheets:
            filas = hoja.max_row or 0
            columnas = hoja.max_column or 0
            if filas > max_filas or columnas > max_columnas:
                raise ValueError(
                    f"El libro Excel supera el máximo permitido "
                    f"({max_filas} filas / {max_columnas} columnas por hoja): "
                    f"hoja {hoja.title!r} tiene {filas} filas y {columnas} columnas"
                )
    finally:
        workbook.close()


def validar_archivo_excel(ruta, *, max_bytes=100 * 1024 * 1024,
                          max_descomprimido=200 * 1024 * 1024,
                          max_filas=_MAX_FILAS_EXCEL,
                          max_columnas=_MAX_COLUMNAS_EXCEL) -> Path:
    """Valida tipo y límites básicos antes de abrir un Excel seleccionado."""
    path = Path(ruta).expanduser()
    if not path.is_file():
        raise FileNotFoundError(f"Archivo Excel no encontrado: {path}")
    if path.suffix.lower() not in _EXTENSIONES_EXCEL:
        raise ValueError(f"Extensión Excel no permitida: {path.suffix or '(sin extensión)'}")
    if path.stat().st_size > max_bytes:
        raise ValueError("El archivo Excel supera el tamaño máximo permitido")

    if path.suffix.lower() in {".xlsx", ".xlsm"}:
        if not zipfile.is_zipfile(path):
            raise ValueError("El archivo no es un libro XLSX/XLSM válido")
        with zipfile.ZipFile(path) as archivo:
            miembros = archivo.infolist()
            if len(miembros) > 10_000:
                raise ValueError("El libro Excel contiene demasiados componentes")
            if sum(miembro.file_size for miembro in miembros) > max_descomprimido:
                raise ValueError("El contenido descomprimido del Excel supera el límite")
        _validar_dimensiones_xlsx(path, max_filas=max_filas, max_columnas=max_columnas)
    return path


def titulo_programa(txt):
    """Siglas MAYÚS + preposiciones lower + resto capitalize."""
    if not txt:
        return ""
    tokens = _SEP_PROG_RE.split(_ESPACIOS_RE.sub(" ", str(txt)).strip())
    result = []
    for i, t in enumerate(tokens):
        if not t:
            continue
        solo  = _NO_LETRAS_RE.sub("", t)
        t_low = normalizar(t)
        if i > 0 and t_low in _PREPOSICIONES:
            result.append(t.lower())
        elif len(solo) <= 4 and solo.isalpha() and t_low not in _PREPOSICIONES:
            result.append(t.upper())
        else:
            result.append(t.capitalize())
    return " ".join(result)


# ── prefijo y audiencia ───────────────────────────────────────────────────────

def prefijo_observacion(nombre, programa):
    """'PrimerNombre 3LETRAS: ' — fallback a 3LETRAS si nombre queda vacío."""
    nombre_limpio = limpiar_nombre(nombre)
    primer        = nombre_limpio.split()[0] if nombre_limpio.strip() else ""
    tres          = (unicodedata.normalize("NFKD", str(programa or ""))
                     .encode("ASCII", "ignore").decode().upper()[:3])
    if primer and tres:
        return f"{primer} {tres}: "
    elif primer:
        return f"{primer}: "
    elif tres:
        return f"{tres}: "
    return ""


def audiencia_suffix(row, cols):
    """Sufijo de audiencia. Filtra fechas pasadas."""
    col = cols.get("prox_aud")
    if not col:
        return ""
    fec = get_date(row.get(col))
    if not fec:
        return ""
    fec_date = fec.date() if hasattr(fec, "date") else fec
    if fec_date < datetime.now().date():
        return ""
    return f". Se cita a audiencia para el día {fecha_es(fec)}."


# ── detecciones ───────────────────────────────────────────────────────────────

def detectar_tribunal(valor):
    if not valor:
        return None
    t = normalizar(valor).upper()
    if "MULCHEN" in t: return "MULCHEN"
    if "LAJA"   in t: return "LAJA"
    if "TOME"   in t: return "TOME"
    return None


def contiene_token(texto, *tokens) -> bool:
    """Compara tokens normalizados por palabra completa."""
    norm = normalizar(texto)
    palabras = set(re.findall(r"[a-z0-9]+", norm))
    return any(normalizar(t) in palabras for t in tokens)


def es_dce(txt):
    if not txt: return False
    t = normalizar(txt).upper()
    return "DCE" in t or "DIAGNOSTICO" in t


def es_derivacion_sin_seg(txt):
    if not txt:
        return False
    norm = normalizar(txt)
    patrones = (
        "opd", "dam", "salud privada", "hospital", "unidad de salud mental",
        "cesfam", "red salud", "consulta externa", "colegio", "chile crece contigo",
    )
    return any(re.search(rf"^{re.escape(p)}(?:\b|$)", norm) for p in patrones)


def tiene_curador_real(val) -> bool:
    """
    True si el valor de la columna curador contiene un RUT chileno,
    o si registra un curador institucional (ej. CAJ Biobío) sin RUT individual.
    Cubre todos los placeholders de RUS: '', '---', 'NO POSEE', 'Sin designar', etc.
    Solo devuelve True cuando hay un curador efectivamente asignado.
    """
    s = str(val or '')
    return bool(_RUT_RE.search(s)) or ('Institución:' in s)


# ── edad ──────────────────────────────────────────────────────────────────────

def calcular_edad_exacta(fec_nacimiento):
    fec = get_date(fec_nacimiento)
    if not fec: return None
    hoy = datetime.now().date()
    nac = fec.date() if hasattr(fec, "date") else fec
    return hoy.year - nac.year - ((hoy.month, hoy.day) < (nac.month, nac.day))


def dias_para_mayoria(fec_nacimiento):
    fec = get_date(fec_nacimiento)
    if not fec: return None
    hoy = datetime.now().date()
    nac = fec.date() if hasattr(fec, "date") else fec
    try:
        mayoria = nac.replace(year=nac.year + 18)
    except ValueError:
        mayoria = nac.replace(year=nac.year + 18, day=28)
    return (mayoria - hoy).days


def fecha_mayoria(fec_nacimiento):
    fec = get_date(fec_nacimiento)
    if not fec: return None
    nac = fec.date() if hasattr(fec, "date") else fec
    try:
        return nac.replace(year=nac.year + 18)
    except ValueError:
        return nac.replace(year=nac.year + 18, day=28)


# ── get helpers ───────────────────────────────────────────────────────────────

def get_int(val):
    if val is None: return None
    s = str(val).strip().lower()
    if s in _VALORES_VACIOS: return None
    try:
        num = pd.to_numeric(s, errors="coerce")
        if pd.isna(num) or not math.isfinite(float(num)):
            return None
        if not float(num).is_integer():
            return None
        return int(num)
    except (OverflowError, ValueError, TypeError):
        return None


def get_date(val):
    """Parsing determinista: ISO → %Y-%m-%d, LATAM → %d/%m/%Y."""
    if val is None:
        return None
    if isinstance(val, (datetime, pd.Timestamp)):
        return val if isinstance(val, datetime) else val.to_pydatetime()
    if isinstance(val, (int, float)) and not isinstance(val, bool) and 30000 < val < 60000:
        return (pd.Timestamp('1899-12-30') + pd.Timedelta(days=val)).to_pydatetime()
    s = str(val).strip()
    if s.lower() in _VALORES_VACIOS:
        return None
    try:
        if _ISO_RE.match(s):
            return pd.to_datetime(s, format="%Y-%m-%d")
        if _LATAM_RE.match(s):
            return pd.to_datetime(s, format="%d/%m/%Y")
        return pd.to_datetime(s, dayfirst=True)
    except (ValueError, TypeError):
        return None
