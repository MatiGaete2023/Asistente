#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CSMP Assistant v8.0 — Fase 4: Generador de Resoluciones .docx
Formato: Arial 12 pt, justificado, interlineado 1.5.
1 RIT por página, Word consolidado.
"""

import re, unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ─── Fechas ───────────────────────────────────────────────────────────────────

MESES_ES = {1:"enero",2:"febrero",3:"marzo",4:"abril",5:"mayo",6:"junio",
            7:"julio",8:"agosto",9:"septiembre",10:"octubre",11:"noviembre",12:"diciembre"}

_ISO_FECHA_RE   = re.compile(r'^\d{4}-\d{1,2}-\d{1,2}')
_LATAM_FECHA_RE = re.compile(r'^\d{1,2}[/\-]\d{1,2}[/\-]\d{4}')

def _parse_fecha(fecha):
    """
    Parser robusto y determinista para fechas que llegan del Excel.
    - ISO (YYYY-MM-DD)         → formato explícito
    - LATAM (DD/MM/YYYY o DD-MM-YYYY) → dayfirst (día primero)
    - Objetos datetime/Timestamp → directo
    Evita el bug donde 3/6/2025 se interpretaba como 6 de marzo (formato US).
    Retorna datetime o None.
    """
    if fecha is None:
        return None
    if isinstance(fecha, (datetime, pd.Timestamp)):
        return fecha if isinstance(fecha, datetime) else fecha.to_pydatetime()
    s = str(fecha).strip()
    if not s or s.lower() in ("nan", "nat", "none", "---", "-"):
        return None
    try:
        if _ISO_FECHA_RE.match(s):
            return pd.to_datetime(s, dayfirst=False)
        if _LATAM_FECHA_RE.match(s):
            # Normalizar guiones a barras y forzar día primero
            return pd.to_datetime(s.replace("-", "/"), dayfirst=True)
        return pd.to_datetime(s, dayfirst=True)
    except (ValueError, TypeError):
        return None

_UNIDADES = ["","un","dos","tres","cuatro","cinco","seis","siete","ocho","nueve",
             "diez","once","doce","trece","catorce","quince","dieciséis",
             "diecisiete","dieciocho","diecinueve"]
_DECENAS  = ["","","veinte","treinta","cuarenta","cincuenta",
             "sesenta","setenta","ochenta","noventa"]
_CENTENAS = ["","ciento","doscientos","trescientos","cuatrocientos","quinientos",
             "seiscientos","setecientos","ochocientos","novecientos"]

def _num_palabras(n: int) -> str:
    if n == 0: return "cero"
    if n == 100: return "cien"
    if n == 1000: return "mil"
    res = []
    miles = n // 1000
    if miles > 0:
        res.append("mil" if miles == 1 else f"{_num_palabras(miles)} mil")
    resto = n % 1000
    c = resto // 100
    if c > 0: res.append(_CENTENAS[c])
    d = resto % 100
    if d > 0:
        if d < 20:
            res.append(_UNIDADES[d])
        else:
            dec, uni = d // 10, d % 10
            if uni == 0:
                res.append(_DECENAS[dec])
            elif dec == 2:
                _veinti = {1:"veintiuno", 2:"veintidós", 3:"veintitrés", 4:"veinticuatro",
                           5:"veinticinco", 6:"veintiséis", 7:"veintisiete",
                           8:"veintiocho", 9:"veintinueve"}
                res.append(_veinti[uni])
            else:
                res.append(f"{_DECENAS[dec]} y {_UNIDADES[uni]}")
    return " ".join(res)

def fecha_en_palabras(fecha=None) -> str:
    """'quince de mayo de dos mil veintiséis'"""
    h = datetime.now() if fecha is None else _parse_fecha(fecha)
    if h is None:
        return "COMPLETAR"
    return f"{_num_palabras(h.day)} de {MESES_ES[h.month]} de {_num_palabras(h.year)}"

def fecha_numerica(fecha=None) -> str:
    """'15 de mayo de 2026'"""
    h = datetime.now() if fecha is None else _parse_fecha(fecha)
    if h is None:
        return "COMPLETAR"
    return f"{h.day} de {MESES_ES[h.month]} de {h.year}"

def duracion_limpia(raw) -> str:
    """'1 año(s)' → '1 año', '6 mes(es)' → '6 meses'"""
    s = re.sub(r'\([^)]*\)', '', str(raw or "")).strip()
    m = re.match(r'^(\d+)\s+mes$', s)
    if m:
        n = int(m.group(1))
        return f"{n} mes" if n == 1 else f"{n} meses"
    return s

def normalizar(txt: str) -> str:
    txt = unicodedata.normalize("NFKD", str(txt or "")).encode("ASCII","ignore").decode()
    return txt.lower().strip()

def _titulo_nombre(nombre: str) -> str:
    """'JUAN PEREZ LOPEZ ()' → 'Juan Perez Lopez' (limpia paréntesis y contenido)."""
    s = re.sub(r'\(.*?\)', '', str(nombre or ''))
    s = re.sub(r'[()]', '', s)
    return ' '.join(p.capitalize() for p in s.split())

def _titulo_programa(programa: str) -> str:
    """
    'AFT - MULCHEN'              → 'AFT Mulchén'
    'RESIDENCIA HOGAR SAN PABLO' → 'Residencia Hogar San Pablo'
    Tokens ≤4 chars + solo letras + posición 0-1 → sigla MAYÚSCULA. Resto capitalize().
    Guiones/rayas aislados se eliminan.
    """
    s = re.sub(r'\s*[-–—]\s*', ' ', str(programa or '').strip())
    tokens = [t for t in s.split() if t]
    resultado = []
    for i, tok in enumerate(tokens):
        if i < 2 and len(tok) <= 4 and tok.isalpha() and tok.isupper():
            resultado.append(tok)
        else:
            resultado.append(tok.capitalize())
    return ' '.join(resultado)

# ─── Detección tipo resolución ────────────────────────────────────────────────

_PATRON_PC_IE   = ["pidiendo cuenta a programa respecto al ie",
                   "pide cuenta al programa",
                   "proyecto de resolucion pidiendo cuenta a programa"]
_PATRON_PC_INFO = ["pidiendo cuenta del informe de avance pendiente",
                   "informe pendiente de entrega",
                   "proyecto de resolucion pidiendo cuenta del informe"]
_PATRON_NOMENCL = ["aplica nomenclaturas", "regularizar informaticamente",
                   "nomenclaturas a fin de regularizar"]

def detectar_tipo(obs: str) -> str | None:
    n = normalizar(obs)
    for p in _PATRON_PC_INFO:
        if p in n: return "PC_INFO"
    for p in _PATRON_PC_IE:
        if p in n: return "PC_IE"
    for p in _PATRON_NOMENCL:
        if p in n: return "NOMENCL"
    return None

def detectar_tribunal(val: str) -> str | None:
    t = normalizar(val).upper()
    if "MULCHEN" in t: return "MULCHEN"
    if "LAJA" in t:    return "LAJA"
    if "TOME" in t:    return "TOME"
    return None

def extraer_ncl(obs: str) -> str:
    n = normalizar(obs)
    if "ingreso efectivo" in n:  return "Ingreso Efectivo"
    if "informe" in n:           return "Ingreso informe cumplimiento X"
    if "prorrog" in n:           return "Prorroga"
    if "egreso" in n:            return "Egreso Serv. Protec. Especializada y otras redes"
    return "NCL — COMPLETAR MANUALMENTE"

# ─── Mapeo columnas ───────────────────────────────────────────────────────────

_COL_ALIASES = {
    "tribunal":   ["TRIBUNAL"],
    "rit":        ["RIT"],
    "nombre":     ["NOMBRE","NOMBRE COMPLETO"],
    "rut":        ["RUT"],
    "derivacion": ["DERIVACION","DERIVACIÓN","PROGRAMA"],
    "duracion":   ["DURACION","DURACIÓN","PLAZO","VIGENCIA"],
    "fec_res":    ["FEC. RESOLUCIÓN","FEC.RESOLUCION","FECHA RESOLUCIÓN","FECHA RESOLUCION"],
    "observacion":["OBSERVACION","OBSERVACIÓN"],
}

def _get_col(df, key):
    for c in df.columns:
        for a in _COL_ALIASES.get(key, []):
            if normalizar(c) == normalizar(a): return c
    return None

def _mapear(df):
    return {k: _get_col(df, k) for k in _COL_ALIASES}

# ─── Helpers formato Word ─────────────────────────────────────────────────────

def _aplicar_formato(run, bold=False, italic=False):
    """Arial 12 pt estándar."""
    run.bold = bold
    run.italic = italic
    run.font.name = "Arial"
    run.font.size = Pt(12)

def _nuevo_parrafo(doc, texto="", bold=False, italic=False,
                   align=WD_ALIGN_PARAGRAPH.JUSTIFY,
                   antes=0, despues=0) -> object:
    """
    Crea párrafo con formato base:
      - Arial 12
      - Justificado
      - Interlineado 1.5
    """
    p = doc.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    pf.space_before = Pt(antes)
    pf.space_after  = Pt(despues)
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE

    if texto:
        r = p.add_run(texto)
        _aplicar_formato(r, bold=bold, italic=italic)

    return p

def _parrafo_mixto(doc, partes: list,
                   align=WD_ALIGN_PARAGRAPH.JUSTIFY,
                   antes=0, despues=0) -> object:
    """
    Crea párrafo con runs de formato mixto.
    partes = [(texto, bold, italic), ...]
    """
    p = doc.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    pf.space_before = Pt(antes)
    pf.space_after  = Pt(despues)
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE

    for texto, bold, italic in partes:
        r = p.add_run(texto)
        _aplicar_formato(r, bold=bold, italic=italic)
    return p

# ─── Encabezados y pies ───────────────────────────────────────────────────────

def _header_mulchen(doc):
    """Encabezado Mulchén — sin datos institucionales (dirección/fono/correo)."""
    # Las líneas de dirección, fono, correo y atención virtual fueron eliminadas
    # a solicitud del usuario. El documento comienza directamente con la ciudad/fecha.
    pass

def _header_laja_nota(doc):
    """Nota inicial de Laja en negrita."""
    _nuevo_parrafo(doc,
        "Laja, la fecha de la resolución es la que se indica en el pie de firma.",
        bold=True)
    doc.add_paragraph()

def _pie_mulchen(doc, fecha_palabras: str):
    _nuevo_parrafo(doc)
    _nuevo_parrafo(doc,
        "Se proveyó y firmó mediante firma electrónica avanzada, según lo dispuesto "
        "en la ley 20.886 y acta 71-2016 de la Excelentísima Corte Suprema.",
        bold=True)
    _nuevo_parrafo(doc,
        f"En Mulchén, {fecha_palabras}, notifiqué por el estado diario la resolución precedente.")

def _pie_laja(doc):
    _nuevo_parrafo(doc)
    _nuevo_parrafo(doc,
        "Proveyó, Juez del juzgado de Letras y Garantía de Laja, "
        "quien suscribe con firma electrónica avanzada.")
    _nuevo_parrafo(doc,
        "En Laja, con esta fecha, notifique por el estado diario la resolución que antecede. "
        "//cmga_csmp")

# ─── Generadores por tipo/tribunal ───────────────────────────────────────────
# Convención: datos extraídos del Excel van en NEGRITA y con título (capitalize).
# B = True (bold), F = False (normal), I = True (italic)
B, F, I = True, False, True

def _gen_pc_ie_mulchen(doc, d: dict, fp: str):
    nom  = _titulo_nombre(d['nombre'])
    prog = _titulo_programa(d['derivacion'])
    _header_mulchen(doc)
    _nuevo_parrafo(doc, f"Mulchén, {fp}.")
    _nuevo_parrafo(doc)
    _parrafo_mixto(doc, [
        ("Advirtiendo el Tribunal que con fecha ", F, F),
        (d['fec_res'], B, F),
        (" se ordenó el ingreso de ", F, F),
        (nom, B, F),
        (", cédula de identidad N° ", F, F),
        (d['rut'], B, F),
        (" a ", F, F),
        (prog, B, F),
        (" por el plazo de ", F, F),
        (d['duracion'], B, F),
        (", no constando en la carpeta digital el ingreso efectivo de la persona "
         "aludida a la institución, pídase cuenta al programa para que informe con "
         "carácter de urgente, en el plazo de 5 días, bajo apercibimiento del artículo "
         "238 del Código de Procedimiento Civil, si se efectuó el ingreso del niño "
         "sujeto de protección y, en la afirmativa, su fecha.", F, F),
    ])
    _nuevo_parrafo(doc)
    _nuevo_parrafo(doc, "Notifíquese al programa interventor por correo electrónico.")
    _nuevo_parrafo(doc)
    _nuevo_parrafo(doc, "Sirva la presente resolución como suficiente y atento oficio remisor.")
    _nuevo_parrafo(doc)
    _nuevo_parrafo(doc, f"RIT: {d['rit']}", bold=True)
    _pie_mulchen(doc, fp)

def _gen_pc_ie_laja(doc, d: dict):
    nom  = _titulo_nombre(d['nombre'])
    prog = _titulo_programa(d['derivacion'])
    _header_laja_nota(doc)
    _nuevo_parrafo(doc, "VISTO:", bold=True)
    _nuevo_parrafo(doc)
    _nuevo_parrafo(doc,
        "Atendido al mérito de los antecedentes que obran en autos y de conformidad a lo dispuesto "
        "en los artículos 13 y 16 de la Ley N° 19.968 que crea Los Tribunales de Familia, se resuelve:")
    _nuevo_parrafo(doc)
    _parrafo_mixto(doc, [
        ("Ofíciese a programa ", F, F),
        (prog, B, F),
        (" a fin de pedir cuenta del ingreso efectivo de ", F, F),
        (nom, B, F),
        (", cédula de identidad N° ", F, F),
        (d['rut'], B, F),
        (".", F, F),
    ])
    _nuevo_parrafo(doc)
    _nuevo_parrafo(doc, "-Notifíquese vía correo al programa interventor y curador.")
    _nuevo_parrafo(doc, "-Notifíquese a los restantes intervinientes por el estado diario.")
    _nuevo_parrafo(doc)
    _nuevo_parrafo(doc, "Sirva la presente resolución de suficiente y atento oficio remisor.")
    _nuevo_parrafo(doc)
    _nuevo_parrafo(doc, f"RIT: {d['rit']}", bold=True)
    _pie_laja(doc)

def _gen_pc_informe_mulchen(doc, d: dict, fp: str):
    nom  = _titulo_nombre(d['nombre'])
    prog = _titulo_programa(d['derivacion'])
    _header_mulchen(doc)
    _nuevo_parrafo(doc, f"Mulchén, {fp}.")
    _nuevo_parrafo(doc)
    _nuevo_parrafo(doc,
        "Atendido que el plazo para la remisión del informe de avance pertinente se encuentra vencido "
        "y teniendo presente las facultades oficiosas del Tribunal contempladas en el artículo 13 "
        "de la Ley 19.968, se resuelve:")
    _nuevo_parrafo(doc)
    _parrafo_mixto(doc, [
        ("Pídase cuenta al organismo interventor, ", F, F),
        (prog, B, F),
        (", a fin de dar estricto cumplimiento a lo ordenado en la presente causa, "
         "en cuanto a la remisión del informe de cumplimiento, correspondiente a ", F, F),
        (nom, B, F),
        (", cédula de identidad N° ", F, F),
        (d['rut'], B, F),
        (".", F, F),
    ])
    _nuevo_parrafo(doc)
    _nuevo_parrafo(doc,
        "Profesionales de dicho programa, deberán remitir dicho informe y dar cumplimiento en el "
        "término de 5 días desde su notificación o solicitar una prórroga para su entrega en caso "
        "de ser necesario.")
    _nuevo_parrafo(doc)
    _nuevo_parrafo(doc, "Notifíquese la presente resolución por correo electrónico al programa interventor.")
    _nuevo_parrafo(doc)
    _nuevo_parrafo(doc, "Sirva la presente resolución de suficiente y atento oficio remisor.")
    _nuevo_parrafo(doc)
    _nuevo_parrafo(doc, f"RIT: {d['rit']}", bold=True)
    _pie_mulchen(doc, fp)

def _gen_nomenclatura_mulchen(doc, d: dict, fp: str, ncl: str):
    _header_mulchen(doc)
    _nuevo_parrafo(doc, f"Mulchén, {fp}.")
    _nuevo_parrafo(doc)
    _parrafo_mixto(doc, [
        ("Advirtiendo el Tribunal que se omitió la incorporación de la nomenclatura \"", F, I),
        (ncl, B, I),
        ("\", aplíquese a la presente resolución para efectos de regularizar el correcto "
         "registro en el sistema de tramitación de causas.", F, I),
    ])
    _nuevo_parrafo(doc)
    _nuevo_parrafo(doc, f"RIT: {d['rit']}", bold=True)
    _pie_mulchen(doc, fp)

def _gen_nomenclatura_laja(doc, d: dict, ncl: str):
    _header_laja_nota(doc)
    _parrafo_mixto(doc, [
        ("Advirtiendo el Tribunal que se omitió la incorporación de la nomenclatura \"", F, I),
        (ncl, B, I),
        ("\", aplíquese a la presente resolución, para efectos de regularizar el correcto "
         "registro en el sistema de seguimiento de tramitación de causas.", F, I),
    ])
    _nuevo_parrafo(doc)
    _nuevo_parrafo(doc, f"RIT: {d['rit']}", bold=True)
    _pie_laja(doc)

# ─── Configurar estilo base del documento ────────────────────────────────────

def _configurar_estilos(doc: Document):
    """Establece Arial 12 interlineado 1.5 como estilo Normal del documento."""
    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = Pt(12)
    pf = style.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    pf.space_before = Pt(0)
    pf.space_after  = Pt(0)
    # Forzar la fuente también en el elemento XML para mayor compatibilidad
    rpr = style.element.get_or_add_rPr()
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:ascii"),    "Arial")
    rFonts.set(qn("w:hAnsi"),    "Arial")
    rFonts.set(qn("w:cs"),       "Arial")
    rpr.insert(0, rFonts)

# ─── Orquestador principal ────────────────────────────────────────────────────

def generar_resoluciones(ruta_excel: str, ruta_salida: str) -> dict:
    """
    Lee Excel con columna OBSERVACION (salida del motor),
    detecta filas que requieren proyecto de resolución y genera
    un Word consolidado (1 RIT por página).

    Retorna:
        {
            "archivo_generado": str | None,
            "total_resoluciones": int,
            "faltantes": list[dict],
            "errores": list[str],
        }
    """
    resultado = {"archivo_generado": None, "total_resoluciones": 0,
                 "faltantes": [], "errores": []}

    try:
        df = pd.read_excel(ruta_excel)
        df.columns = [str(c).strip() for c in df.columns]
        df = df.dropna(how="all").fillna("")
    except Exception as e:
        resultado["errores"].append(f"Error cargando Excel: {e}")
        return resultado

    cols = _mapear(df)
    if not cols["observacion"]:
        resultado["errores"].append(
            "Columna OBSERVACION no encontrada. Ejecuta el motor primero."
        )
        return resultado

    fecha_hoy_palabras = fecha_en_palabras()

    doc = Document()
    _configurar_estilos(doc)
    for section in doc.sections:
        section.top_margin    = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin   = Inches(1.1)
        section.right_margin  = Inches(1.1)

    resoluciones_generadas = 0
    primera = True

    for idx, row in df.iterrows():
        obs = str(row.get(cols["observacion"], "")).strip()
        tipo = detectar_tipo(obs)
        if not tipo:
            continue

        tribunal_raw = str(row.get(cols["tribunal"] or "", "")).strip()
        tribunal     = detectar_tribunal(tribunal_raw)

        d = {
            "rit":       str(row.get(cols["rit"] or "", "")).strip() or "COMPLETAR",
            "nombre":    str(row.get(cols["nombre"] or "", "")).strip() or "COMPLETAR",
            "rut":       str(row.get(cols["rut"] or "", "")).strip() or "COMPLETAR",
            "derivacion":str(row.get(cols["derivacion"] or "", "")).strip() or "COMPLETAR",
            "duracion":  duracion_limpia(row.get(cols["duracion"] or "", "")),
            "fec_res":   (fecha_numerica(row.get(cols["fec_res"], None))
                          if cols["fec_res"] and str(row.get(cols["fec_res"], "")).strip()
                          else "COMPLETAR"),
        }

        if not primera:
            doc.add_page_break()
        primera = False
        resoluciones_generadas += 1

        # Sin plantilla (Tomé u otro)
        if tribunal not in ("LAJA", "MULCHEN"):
            resultado["faltantes"].append({
                "rit": d["rit"], "tribunal": tribunal_raw,
                "tipo": tipo, "fila_excel": idx + 2,
            })
            _nuevo_parrafo(doc, f"[SIN PLANTILLA — {tribunal_raw or 'DESCONOCIDO'}]", bold=True)
            _nuevo_parrafo(doc, f"RIT: {d['rit']}  |  Tipo: {tipo}")
            _nuevo_parrafo(doc, "Completa este bloque manualmente.", italic=True)
            continue

        try:
            if tipo == "PC_IE":
                if tribunal == "MULCHEN": _gen_pc_ie_mulchen(doc, d, fecha_hoy_palabras)
                else:                     _gen_pc_ie_laja(doc, d)

            elif tipo == "PC_INFO":
                if tribunal == "MULCHEN":
                    _gen_pc_informe_mulchen(doc, d, fecha_hoy_palabras)
                else:
                    resultado["faltantes"].append({
                        "rit": d["rit"], "tribunal": "LAJA", "tipo": "PC_INFO",
                        "fila_excel": idx + 2, "nota": "Sin plantilla PC_Informe para Laja",
                    })
                    _nuevo_parrafo(doc, "[SIN PLANTILLA — PC_INFO LAJA]", bold=True)
                    _nuevo_parrafo(doc, f"RIT: {d['rit']}")

            elif tipo == "NOMENCL":
                ncl = extraer_ncl(obs)
                if tribunal == "MULCHEN": _gen_nomenclatura_mulchen(doc, d, fecha_hoy_palabras, ncl)
                else:                     _gen_nomenclatura_laja(doc, d, ncl)

        except Exception as e:
            resultado["errores"].append(f"Fila {idx+2} RIT {d['rit']}: {e}")
            _nuevo_parrafo(doc, f"[ERROR: {e}]", bold=True)

    if resoluciones_generadas == 0:
        resultado["errores"].append("No se detectaron filas con proyecto de resolución.")
        return resultado

    Path(ruta_salida).mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre = f"Resoluciones_{ts}.docx"
    ruta_completa = Path(ruta_salida) / nombre

    try:
        doc.save(str(ruta_completa))
        resultado["archivo_generado"] = str(ruta_completa)
        resultado["total_resoluciones"] = resoluciones_generadas
    except Exception as e:
        resultado["errores"].append(f"Error guardando Word: {e}")

    return resultado


def generar_informe_faltantes(faltantes: list, ruta_salida: str) -> str | None:
    """Excel con filas sin plantilla."""
    if not faltantes: return None
    df = pd.DataFrame(faltantes)
    Path(ruta_salida).mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    ruta = Path(ruta_salida) / f"Resoluciones_FALTANTES_{ts}.xlsx"
    df.to_excel(str(ruta), index=False)
    return str(ruta)
