# -*- coding: utf-8 -*-
"""
generador_correos.py — CSMP Assistant v8.0

Genera borradores Outlook (Drafts) vía win32com. NUNCA envía.

Tipos de borrador:
  A) PROGRAMA+TRIBUNAL con informes VENCIDOS      → 1 borrador al programa
  B) TRIBUNAL con ≥1 informe vencido              → 1 borrador al tribunal
  C) PROGRAMA+TRIBUNAL con informes POR VENCER    → 1 borrador al programa
  D) PROGRAMA+TRIBUNAL en LISTA DE ESPERA         → 1 borrador al programa

Formato saludo en TODOS los correos:
  SRES. NOMBRE PROGRAMA
  PRESENTE.
"""

import re
import unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd

from .contactos_programas import CatastroContactos

CC_FIJO = "ucc_concepcion@pjud.cl"
DIAS_POR_VENCER = 45

PATRON_ESPERA = "se remite correo electronico al programa consultando respecto de fecha estimada de ingreso"


# ─── Helpers nombre ───────────────────────────────────────────────────────────

def _formatear_nombre_programa(nombre: str) -> str:
    """
    'AFT - MULCHEN'              → 'AFT Mulchén'
    'RESIDENCIA HOGAR SAN PABLO' → 'Residencia Hogar San Pablo'
    'RFA – CASTELLON'            → 'RFA Castellón'
    Regla: tokens ≤4 chars + solo letras + todo mayúscula en posición 0-1 → sigla (MAYÚSCULA).
    Resto → capitalize().
    Guiones/rayas aislados se eliminan.
    """
    if not nombre:
        return nombre
    s = re.sub(r'\s*[-–—]\s*', ' ', nombre.strip())
    tokens = [t for t in s.split() if t]
    resultado = []
    for i, tok in enumerate(tokens):
        if i < 2 and len(tok) <= 4 and tok.isalpha() and tok.isupper():
            resultado.append(tok)          # sigla → MAYÚSCULA
        else:
            resultado.append(tok.capitalize())
    return ' '.join(resultado)


def _titulo_nombre(nombre: str) -> str:
    """'JUAN PEREZ LOPEZ ()' -> 'Juan Perez Lopez' (limpia parentesis)"""
    s = re.sub(r'\(.*?\)', '', str(nombre or ''))
    s = re.sub(r'[()]', '', s)
    return ' '.join(p.capitalize() for p in s.split())


def _normalizar_texto(txt: str) -> str:
    return unicodedata.normalize("NFKD", str(txt or "")).encode("ASCII", "ignore").decode().lower()


# ─── Helpers tribunal ─────────────────────────────────────────────────────────

def _detectar_clave_tribunal(valor: str) -> str | None:
    v = _normalizar_texto(valor).upper()
    if "MULCHEN" in v: return "MULCHEN"
    if "LAJA" in v:    return "LAJA"
    if "TOME" in v:    return "TOME"
    return None


# ─── Helpers HTML ─────────────────────────────────────────────────────────────

def _fila_fecha(rit, tribunal, rut, nombre, fecha) -> str:
    fecha_str  = pd.to_datetime(fecha, dayfirst=True).strftime("%d/%m/%Y") if pd.notna(fecha) else "---"
    rut_str    = str(rut).strip() if rut and str(rut).strip() not in ("", "nan", "None") else "---"
    nombre_fmt = _titulo_nombre(str(nombre))
    return (
        f'<tr>'
        f'<td style="padding:5px 10px;border:1px solid #ddd">{rit}</td>'
        f'<td style="padding:5px 10px;border:1px solid #ddd">{tribunal}</td>'
        f'<td style="padding:5px 10px;border:1px solid #ddd">{rut_str}</td>'
        f'<td style="padding:5px 10px;border:1px solid #ddd">{nombre_fmt}</td>'
        f'<td style="padding:5px 10px;border:1px solid #ddd;text-align:center">{fecha_str}</td>'
        f'</tr>'
    )

def _fila_espera(rit, tribunal, rut, nombre, dias_espera) -> str:
    rut_str    = str(rut).strip() if rut and str(rut).strip() not in ("", "nan", "None") else "---"
    nombre_fmt = _titulo_nombre(str(nombre))
    try:
        dias_str = str(int(float(str(dias_espera)))) if dias_espera not in (None, "", "nan") else "---"
    except Exception:
        dias_str = str(dias_espera)
    return (
        f'<tr>'
        f'<td style="padding:5px 10px;border:1px solid #ddd">{rit}</td>'
        f'<td style="padding:5px 10px;border:1px solid #ddd">{tribunal}</td>'
        f'<td style="padding:5px 10px;border:1px solid #ddd">{rut_str}</td>'
        f'<td style="padding:5px 10px;border:1px solid #ddd">{nombre_fmt}</td>'
        f'<td style="padding:5px 10px;border:1px solid #ddd;text-align:center">{dias_str}</td>'
        f'</tr>'
    )

def _bloque_programa_tribunal(programa: str, filas_html: str) -> str:
    prog_fmt = _formatear_nombre_programa(programa)
    return (
        f'<p style="margin-top:16px"><strong>{prog_fmt}</strong></p>'
        f'<table style="border-collapse:collapse;font-size:10pt;width:100%;max-width:820px">'
        f'<thead><tr style="background:#1F4E78;color:#fff">'
        f'<th style="padding:6px 10px;border:1px solid #ccc;text-align:left">RIT</th>'
        f'<th style="padding:6px 10px;border:1px solid #ccc;text-align:left">Tribunal</th>'
        f'<th style="padding:6px 10px;border:1px solid #ccc;text-align:left">RUT NNA</th>'
        f'<th style="padding:6px 10px;border:1px solid #ccc;text-align:left">Nombre NNA</th>'
        f'<th style="padding:6px 10px;border:1px solid #ccc;text-align:center">Fecha vencimiento</th>'
        f'</tr></thead>'
        f'<tbody>{filas_html}</tbody>'
        f'</table>'
    )

def _cargar_plantilla(nombre_archivo: str, **reemplazos) -> str:
    ruta = Path(__file__).parent / "plantillas" / nombre_archivo
    html = ruta.read_text(encoding="utf-8")
    for clave, valor in reemplazos.items():
        html = html.replace(f"{{{{{clave}}}}}", str(valor))
    return html


# ─── Outlook ──────────────────────────────────────────────────────────────────

def _crear_borrador_outlook(para: list[str], cc: list[str],
                             asunto: str, cuerpo_html: str) -> bool:
    try:
        import win32com.client as win32
        outlook = win32.Dispatch("Outlook.Application")
        mail    = outlook.CreateItem(0)
        mail.To = "; ".join(para)
        mail.CC = "; ".join(cc)
        mail.Subject = asunto

        # Obtener firma del usuario.
        # Método recomendado para Outlook moderno:
        #   1. Display(False) — abre inspector en background, inyecta firma automáticamente
        #   2. Leer HTMLBody (ya contiene la firma)
        #   3. Cerrar el inspector
        #   4. Reensamblar: firma + contenido del correo
        # Fallback si Display falla: inyectar solo el cuerpo sin firma.
        html_firma = ""
        inspector  = None
        try:
            mail.Display(False)           # abre en background, carga la firma
            inspector  = mail.GetInspector
            html_firma = mail.HTMLBody or ""
            inspector.Close(0)            # cerrar sin guardar (0 = olDiscardChanges)
            inspector = None
        except Exception:
            try:
                html_firma = mail.HTMLBody or ""
            except Exception:
                html_firma = ""

        interno  = _extraer_body(cuerpo_html)
        body_tag = re.search(r'<body[^>]*>', html_firma, re.IGNORECASE)
        if body_tag:
            pos = body_tag.end()
            mail.HTMLBody = html_firma[:pos] + "\n" + interno + "\n<br>" + html_firma[pos:]
        elif html_firma:
            mail.HTMLBody = cuerpo_html + html_firma
        else:
            mail.HTMLBody = cuerpo_html  # sin firma disponible

        mail.Save()
        return True
    except ImportError:
        raise RuntimeError("pywin32 no disponible. Instala: pip install pywin32")
    except Exception as e:
        raise RuntimeError(f"Error creando borrador Outlook: {e}")

def _extraer_body(html: str) -> str:
    m = re.search(r'<body[^>]*>(.*?)</body>', html, re.IGNORECASE | re.DOTALL)
    return m.group(1).strip() if m else html


# ─── Generador principal ──────────────────────────────────────────────────────

class GeneradorCorreos:
    def __init__(self, config: dict, ruta_catastro: str, contactos_tribunales: dict):
        self.cc_fijo        = config.get("cc_fijo", CC_FIJO)
        self.catastro       = CatastroContactos(ruta_catastro)
        self.contactos_trib = contactos_tribunales

    # ── Informes (vencidos / por vencer) ─────────────────────────────────────

    def procesar(self, df: pd.DataFrame) -> dict:
        """
        Excel informes. Cols: RIT, TRIBUNAL, RUT, NOMBRE, DERIVACIÓN, FECHA VENCIMIENTO.
        Genera tipos A, B, C.
        """
        resultado = self._vacio()
        df = df.copy()
        df.columns = [str(c).strip() for c in df.columns]

        col_rit    = self._col(df, ["RIT"])
        col_trib   = self._col(df, ["TRIBUNAL"])
        col_rut    = self._col(df, ["RUT", "RUT NNA", "RUT LITIGANTE"])
        col_nombre = self._col(df, ["NOMBRE", "NOMBRE COMPLETO"])
        col_prog   = self._col(df, ["DERIVACIÓN", "DERIVACION", "PROGRAMA"])
        col_fecha  = self._col(df, ["FECHA VENCIMIENTO", "FEC.VENCIMIENTO",
                                     "FEC. VENCIMIENTO"])

        if not all([col_rit, col_trib, col_nombre, col_prog, col_fecha]):
            faltantes = [n for n, c in [("RIT", col_rit), ("TRIBUNAL", col_trib),
                         ("NOMBRE", col_nombre), ("PROGRAMA", col_prog),
                         ("FECHA VENCIMIENTO", col_fecha)] if not c]
            resultado["errores"].append(f"Columnas no encontradas: {faltantes}")
            return resultado

        df[col_fecha] = pd.to_datetime(df[col_fecha], errors="coerce", dayfirst=True)
        hoy = datetime.now().date()
        df["_clave_trib"] = df[col_trib].apply(_detectar_clave_tribunal)
        df["_dias"] = df[col_fecha].apply(
            lambda f: (f.date() - hoy).days if pd.notna(f) else None)
        df["_tipo"] = df["_dias"].apply(
            lambda d: "VENCIDO" if d is not None and d <= 0
                      else ("POR_VENCER" if d is not None and 0 < d <= DIAS_POR_VENCER
                            else "FUTURO"))

        vencidos   = df[df["_tipo"] == "VENCIDO"]
        por_vencer = df[df["_tipo"] == "POR_VENCER"]

        for (prog, trib_raw), g in vencidos.groupby([col_prog, col_trib]):
            self._prog_borrador(g, prog, trib_raw, col_rit, col_trib, col_rut,
                                col_nombre, col_fecha, "vencidos", resultado)

        for clave, g in vencidos.groupby("_clave_trib"):
            if clave:
                self._trib_borrador(g, clave, col_prog, col_rit, col_trib,
                                    col_rut, col_nombre, col_fecha, resultado)

        for (prog, trib_raw), g in por_vencer.groupby([col_prog, col_trib]):
            self._prog_borrador(g, prog, trib_raw, col_rit, col_trib, col_rut,
                                col_nombre, col_fecha, "por_vencer", resultado)

        return resultado

    # ── Lista de espera ───────────────────────────────────────────────────────

    def procesar_espera(self, df: pd.DataFrame) -> dict:
        """
        Excel ESPERA (salida del motor con columna OBSERVACION).
        Genera borrador tipo D: 1 por PROGRAMA+TRIBUNAL donde OBSERVACION
        contiene el patrón de lista de espera.
        Cols: RIT, TRIBUNAL, RUT, NOMBRE, DERIVACIÓN, T ESPERA, OBSERVACION.
        """
        resultado = self._vacio()
        df = df.copy()
        df.columns = [str(c).strip() for c in df.columns]

        col_rit    = self._col(df, ["RIT"])
        col_trib   = self._col(df, ["TRIBUNAL"])
        col_rut    = self._col(df, ["RUT", "RUT NNA", "RUT LITIGANTE"])
        col_nombre = self._col(df, ["NOMBRE", "NOMBRE COMPLETO"])
        col_prog   = self._col(df, ["DERIVACIÓN", "DERIVACION", "PROGRAMA"])
        col_espera = self._col(df, ["T ESPERA", "T_ESPERA", "DIAS_ESPERA",
                                     "DÍAS DE ESPERA", "DIAS DE ESPERA", "TESPERA"])
        col_obs    = self._col(df, ["OBSERVACION", "OBSERVACIÓN"])

        faltantes = [n for n, c in [("RIT", col_rit), ("TRIBUNAL", col_trib),
                     ("NOMBRE", col_nombre), ("PROGRAMA", col_prog),
                     ("T ESPERA", col_espera)] if not c]
        if faltantes:
            resultado["errores"].append(f"Columnas no encontradas: {faltantes}")
            return resultado

        # Filtrar por observación
        if col_obs:
            df_filtrado = df[df[col_obs].apply(
                lambda obs: PATRON_ESPERA in _normalizar_texto(obs)
            )].copy()
        else:
            df_filtrado = df.copy()

        if df_filtrado.empty:
            resultado["errores"].append(
                "No hay filas con observación de lista de espera. "
                "Procesa primero con el motor (ESPERA)."
            )
            return resultado

        for (prog, trib_raw), grupo in df_filtrado.groupby([col_prog, col_trib]):
            contacto = self.catastro.resolver(str(prog))
            if not contacto:
                msg = f"Sin contacto: '{prog}'"
                if msg not in resultado["grupos_sin_contacto"]:
                    resultado["grupos_sin_contacto"].append(msg)
                continue

            nombre_prog_fmt = _formatear_nombre_programa(contacto["nombre"])
            filas_html = ""
            for _, row in grupo.iterrows():
                rut = str(row[col_rut]) if col_rut and col_rut in row.index else "---"
                filas_html += _fila_espera(
                    str(row[col_rit]), str(row[col_trib]),
                    rut, str(row[col_nombre]), row[col_espera]
                )

            cuerpo = _cargar_plantilla(
                "correo_programa_espera.html",
                NOMBRE_PROGRAMA=nombre_prog_fmt,
                TRIBUNAL=str(trib_raw).strip(),
                FILAS=filas_html,
            )
            asunto = f"Lista de espera — {nombre_prog_fmt} / {str(trib_raw).strip()}"

            try:
                _crear_borrador_outlook(
                    para=[contacto["mail"]],
                    cc=[self.cc_fijo],
                    asunto=asunto,
                    cuerpo_html=cuerpo,
                )
                resultado["borradores_creados"] += 1
                resultado["detalle"].append(
                    f"✓ [ESPERA] {nombre_prog_fmt} / {str(trib_raw).strip()} "
                    f"→ {contacto['mail']} ({len(grupo)} registros)"
                )
            except RuntimeError as e:
                resultado["errores"].append(str(e))

        return resultado

    # ── Borradores internos ───────────────────────────────────────────────────

    def _prog_borrador(self, grupo, prog, trib_raw, col_rit, col_trib, col_rut,
                        col_nombre, col_fecha, tipo, resultado):
        contacto = self.catastro.resolver(str(prog))
        if not contacto:
            msg = f"Sin contacto: '{prog}'"
            if msg not in resultado["grupos_sin_contacto"]:
                resultado["grupos_sin_contacto"].append(msg)
            return

        nombre_prog_fmt = _formatear_nombre_programa(contacto["nombre"])
        filas_html = ""
        for _, row in grupo.iterrows():
            rut = str(row[col_rut]) if col_rut and col_rut in row.index else "---"
            filas_html += _fila_fecha(str(row[col_rit]), str(row[col_trib]),
                                       rut, str(row[col_nombre]), row[col_fecha])

        plantilla = ("correo_programa_vencidos.html" if tipo == "vencidos"
                     else "correo_programa_por_vencer.html")
        asunto_txt = "vencidos" if tipo == "vencidos" else "por vencer"
        asunto = f"Informes de avance {asunto_txt} — {nombre_prog_fmt} / {str(trib_raw).strip()}"

        cuerpo = _cargar_plantilla(plantilla, NOMBRE_PROGRAMA=nombre_prog_fmt,
                                    TRIBUNAL=str(trib_raw).strip(), FILAS=filas_html)
        try:
            _crear_borrador_outlook(para=[contacto["mail"]], cc=[self.cc_fijo],
                                     asunto=asunto, cuerpo_html=cuerpo)
            resultado["borradores_creados"] += 1
            resultado["detalle"].append(
                f"✓ [{tipo.upper()}] {nombre_prog_fmt} / {str(trib_raw).strip()} "
                f"→ {contacto['mail']} ({len(grupo)} registros)"
            )
        except RuntimeError as e:
            resultado["errores"].append(str(e))

    def _trib_borrador(self, grupo_trib, clave_trib, col_prog, col_rit, col_trib,
                        col_rut, col_nombre, col_fecha, resultado):
        mails = self.contactos_trib.get(clave_trib, [])
        if not mails:
            resultado["errores"].append(f"Sin mails de tribunal: {clave_trib}")
            return

        bloques = ""
        for prog, g in grupo_trib.groupby(col_prog):
            filas = ""
            for _, row in g.iterrows():
                rut = str(row[col_rut]) if col_rut and col_rut in row.index else "---"
                filas += _fila_fecha(str(row[col_rit]), str(row[col_trib]),
                                     rut, str(row[col_nombre]), row[col_fecha])
            bloques += _bloque_programa_tribunal(str(prog).strip(), filas)

        nombre_trib = {"LAJA": "Jgdo. L. y G. de Laja",
                       "MULCHEN": "Jgdo. L. y G. de Mulchén",
                       "TOME": "Juzgado de Familia Tomé"}.get(clave_trib, clave_trib)

        cuerpo = _cargar_plantilla("correo_tribunal_vencidos.html", BLOQUES_PROGRAMA=bloques)
        asunto = f"Informes vencidos en RUS — {nombre_trib} ({len(grupo_trib)} registros)"

        try:
            _crear_borrador_outlook(para=mails, cc=[self.cc_fijo],
                                     asunto=asunto, cuerpo_html=cuerpo)
            resultado["borradores_creados"] += 1
            resultado["detalle"].append(
                f"✓ [TRIBUNAL] {nombre_trib} → {len(mails)} dest. ({len(grupo_trib)} registros)"
            )
        except RuntimeError as e:
            resultado["errores"].append(str(e))

    # ── Helpers ───────────────────────────────────────────────────────────────

    @staticmethod
    def _vacio() -> dict:
        return {"borradores_creados": 0, "grupos_sin_contacto": [],
                "errores": [], "detalle": []}

    @staticmethod
    def _col(df: pd.DataFrame, aliases: list[str]) -> str | None:
        for c in df.columns:
            for a in aliases:
                if _normalizar_texto(c) == _normalizar_texto(a):
                    return c
        return None
