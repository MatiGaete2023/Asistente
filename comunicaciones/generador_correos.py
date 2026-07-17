# -*- coding: utf-8 -*-
"""
generador_correos.py — CSMP Assistant v9.0.1

Genera borradores Outlook (Drafts) vía win32com. NUNCA envía.

Tipos de borrador:
  A) PROGRAMA+TRIBUNAL con informes VENCIDOS      → 1 borrador al programa
  B) PROGRAMA+TRIBUNAL con informes POR VENCER    → 1 borrador al programa
  C) PROGRAMA+TRIBUNAL en LISTA DE ESPERA         → 1 borrador al programa

Formato saludo en TODOS los correos:
  SRES. NOMBRE PROGRAMA
  PRESENTE.
"""

import re
from datetime import datetime
from html import escape
from pathlib import Path
import os
import uuid

import pandas as pd

from .contactos_programas import CatastroContactos
from motor.columnas_comunes import (
    normalizar, detectar_tribunal, titulo_programa, es_dce,
    ALIAS_RIT, ALIAS_TRIBUNAL, ALIAS_RUT, ALIAS_NOMBRE, ALIAS_PROGRAMA,
    ALIAS_VENCIMIENTO, ALIAS_ESPERA, TRIBUNAL_DISPLAY,
)

CC_FIJO = "ucc_concepcion@pjud.cl"
DIAS_POR_VENCER = 30

PATRON_ESPERA = "se remite correo electronico al programa consultando respecto de la fecha estimada de ingreso efectivo"


# ─── Helpers nombre ───────────────────────────────────────────────────────────

def _titulo_nombre(nombre: str) -> str:
    """'JUAN PEREZ LOPEZ ()' -> 'Juan Perez Lopez' (limpia parentesis)"""
    s = re.sub(r'\(.*?\)', '', str(nombre or ''))
    s = re.sub(r'[()]', '', s)
    return ' '.join(p.capitalize() for p in s.split())



# ─── Helpers tribunal ─────────────────────────────────────────────────────────


# ─── Helpers HTML ─────────────────────────────────────────────────────────────

def _texto_html(valor, default="---") -> str:
    try:
        if pd.isna(valor):
            return escape(default, quote=True)
    except (TypeError, ValueError):
        pass
    texto = str(valor).strip() if valor is not None else ""
    if texto.lower() in ("", "nan", "none", "nat"):
        texto = default
    return escape(texto, quote=True)


def _fila_fecha(rit, tribunal, rut, nombre, fecha) -> str:
    fecha_val = pd.to_datetime(fecha, errors="coerce", dayfirst=True)
    fecha_str = fecha_val.strftime("%d/%m/%Y") if pd.notna(fecha_val) else "---"
    rut_str = _texto_html(rut)
    nombre_fmt = escape(_titulo_nombre(str(nombre)), quote=True)
    return (
        f'<tr>'
        f'<td style="padding:5px 10px;border:1px solid #ddd">{_texto_html(rit)}</td>'
        f'<td style="padding:5px 10px;border:1px solid #ddd">{_texto_html(tribunal)}</td>'
        f'<td style="padding:5px 10px;border:1px solid #ddd">{rut_str}</td>'
        f'<td style="padding:5px 10px;border:1px solid #ddd">{nombre_fmt}</td>'
        f'<td style="padding:5px 10px;border:1px solid #ddd;text-align:center">{escape(fecha_str)}</td>'
        f'</tr>'
    )

def _fila_espera(rit, tribunal, rut, nombre, dias_espera) -> str:
    rut_str = _texto_html(rut)
    nombre_fmt = escape(_titulo_nombre(str(nombre)), quote=True)
    try:
        dias_str = str(int(float(str(dias_espera)))) if dias_espera not in (None, "", "nan") else "---"
    except Exception:
        dias_str = str(dias_espera)
    return (
        f'<tr>'
        f'<td style="padding:5px 10px;border:1px solid #ddd">{_texto_html(rit)}</td>'
        f'<td style="padding:5px 10px;border:1px solid #ddd">{_texto_html(tribunal)}</td>'
        f'<td style="padding:5px 10px;border:1px solid #ddd">{rut_str}</td>'
        f'<td style="padding:5px 10px;border:1px solid #ddd">{nombre_fmt}</td>'
        f'<td style="padding:5px 10px;border:1px solid #ddd;text-align:center">{_texto_html(dias_str)}</td>'
        f'</tr>'
    )

def _bloque_programa_tribunal(programa: str, filas_html: str) -> str:
    prog_fmt = escape(titulo_programa(programa), quote=True)
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
        seguro = str(valor) if clave == "FILAS" else escape(str(valor), quote=True)
        html = html.replace(f"{{{{{clave}}}}}", seguro)
    if re.search(r"\{\{[A-Z_]+\}\}", html):
        raise ValueError(f"Plantilla incompleta: {nombre_archivo}")
    return html


# ─── Outlook ──────────────────────────────────────────────────────────────────

def exportar_borrador_html(borrador: dict, ruta_salida: str) -> bool:
    """Exporta un borrador a HTML para revisión sin Outlook."""
    destino = Path(ruta_salida).expanduser()
    destino.mkdir(parents=True, exist_ok=True)
    asunto = str(borrador.get("asunto", "borrador")).strip() or "borrador"
    base = re.sub(r"[^A-Za-z0-9áéíóúÁÉÍÓÚñÑüÜ_-]+", "_", asunto).strip("_")
    base = (base or "borrador")[:80]
    archivo = destino / f"{base}_{datetime.now():%Y%m%d_%H%M%S_%f}.html"
    temporal = archivo.with_name(f".{archivo.stem}.{uuid.uuid4().hex}.tmp.html")
    cuerpo = str(borrador.get("cuerpo_html", ""))
    encabezado = (
        "<!DOCTYPE html><html lang=\"es\"><head><meta charset=\"UTF-8\">"
        f"<title>{escape(asunto, quote=True)}</title></head><body>"
        f"<p><strong>Para:</strong> {escape('; '.join(borrador.get('para', [])), quote=True)}<br>"
        f"<strong>CC:</strong> {escape('; '.join(borrador.get('cc', [])), quote=True)}<br>"
        f"<strong>Asunto:</strong> {escape(asunto, quote=True)}</p><hr>"
    )
    try:
        with open(temporal, "w", encoding="utf-8") as f:
            f.write(encabezado)
            f.write(cuerpo)
            f.write("</body></html>\n")
            f.flush()
            os.fsync(f.fileno())
        os.replace(temporal, archivo)
        return True
    except Exception:
        temporal.unlink(missing_ok=True)
        raise


def crear_exportador_html(ruta_salida: str):
    """Retorna un despachador compatible con GeneradorCorreos para dry-run."""
    return lambda borrador: exportar_borrador_html(borrador, ruta_salida)


def _crear_borrador_outlook(para: list[str], cc: list[str],
                             asunto: str, cuerpo_html: str) -> bool:
    pythoncom = None
    com_inicializado = False
    inspector = None
    guardado = False
    try:
        import pythoncom
        import win32com.client as win32
        pythoncom.CoInitialize()
        com_inicializado = True
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
        try:
            mail.Display(False)           # Outlook inyecta la firma al abrir el inspector
            inspector = mail.GetInspector
            html_firma = mail.HTMLBody or ""
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
        guardado = True
        return True
    except ImportError as exc:
        raise RuntimeError("pywin32 no disponible. Instala: pip install pywin32") from exc
    except Exception as e:
        raise RuntimeError(f"Error creando borrador Outlook: {e}") from e
    finally:
        if inspector is not None:
            try:
                # OlInspectorClose: 0=olSave, 1=olDiscard.
                inspector.Close(0 if guardado else 1)
            except Exception:
                pass
        if pythoncom is not None and com_inicializado:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

def _extraer_body(html: str) -> str:
    m = re.search(r'<body[^>]*>(.*?)</body>', html, re.IGNORECASE | re.DOTALL)
    return m.group(1).strip() if m else html


# ─── Generador principal ──────────────────────────────────────────────────────

class GeneradorCorreos:
    def __init__(self, config: dict, ruta_catastro: str, despachador=None):
        # Política institucional: el CC no se cambia desde configuración local.
        self.cc_fijo        = CC_FIJO
        self.catastro       = CatastroContactos(ruta_catastro)
        self._despachador   = despachador or self._crear_borrador

    # ── Informes (vencidos / por vencer) ─────────────────────────────────────

    def procesar(self, df: pd.DataFrame) -> dict:
        """
        Excel informes. Cols: RIT, TRIBUNAL, RUT, NOMBRE, DERIVACIÓN, FECHA VENCIMIENTO.
        Genera borradores por programa, tribunal y estado.
        """
        resultado = self._vacio()
        df = df.copy()
        df.columns = [str(c).strip() for c in df.columns]

        col_rit    = self._col(df, ALIAS_RIT)
        col_trib   = self._col(df, ALIAS_TRIBUNAL)
        col_rut    = self._col(df, ALIAS_RUT)
        col_nombre = self._col(df, ALIAS_NOMBRE)
        col_prog   = self._col(df, ALIAS_PROGRAMA)
        col_fecha  = self._col(df, ALIAS_VENCIMIENTO)

        if not all([col_rit, col_trib, col_nombre, col_prog, col_fecha]):
            faltantes = [n for n, c in [("RIT", col_rit), ("TRIBUNAL", col_trib),
                         ("NOMBRE", col_nombre), ("PROGRAMA", col_prog),
                         ("FECHA VENCIMIENTO", col_fecha)] if not c]
            resultado["errores"].append(f"Columnas no encontradas: {faltantes}")
            return resultado

        # Degradación suave (decisión D9 del plan): las filas con datos
        # faltantes/ inválidos se EXCLUYEN con aviso, pero el resto del lote
        # sigue generando borradores. Un lote diario nunca se aborta entero
        # por una fila defectuosa.
        df = self._excluir_filas_incompletas(df, {
            "RIT": col_rit,
            "TRIBUNAL": col_trib,
            "NOMBRE": col_nombre,
            "PROGRAMA": col_prog,
            "FECHA VENCIMIENTO": col_fecha,
        }, resultado)
        if df.empty:
            return resultado

        fechas = pd.to_datetime(df[col_fecha], errors="coerce", dayfirst=True)
        invalidas = int(fechas.isna().sum())
        if invalidas:
            resultado["errores"].append(
                f"FECHA VENCIMIENTO inválida en {invalidas} fila(s); esas filas se omitieron"
            )
            df = df[fechas.notna()].copy()
            fechas = fechas[fechas.notna()]
            if df.empty:
                return resultado
        df[col_fecha] = fechas
        hoy = datetime.now().date()
        df["_clave_trib"] = df[col_trib].apply(detectar_tribunal)
        tribunales_invalidos = int(df["_clave_trib"].isna().sum())
        if tribunales_invalidos:
            resultado["errores"].append(
                f"TRIBUNAL no reconocido en {tribunales_invalidos} fila(s); "
                "se agruparon por el valor original de la celda"
            )
        df["_clave_trib"] = df.apply(
            lambda r: r["_clave_trib"] or normalizar(r[col_trib]), axis=1)
        df["_clave_prog"] = df[col_prog].apply(self._clave_programa)
        df["_dias"] = df[col_fecha].apply(
            lambda f: (f.date() - hoy).days if pd.notna(f) else None)
        df["_tipo"] = df["_dias"].apply(
            lambda d: "VENCIDO" if d is not None and d < 0
                      else ("POR_VENCER" if d is not None and 0 <= d <= DIAS_POR_VENCER
                            else "FUTURO"))

        vencidos   = df[df["_tipo"] == "VENCIDO"]
        por_vencer = df[df["_tipo"] == "POR_VENCER"]

        for _, g in vencidos.groupby(["_clave_prog", "_clave_trib"], dropna=False):
            prog = g.iloc[0][col_prog]
            trib_raw = g.iloc[0][col_trib]
            self._prog_borrador(g, prog, trib_raw, col_rit, col_trib, col_rut,
                                col_nombre, col_fecha, "vencidos", resultado)


        for _, g in por_vencer.groupby(["_clave_prog", "_clave_trib"], dropna=False):
            prog = g.iloc[0][col_prog]
            trib_raw = g.iloc[0][col_trib]
            self._prog_borrador(g, prog, trib_raw, col_rit, col_trib, col_rut,
                                col_nombre, col_fecha, "por_vencer", resultado)

        return resultado

    # ── Lista de espera ───────────────────────────────────────────────────────

    def procesar_espera(self, df: pd.DataFrame) -> dict:
        """
        Excel ESPERA, procesado o de origen.
        Genera borrador tipo D: 1 por PROGRAMA+TRIBUNAL donde OBSERVACION
        contiene el patrón de lista de espera. Si no existe OBSERVACION, aplica
        directamente el mismo umbral de ESPERA (T ESPERA >= 30).
        Cols: RIT, TRIBUNAL, RUT, NOMBRE, DERIVACIÓN, T ESPERA, OBSERVACION.
        """
        resultado = self._vacio()
        df = df.copy()
        df.columns = [str(c).strip() for c in df.columns]

        col_rit    = self._col(df, ALIAS_RIT)
        col_trib   = self._col(df, ALIAS_TRIBUNAL)
        col_rut    = self._col(df, ALIAS_RUT)
        col_nombre = self._col(df, ALIAS_NOMBRE)
        col_prog   = self._col(df, ALIAS_PROGRAMA)
        col_espera = self._col(df, ALIAS_ESPERA)
        col_obs    = self._col(df, ["OBSERVACION", "OBSERVACIÓN"])

        faltantes = [n for n, c in [("RIT", col_rit), ("TRIBUNAL", col_trib),
                     ("NOMBRE", col_nombre), ("PROGRAMA", col_prog),
                     ("T ESPERA", col_espera)] if not c]
        if faltantes:
            resultado["errores"].append(f"Columnas no encontradas: {faltantes}")
            return resultado

        # Degradación suave (D9): filas defectuosas fuera con aviso, el resto
        # del lote sigue. Nunca abortar todo por una fila.
        df = self._excluir_filas_incompletas(df, {
            "RIT": col_rit,
            "TRIBUNAL": col_trib,
            "NOMBRE": col_nombre,
            "PROGRAMA": col_prog,
            "T ESPERA": col_espera,
        }, resultado)
        if df.empty:
            return resultado

        tribunales_invalidos = int(df[col_trib].apply(detectar_tribunal).isna().sum())
        if tribunales_invalidos:
            resultado["errores"].append(
                f"TRIBUNAL no reconocido en {tribunales_invalidos} fila(s); "
                "se agruparon por el valor original de la celda"
            )

        esperas_invalidas = df[col_espera].apply(
            lambda valor: self._entero_no_negativo(valor) is None
        )
        if esperas_invalidas.any():
            resultado["errores"].append(
                f"T ESPERA inválido en {int(esperas_invalidas.sum())} fila(s); "
                "esas filas se omitieron"
            )
            df = df[~esperas_invalidas].copy()
            if df.empty:
                return resultado

        # Un archivo de origen también es válido: con T ESPERA suficiente el
        # motor habría producido E-05, que es precisamente el criterio para
        # este correo. Una OBSERVACION existente conserva prioridad para no
        # incluir filas que el motor ya descartó por alguna regla especial.
        if col_obs:
            df_filtrado = df[df[col_obs].apply(
                lambda obs: PATRON_ESPERA in normalizar(obs)
            )].copy()
        else:
            df_filtrado = df[df[col_espera].apply(
                lambda valor: self._entero_no_negativo(valor) >= 30
            )].copy()
            # El motor omite E-05 cuando no puede identificar el tribunal.
            df_filtrado = df_filtrado[df_filtrado[col_trib].apply(
                lambda valor: detectar_tribunal(valor) is not None
            )].copy()

        if df_filtrado.empty:
            resultado["errores"].append(
                "No hay filas elegibles para correo de lista de espera."
            )
            return resultado

        df_filtrado["_clave_prog"] = df_filtrado[col_prog].apply(self._clave_programa)
        df_filtrado["_clave_trib"] = df_filtrado[col_trib].apply(
            lambda v: detectar_tribunal(v) or normalizar(v))
        for _, grupo in df_filtrado.groupby(["_clave_prog", "_clave_trib"], dropna=False):
            prog = grupo.iloc[0][col_prog]
            trib_raw = grupo.iloc[0][col_trib]
            contacto = self.catastro.resolver(str(prog), permitir_fuzzy=True)
            if not contacto or not self._mail_valido(contacto.get("mail", "")):
                msg = f"Sin contacto: '{prog}'"
                if msg not in resultado["grupos_sin_contacto"]:
                    resultado["grupos_sin_contacto"].append(msg)
                continue

            nombre_prog_fmt = titulo_programa(contacto["nombre"])
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
                TRIBUNAL=TRIBUNAL_DISPLAY.get(detectar_tribunal(trib_raw), str(trib_raw).strip()),
                FILAS=filas_html,
            )
            asunto = f"Lista de espera — {nombre_prog_fmt} / {TRIBUNAL_DISPLAY.get(detectar_tribunal(trib_raw), str(trib_raw).strip())}"

            try:
                self._despachador({"para":[contacto["mail"]], "cc":[self.cc_fijo], "asunto":asunto, "cuerpo_html":cuerpo, "n_registros":len(grupo)})
                resultado["borradores_creados"] += 1
                resultado["detalle"].append(
                    f"✓ [ESPERA] {nombre_prog_fmt} / {str(trib_raw).strip()} "
                    f"→ {contacto['mail']} ({len(grupo)} registros)"
                )
            except Exception as e:
                resultado["errores"].append(str(e))

        return resultado

    # ── Borradores internos ───────────────────────────────────────────────────

    def _prog_borrador(self, grupo, prog, trib_raw, col_rit, col_trib, col_rut,
                        col_nombre, col_fecha, tipo, resultado):
        contacto = self.catastro.resolver(str(prog), permitir_fuzzy=True)
        if not contacto or not self._mail_valido(contacto.get("mail", "")):
            msg = f"Sin contacto: '{prog}'"
            if msg not in resultado["grupos_sin_contacto"]:
                resultado["grupos_sin_contacto"].append(msg)
            return

        nombre_prog_fmt = titulo_programa(contacto["nombre"])
        filas_html = ""
        for _, row in grupo.iterrows():
            rut = str(row[col_rut]) if col_rut and col_rut in row.index else "---"
            filas_html += _fila_fecha(str(row[col_rit]), str(row[col_trib]),
                                       rut, str(row[col_nombre]), row[col_fecha])

        plantilla = ("correo_programa_vencidos.html" if tipo == "vencidos"
                     else "correo_programa_por_vencer.html")
        asunto_txt = "vencidos" if tipo == "vencidos" else "por vencer"
        clave = detectar_tribunal(trib_raw)
        trib_fmt = TRIBUNAL_DISPLAY.get(clave, str(trib_raw).strip())
        if clave is None:
            resultado["errores"].append(f"Tribunal no reconocido: {trib_raw}")
        # Mejoras §1.3: el criterio DCE se aplica sobre la DERIVACIÓN de RUS
        # (mismo dato que usa el motor), no sobre el nombre del catastro.
        etiqueta = "informes diagnósticos" if es_dce(prog) else "informes de avance"
        asunto = f"{etiqueta.capitalize()} {asunto_txt} — {nombre_prog_fmt} / {trib_fmt}"

        cuerpo = _cargar_plantilla(plantilla, NOMBRE_PROGRAMA=nombre_prog_fmt,
                                    TRIBUNAL=trib_fmt, FILAS=filas_html, ETIQUETA_INFORMES=etiqueta)
        try:
            self._despachador({"para":[contacto["mail"]], "cc":[self.cc_fijo], "asunto":asunto, "cuerpo_html":cuerpo, "n_registros":len(grupo)})
            resultado["borradores_creados"] += 1
            resultado["detalle"].append(
                f"✓ [{tipo.upper()}] {nombre_prog_fmt} / {str(trib_raw).strip()} "
                f"→ {contacto['mail']} ({len(grupo)} registros)"
            )
        except Exception as e:
            resultado["errores"].append(str(e))


    @staticmethod
    def _crear_borrador(borrador: dict) -> bool:
        return _crear_borrador_outlook(
            para=borrador["para"], cc=borrador["cc"],
            asunto=borrador["asunto"], cuerpo_html=borrador["cuerpo_html"]
        )

    # ── Helpers ───────────────────────────────────────────────────────────────

    @staticmethod
    def _vacio() -> dict:
        return {"borradores_creados": 0, "grupos_sin_contacto": [],
                "errores": [], "detalle": []}

    @staticmethod
    def _col(df: pd.DataFrame, aliases: list[str]) -> str | None:
        for c in df.columns:
            for a in aliases:
                if normalizar(c) == normalizar(a):
                    return c
        return None

    @staticmethod
    def _mail_valido(mail: str) -> bool:
        texto = str(mail).strip()
        if re.search(r"[\s;,\r\n]", texto):
            return False
        return bool(re.fullmatch(r"[^@]+@[^@]+\.[^@]+", texto))

    @staticmethod
    def _excluir_filas_incompletas(df: pd.DataFrame, campos: dict[str, str],
                                   resultado: dict) -> pd.DataFrame:
        """Excluye (con aviso) las filas con campos requeridos vacíos.

        Retorna el DataFrame filtrado — el lote restante sigue procesándose.
        """
        mascara_total = pd.Series(False, index=df.index)
        for etiqueta, columna in campos.items():
            serie = df[columna]
            vacios = serie.isna() | serie.astype(str).str.strip().str.lower().isin(
                ("", "nan", "none", "nat", "<na>")
            )
            cantidad = int(vacios.sum())
            if cantidad:
                resultado["errores"].append(
                    f"{etiqueta} vacío en {cantidad} fila(s); esas filas se omitieron"
                )
            mascara_total |= vacios
        return df[~mascara_total].copy()

    @staticmethod
    def _entero_no_negativo(valor):
        try:
            numero = float(str(valor).strip())
            return int(numero) if numero >= 0 and numero.is_integer() else None
        except (OverflowError, TypeError, ValueError):
            return None

    def _clave_programa(self, valor) -> str:
        contacto = self.catastro.resolver(str(valor), permitir_fuzzy=True)
        return normalizar(contacto["nombre"] if contacto else valor)
