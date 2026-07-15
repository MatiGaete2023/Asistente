# -*- coding: utf-8 -*-
"""
contactos_programas.py — Resolución programa → {mail, director}

Carga catastro_programas.json y resuelve contacto por matching fuzzy
normalizado (sin tildes, sin mayúsculas, sin espacios dobles, sin guiones).

API:
    from comunicaciones.contactos_programas import CatastroContactos
    cat = CatastroContactos("ruta/catastro_programas.json")
    contacto = cat.resolver("AFT  MULCHEN")
    # → {"nombre": "AFT - MULCHEN", "mail": "...", "director": "..."}
    # → None si no hay match
"""

import json
import re
import unicodedata
from pathlib import Path

_PAREN_RE        = re.compile(r"\(.*?\)")
_PAREN_SUELTO_RE = re.compile(r"[()]")
_SEPARADOR_RE    = re.compile(r"[-–—_/\\.]")
_ESPACIOS_RE     = re.compile(r"\s+")


def _normalizar(txt: str) -> str:
    """Normalización agresiva para matching fuzzy."""
    if not txt:
        return ""
    txt = unicodedata.normalize("NFKD", str(txt)).encode("ASCII", "ignore").decode()
    txt = txt.lower()
    txt = _SEPARADOR_RE.sub(" ", txt)   # guiones y separadores → espacio
    txt = _ESPACIOS_RE.sub(" ", txt)    # espacios múltiples → uno
    return txt.strip()


def _tokens_significativos(txt: str) -> set:
    """Tokens >3 chars tras limpiar paréntesis. Para matching token-set (nivel 3)."""
    s = _PAREN_RE.sub("", _normalizar(txt))
    s = _PAREN_SUELTO_RE.sub("", s)
    return {t for t in s.split() if len(t) > 3}


def _primer_nombre_dos_apellidos(nombre_completo: str) -> str:
    """
    'Gissela Loreto Montoya Salvo' → 'Gissela Montoya Salvo'
    'Nancy Carola Oliva Peña'      → 'Nancy Oliva Peña'
    'Oscar Vasquez Vergara'        → 'Oscar Vasquez Vergara'
    'Luis Alveal Riquelme'         → 'Luis Alveal Riquelme'

    Regla: tomar primer token como nombre, saltar tokens intermedios que
    parezcan segundo nombre (si hay ≥4 tokens), tomar los 2 últimos como
    apellidos. Si hay exactamente 3 tokens → nombre + 2 apellidos directos.
    Si hay ≤2 tokens → retornar tal cual.
    """
    partes = [p for p in nombre_completo.strip().split() if p]
    if len(partes) == 0:
        return ""
    if len(partes) <= 2:
        return " ".join(p.capitalize() for p in partes)
    if len(partes) == 3:
        # nombre apellido1 apellido2 → directo
        return " ".join(p.capitalize() for p in [partes[0], partes[1], partes[2]])
    # ≥4 tokens: nombre [segundonombre...] apellido1 apellido2
    nombre    = partes[0].capitalize()
    apellido1 = partes[-2].capitalize()
    apellido2 = partes[-1].capitalize()
    return f"{nombre} {apellido1} {apellido2}"


class CatastroContactos:
    def __init__(self, ruta_json: str):
        self._ruta = Path(ruta_json)
        self._registros = []        # lista de dicts originales
        self._indice = {}           # norm_nombre → registro

        self._cargar()

    def _cargar(self):
        if not self._ruta.exists():
            raise FileNotFoundError(f"Catastro no encontrado: {self._ruta}")

        with open(self._ruta, encoding="utf-8") as f:
            raw = f.read()

        # Formato: objetos JSON separados sin array wrapper
        raw = raw.strip()
        if not raw.startswith("["):
            raw = "[" + raw.rstrip(",") + "]"
        raw = re.sub(r",\s*\]", "]", raw)

        data = json.loads(raw)

        for r in data:
            nombre   = r.get("Nombre", "").strip()
            mail     = r.get("Mail", "").strip()
            director = r.get("Director", "").strip()
            if not nombre:
                continue
            registro = {
                "nombre":   nombre,
                "mail":     mail.lower(),
                "director": director,
                "saludo":   _primer_nombre_dos_apellidos(director) if director else "",
                "_tokens":  _tokens_significativos(nombre),   # precomputado para nivel 3
            }
            self._registros.append(registro)
            self._indice[_normalizar(nombre)] = registro

        # Cargar aliases (nombre RUS → nombre catastro)
        # Permite resolver programas cuyo nombre en RUS difiere del catastro
        # sin modificar el catastro principal.
        ruta_aliases = self._ruta.parent / "aliases_programas.json"
        if ruta_aliases.exists():
            try:
                with open(ruta_aliases, encoding="utf-8") as f:
                    aliases = json.load(f)
                for a in aliases:
                    alias_norm    = _normalizar(a.get("alias", ""))
                    catastro_norm = _normalizar(a.get("nombre_catastro", ""))
                    if alias_norm and catastro_norm and catastro_norm in self._indice:
                        # Registrar el alias apuntando al mismo registro del catastro
                        self._indice[alias_norm] = self._indice[catastro_norm]
            except Exception:
                pass  # aliases opcionales — nunca bloquear la carga principal

    def resolver(self, nombre_programa: str) -> dict | None:
        """
        Busca contacto para nombre_programa.
        1. Exact match normalizado.
        2. Substring bidireccional.
        3. Token-set: comparte ≥2 tokens significativos (>3 chars) y ratio ≥45%.
           Resuelve variantes como 'AFT EL CONQUISTADOR YUMBEL' vs
           'AFT - EL CONQUISTADOR DE YUMBEL'.
        Retorna dict o None.
        """
        clave = _normalizar(nombre_programa)

        # 1. Exact
        if clave in self._indice:
            return self._indice[clave]

        # 2. Substring bidireccional
        mejor = None
        mejor_len = 0
        for k, reg in self._indice.items():
            if k in clave or clave in k:
                if len(k) > mejor_len:
                    mejor = reg
                    mejor_len = len(k)
        if mejor:
            return mejor

        # 3. Token-set — usa tokens precomputados del catastro (no recalcula)
        tokens_q = _tokens_significativos(nombre_programa)
        if not tokens_q:
            return None

        mejor_score  = (0, 0.0)
        mejor_nivel3 = None
        for reg in self._registros:
            tokens_k = reg["_tokens"]
            inter    = tokens_q & tokens_k
            if len(inter) < 2:
                continue
            ratio = len(inter) / max(len(tokens_q), len(tokens_k), 1)
            score = (len(inter), ratio)
            if ratio >= 0.45 and score > mejor_score:
                mejor_score  = score
                mejor_nivel3 = reg

        return mejor_nivel3

    def todos(self) -> list:
        return list(self._registros)
