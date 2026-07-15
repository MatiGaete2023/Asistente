# -*- coding: utf-8 -*-
"""
contactos_programas.py — Resolución programa → {mail, director}

Carga catastro_programas.json y resuelve automáticamente solo nombres o
aliases exactos normalizados. El matching fuzzy requiere opt-in explícito.

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
    'Ana María Pérez Soto' → 'Ana Pérez Soto'
    'Tomás Díaz Rojas'     → 'Tomás Díaz Rojas'

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
        self._claves_ambiguas = set()

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
        if not isinstance(data, list) or not data:
            raise ValueError(f"Catastro vacío o inválido: {self._ruta}")

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
            clave = _normalizar(nombre)
            existente = self._indice.get(clave)
            if existente is not None and (
                    existente["nombre"] != registro["nombre"] or
                    existente["mail"] != registro["mail"]):
                # Nunca elegir silenciosamente el último registro cuando dos
                # nombres diferentes colapsan a la misma clave normalizada.
                self._claves_ambiguas.add(clave)
            else:
                self._indice[clave] = registro

        # Cargar aliases (nombre RUS → nombre catastro)
        # Permite resolver programas cuyo nombre en RUS difiere del catastro
        # sin modificar el catastro principal.
        ruta_aliases = self._ruta.parent / "aliases_programas.json"
        # Los aliases son opcionales: una entrada malformada se OMITE con
        # aviso, nunca bloquea la carga principal (diseño original v8.x).
        if ruta_aliases.exists():
            try:
                with open(ruta_aliases, encoding="utf-8") as f:
                    aliases = json.load(f)
            except (OSError, json.JSONDecodeError):
                aliases = []
            if not isinstance(aliases, list):
                aliases = []
            for a in aliases:
                if not isinstance(a, dict):
                    continue
                alias_norm = _normalizar(a.get("alias", ""))
                catastro_norm = _normalizar(a.get("nombre_catastro", ""))
                if not alias_norm or not catastro_norm:
                    continue
                if (catastro_norm in self._claves_ambiguas or
                        catastro_norm not in self._indice):
                    continue

                destino = self._indice[catastro_norm]
                existente = self._indice.get(alias_norm)
                if existente is not None and existente is not destino:
                    self._claves_ambiguas.add(alias_norm)
                    self._indice.pop(alias_norm, None)
                    continue
                if alias_norm not in self._claves_ambiguas:
                    self._indice[alias_norm] = destino

    def resolver(self, nombre_programa: str, *, permitir_fuzzy: bool = False) -> dict | None:
        """
        Busca contacto para nombre_programa.
        1. Exact match normalizado (incluye aliases aprobados).
        2. Con ``permitir_fuzzy=True``, substring bidireccional único.
        3. Con opt-in, token-set: comparte ≥2 tokens y ratio ≥45%.
           Resuelve variantes como 'AFT EL CONQUISTADOR YUMBEL' vs
           'AFT - EL CONQUISTADOR DE YUMBEL'.
        Retorna dict o None.
        """
        clave = _normalizar(nombre_programa)
        if not clave or clave in self._claves_ambiguas:
            return None

        # 1. Exact
        if clave in self._indice:
            return self._indice[clave]

        # La generación automática solo admite coincidencias exactas o aliases
        # aprobados. El fuzzy matching queda disponible únicamente para una
        # futura pantalla de selección humana.
        if not permitir_fuzzy or len(clave) < 5:
            return None

        # 2. Substring bidireccional. Solo resolver si hay exactamente un
        # contacto candidato; ante ambigüedad se exige intervención humana.
        candidatos = {}
        for k, reg in self._indice.items():
            if k in self._claves_ambiguas:
                continue
            if k in clave or clave in k:
                candidatos[(reg["nombre"], reg["mail"])] = reg
        if len(candidatos) == 1:
            return next(iter(candidatos.values()))
        if len(candidatos) > 1:
            return None

        # 3. Token-set — usa tokens precomputados del catastro (no recalcula)
        tokens_q = _tokens_significativos(nombre_programa)
        if not tokens_q:
            return None

        mejor_score = (0, 0.0)
        mejores = {}
        for reg in self._registros:
            if _normalizar(reg["nombre"]) in self._claves_ambiguas:
                continue
            tokens_k = reg["_tokens"]
            inter    = tokens_q & tokens_k
            if len(inter) < 2:
                continue
            ratio = len(inter) / max(len(tokens_q), len(tokens_k), 1)
            score = (len(inter), ratio)
            if ratio < 0.45:
                continue
            if score > mejor_score:
                mejor_score = score
                mejores = {(reg["nombre"], reg["mail"]): reg}
            elif score == mejor_score:
                mejores[(reg["nombre"], reg["mail"])] = reg

        return next(iter(mejores.values())) if len(mejores) == 1 else None

    def todos(self) -> list:
        return list(self._registros)
