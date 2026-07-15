#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
motor/textos.py — v8.13
Fuente unica de verdad para textos de observacion.
Elimina el bug recurrente: strings hardcodeados en reglas_*.py que se
re-tipean cada sesion y derivan del texto acordado con el usuario.

Editar SOLO motor/textos_observaciones.json. Nunca escribir texto de
observacion directamente en un archivo .py.
"""

import json
from pathlib import Path

_RUTA = Path(__file__).parent / "textos_observaciones.json"
_CACHE = None


class _SafeDict(dict):
    """Placeholder sin dato -> queda visible como {CAMPO}, nunca crash ni vacio."""
    def __missing__(self, key):
        return "{" + key + "}"


def _cargar():
    global _CACHE
    if _CACHE is None:
        with open(_RUTA, encoding="utf-8") as f:
            _CACHE = json.load(f)
    return _CACHE


def render(modo: str, id_regla: str, **datos) -> str:
    """
    modo: 'ESPERA' | 'CUMPLIMIENTO' | 'INFORMES'
    id_regla: clave dentro del modo (ej. 'R7_CURADOR')
    datos: placeholders nombrados (PNOMBRE, PROGRAMA, FECHA_RESOLUCION, etc.)
    """
    try:
        entry = _cargar()[modo][id_regla]
    except KeyError:
        return f"[TEXTO NO DEFINIDO: {modo}.{id_regla}]"
    return entry["texto"].format_map(_SafeDict(datos))


def confirmado(modo: str, id_regla: str) -> bool:
    try:
        return bool(_cargar()[modo][id_regla].get("confirmado", False))
    except KeyError:
        return False


def listar_no_confirmados():
    """Utilidad de auditoria: retorna [(modo, id_regla), ...] con confirmado=False."""
    data = _cargar()
    out = []
    for modo, reglas in data.items():
        if modo == "_meta":
            continue
        for id_regla, entry in reglas.items():
            if not entry.get("confirmado", False):
                out.append((modo, id_regla))
    return out
