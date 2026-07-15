#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
motor/textos.py — v9.0.1
Fuente unica de verdad para textos de observacion.
Elimina el bug recurrente: strings hardcodeados en reglas_*.py que se
re-tipean cada sesion y derivan del texto acordado con el usuario.

Editar SOLO motor/textos_observaciones.json. Nunca escribir texto de
observacion directamente en un archivo .py.
"""

import json
import os
import shutil
import sys
import tempfile
from pathlib import Path
from threading import RLock

_RUTA_EMPAQUETADA = Path(__file__).parent / "textos_observaciones.json"


def _ruta_catalogo() -> Path:
    """Usa copia editable para ejecutables o instalaciones no editables."""
    instalado = any(
        parte.lower() in {"site-packages", "dist-packages"}
        for parte in _RUTA_EMPAQUETADA.parts
    )
    if not getattr(sys, "frozen", False) and not instalado:
        return _RUTA_EMPAQUETADA

    base = Path(os.environ.get("LOCALAPPDATA") or Path.home()) / "CSMP_RUS"
    destino = base / "textos_observaciones.json"
    if not destino.exists():
        base.mkdir(parents=True, exist_ok=True)
        fd, temporal_nombre = tempfile.mkstemp(
            prefix=f".{destino.stem}.", suffix=".tmp", dir=base
        )
        os.close(fd)
        temporal = Path(temporal_nombre)
        try:
            shutil.copy2(_RUTA_EMPAQUETADA, temporal)
            os.replace(temporal, destino)
        except Exception:
            temporal.unlink(missing_ok=True)
            raise
    return destino


_RUTA = _ruta_catalogo()
_CACHE = None
_CACHE_LOCK = RLock()


def ruta_catalogo_textos() -> Path:
    return _RUTA


class ErrorCatalogoTextos(RuntimeError):
    """El catálogo no permite producir una observación completa y segura."""


def _cargar():
    global _CACHE
    with _CACHE_LOCK:
        if _CACHE is None:
            with open(_RUTA, encoding="utf-8") as f:
                _CACHE = json.load(f)
        return _CACHE


def invalidar_cache():
    """Invalida el catálogo de forma segura entre hilos."""
    global _CACHE
    with _CACHE_LOCK:
        _CACHE = None


def render(modo: str, id_regla: str, **datos) -> str:
    """
    modo: 'ESPERA' | 'CUMPLIMIENTO' | 'INFORMES'
    id_regla: clave dentro del modo (ej. 'R7_CURADOR')
    datos: placeholders nombrados (PNOMBRE, PROGRAMA, FECHA_RESOLUCION, etc.)
    """
    try:
        entry = _cargar()[modo][id_regla]
    except KeyError as exc:
        raise ErrorCatalogoTextos(
            f"Texto no definido: {modo}.{id_regla}"
        ) from exc
    try:
        return entry["texto"].format_map(datos)
    except KeyError as exc:
        raise ErrorCatalogoTextos(
            f"Falta el placeholder {exc.args[0]} para {modo}.{id_regla}"
        ) from exc
    except (TypeError, ValueError) as exc:
        raise ErrorCatalogoTextos(
            f"Texto inválido en {modo}.{id_regla}: {exc}"
        ) from exc


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
