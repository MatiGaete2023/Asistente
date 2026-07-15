#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Logs — Manager de bitácora diaria con rotación automática."""
import hashlib
import logging
from datetime import datetime, timedelta
from pathlib import Path

def get_logger(ruta_logs: str, nombre: str = "csmp",
               dias_retencion: int = 90) -> logging.Logger:
    """Retorna logger configurado con archivo diario."""
    dias_retencion = int(dias_retencion)
    if not 1 <= dias_retencion <= 3650:
        raise ValueError("dias_retencion debe estar entre 1 y 3650")

    directorio = Path(ruta_logs).expanduser().resolve()
    directorio.mkdir(parents=True, exist_ok=True)
    nombre_archivo = f"csmp_{datetime.now():%Y%m%d}.log"
    ruta = directorio / nombre_archivo

    identificador = hashlib.sha256(str(directorio).encode("utf-8")).hexdigest()[:16]
    logger = logging.getLogger(f"{nombre}.{identificador}")
    logger.propagate = False
    if not any(
        isinstance(handler, logging.FileHandler)
        and Path(handler.baseFilename) == ruta
        for handler in logger.handlers
    ):
        handler = logging.FileHandler(str(ruta), encoding="utf-8")
        handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
        logger.addHandler(handler)
    logger.setLevel(logging.INFO)

    _rotar(directorio, dias=dias_retencion)
    return logger


def _rotar(ruta_logs: Path, dias: int):
    """Elimina logs más antiguos que `dias` días."""
    umbral = datetime.now() - timedelta(days=dias)
    for archivo in ruta_logs.glob("csmp_*.log"):
        try:
            fecha_str = archivo.stem.replace("csmp_", "")
            fecha = datetime.strptime(fecha_str, "%Y%m%d")
            if fecha < umbral:
                archivo.unlink()
        except ValueError:
            # Archivos con otro patrón no pertenecen a la rotación diaria.
            continue
