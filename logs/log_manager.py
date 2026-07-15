#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Logs — Manager de bitácora diaria con rotación automática."""
import logging
from datetime import datetime, timedelta
from pathlib import Path

def get_logger(ruta_logs: str, nombre: str = "csmp") -> logging.Logger:
    """Retorna logger configurado con archivo diario."""
    Path(ruta_logs).mkdir(parents=True, exist_ok=True)
    nombre_archivo = f"csmp_{datetime.now():%Y%m%d}.log"
    ruta = Path(ruta_logs) / nombre_archivo

    logger = logging.getLogger(nombre)
    if not logger.handlers:
        handler = logging.FileHandler(str(ruta), encoding="utf-8")
        handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
        logger.addHandler(handler)
        logger.setLevel(logging.INFO)

    _rotar(Path(ruta_logs), dias=90)
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
        except Exception:
            pass
