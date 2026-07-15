# -*- coding: utf-8 -*-
"""Composición segura de observaciones e incidencias de validación."""
from datetime import datetime
import re
import pandas as pd
from .utilidades import prefijo_observacion, get_date

def prefijo(nombre, programa) -> str:
    return prefijo_observacion(nombre, programa)

def _normalizar_fragmento(fragmento):
    if fragmento is None:
        return ""
    texto = str(fragmento).strip()
    if not texto:
        return ""
    texto = re.sub(r"\s+", " ", texto)
    texto = texto.rstrip(". ") + "."
    texto = re.sub(r"\.{2,}", ".", texto)
    return texto

def componer(pfx, fragmentos) -> str:
    partes = [_normalizar_fragmento(f) for f in fragmentos]
    partes = [p for p in partes if p]
    if not partes:
        return ""
    texto = " ".join(partes)
    texto = re.sub(r"\s+", " ", texto).replace("..", ".")
    return f"{pfx or ''}{texto}"

class Incidencias:
    def __init__(self):
        self._items = []
    def agregar(self, fila_excel, rit, regla, motivo):
        self._items.append({"FILA_EXCEL": fila_excel, "RIT": rit or "", "REGLA": regla, "MOTIVO": motivo})
    def __len__(self):
        return len(self._items)
    def extender(self, otra):
        if otra:
            self._items.extend(otra._items)
    def como_dataframe(self):
        return pd.DataFrame(self._items, columns=["FILA_EXCEL", "RIT", "REGLA", "MOTIVO"])

def fecha_valida(valor):
    fec = get_date(valor)
    if fec is None:
        return None
    return fec if isinstance(fec, datetime) else pd.to_datetime(fec).to_pydatetime()
