#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
motor/textos_confirmacion.py — v8.14 (S4)

Flujo de confirmación de textos que NO depende de sesiones de chat con
una IA. Resuelve F-7 estructuralmente:

  1. exportar_pendientes(ruta_xlsx) -> escribe un .xlsx con los textos
     que tienen confirmado=False actualmente en
     motor/textos_observaciones.json.

  2. Matías abre ese Excel y, por fila:
       - Si el texto está bien tal cual: deja TEXTO_CORRECTO vacío y
         escribe SI en APROBADO.
       - Si el texto necesita cambios: escribe la redacción exacta en
         TEXTO_CORRECTO (y opcionalmente SI en APROBADO si ya la da
         por buena así).

  3. importar_confirmaciones(ruta_xlsx) -> aplica esos cambios a
     motor/textos_observaciones.json:
       - TEXTO_CORRECTO no vacío -> reemplaza "texto" (COPY-PASTE
         literal de lo que escribió Matías, NUNCA parafraseado).
       - APROBADO == "SI" -> marca "confirmado": true.
     Hace backup con timestamp del JSON antes de sobrescribirlo.

Restricción inviolable del proyecto (ver memoria / PLAN v8.14 §2.1):
el texto que termina en el JSON es SIEMPRE lo que el usuario escribió,
verbatim. Este módulo no parafrasea, no correige ortografía, no agrega
ni quita puntuación.
"""

import json
import shutil
from datetime import datetime
from pathlib import Path

import pandas as pd

from .textos import _RUTA

COLUMNAS = ["MODO", "ID", "TEXTO_ACTUAL", "TEXTO_CORRECTO", "APROBADO"]


def _cargar_json(ruta_json: Path) -> dict:
    with open(ruta_json, encoding="utf-8") as f:
        return json.load(f)


def _listar_no_confirmados_de(data: dict):
    out = []
    for modo, reglas in data.items():
        if modo == "_meta":
            continue
        for id_regla, entry in reglas.items():
            if not entry.get("confirmado", False):
                out.append((modo, id_regla))
    return out


def contar_pendientes(ruta_json=None) -> int:
    """Cantidad de textos con confirmado=False. Usado para el badge del
    botón en la GUI: '📋 Textos pendientes (N)'."""
    ruta_json = Path(ruta_json) if ruta_json else _RUTA
    data = _cargar_json(ruta_json)
    return len(_listar_no_confirmados_de(data))


def exportar_pendientes(ruta_xlsx, ruta_json=None) -> int:
    """Escribe TEXTOS_PENDIENTES.xlsx. Retorna cantidad de filas."""
    ruta_json = Path(ruta_json) if ruta_json else _RUTA
    data = _cargar_json(ruta_json)
    pendientes = _listar_no_confirmados_de(data)

    filas = []
    for modo, id_regla in pendientes:
        entry = data[modo][id_regla]
        filas.append({
            "MODO": modo,
            "ID": id_regla,
            "TEXTO_ACTUAL": entry["texto"],
            "TEXTO_CORRECTO": "",
            "APROBADO": "",
        })

    df = pd.DataFrame(filas, columns=COLUMNAS)
    df.to_excel(ruta_xlsx, index=False)
    return len(filas)


def _backup_json(ruta_json: Path) -> Path:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    destino = ruta_json.parent / f"{ruta_json.stem}.backup_{ts}{ruta_json.suffix}"
    shutil.copy2(ruta_json, destino)
    return destino


def importar_confirmaciones(ruta_xlsx, ruta_json=None) -> dict:
    """
    Aplica TEXTOS_PENDIENTES.xlsx (ya editado) a
    motor/textos_observaciones.json. Retorna:
        {
          "actualizados": [ "MODO.ID", ... ],   # texto reemplazado
          "confirmados":  [ "MODO.ID", ... ],   # confirmado=true
          "sin_cambios":  [ "MODO.ID", ... ],
          "errores":      [ "mensaje", ... ],
          "backup":       "ruta al JSON previo respaldado",
        }
    Lanza ValueError si al Excel le faltan columnas requeridas.
    """
    ruta_json = Path(ruta_json) if ruta_json else _RUTA
    df = pd.read_excel(ruta_xlsx)
    faltan = [c for c in COLUMNAS if c not in df.columns]
    if faltan:
        raise ValueError(f"Faltan columnas en el Excel: {faltan}")

    data = _cargar_json(ruta_json)
    resumen = {"actualizados": [], "confirmados": [], "sin_cambios": [], "errores": []}

    for _, row in df.iterrows():
        modo = str(row.get("MODO", "") or "").strip()
        id_regla = str(row.get("ID", "") or "").strip()
        texto_correcto = str(row.get("TEXTO_CORRECTO", "") or "").strip()
        aprobado = str(row.get("APROBADO", "") or "").strip().upper()

        if not modo or not id_regla:
            continue
        if modo not in data or id_regla not in data.get(modo, {}):
            resumen["errores"].append(f"{modo}.{id_regla}: no existe en el JSON actual")
            continue

        entry = data[modo][id_regla]
        cambiado = False

        if texto_correcto:
            entry["texto"] = texto_correcto  # copy-paste literal, nunca parafraseado
            resumen["actualizados"].append(f"{modo}.{id_regla}")
            cambiado = True

        if aprobado == "SI":
            entry["confirmado"] = True
            resumen["confirmados"].append(f"{modo}.{id_regla}")
            cambiado = True

        if not cambiado:
            resumen["sin_cambios"].append(f"{modo}.{id_regla}")

    backup = _backup_json(ruta_json)
    with open(ruta_json, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    if ruta_json == _RUTA:
        # invalidar cache de motor/textos.py para que render() refleje
        # el cambio en la MISMA sesión, sin reiniciar el programa.
        import motor.textos as _textos_mod
        _textos_mod._CACHE = None

    resumen["backup"] = str(backup)
    return resumen
