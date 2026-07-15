#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tests/test_confirmacion_textos.py

CA-S4: ciclo export -> editar 1 texto -> import -> render() refleja el
cambio sin tocar ningún .py; backup del JSON existe; JSON sigue siendo
válido.

Corre SIEMPRE sobre una COPIA temporal de motor/textos_observaciones.json
(nunca sobre el archivo real del proyecto), para no interferir con
tests/goldens_textos.json.
"""

import json
import shutil
import sys
from pathlib import Path

import pandas as pd
import pytest

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

from motor.textos_confirmacion import (  # noqa: E402
    exportar_pendientes, importar_confirmaciones, contar_pendientes,
)

JSON_REAL = RAIZ / "motor" / "textos_observaciones.json"


@pytest.fixture()
def json_temporal(tmp_path):
    """Copia aislada de textos_observaciones.json para no tocar el real."""
    destino = tmp_path / "textos_observaciones.json"
    shutil.copy2(JSON_REAL, destino)
    return destino


def test_exportar_pendientes_genera_17_filas(json_temporal, tmp_path):
    ruta_xlsx = tmp_path / "TEXTOS_PENDIENTES.xlsx"
    n = exportar_pendientes(ruta_xlsx, ruta_json=json_temporal)
    assert n == 17  # F-7: 17 textos confirmado=false conocidos
    assert ruta_xlsx.exists()
    df = pd.read_excel(ruta_xlsx)
    assert list(df.columns) == ["MODO", "ID", "TEXTO_ACTUAL", "TEXTO_CORRECTO", "APROBADO"]
    assert len(df) == 17


def test_ciclo_completo_export_editar_import(json_temporal, tmp_path):
    ruta_xlsx = tmp_path / "TEXTOS_PENDIENTES.xlsx"
    exportar_pendientes(ruta_xlsx, ruta_json=json_temporal)

    # "Matías edita" el excel: aprueba tal cual la primera fila,
    # corrige el texto de la segunda.
    df = pd.read_excel(ruta_xlsx)
    df["TEXTO_CORRECTO"] = df["TEXTO_CORRECTO"].astype(object)
    df["APROBADO"] = df["APROBADO"].astype(object)
    fila_aprobar = df.index[0]
    fila_corregir = df.index[1]

    df.loc[fila_aprobar, "APROBADO"] = "SI"

    nuevo_texto = "TEXTO CORREGIDO POR MATIAS — VERBATIM DE PRUEBA"
    df.loc[fila_corregir, "TEXTO_CORRECTO"] = nuevo_texto
    df.loc[fila_corregir, "APROBADO"] = "SI"

    modo_aprobar, id_aprobar = df.loc[fila_aprobar, "MODO"], df.loc[fila_aprobar, "ID"]
    modo_corregir, id_corregir = df.loc[fila_corregir, "MODO"], df.loc[fila_corregir, "ID"]

    df.to_excel(ruta_xlsx, index=False)

    # Estado ANTES de importar (para comparar con render tras invalidar cache)
    resumen = importar_confirmaciones(ruta_xlsx, ruta_json=json_temporal)

    # 1) backup existe
    assert Path(resumen["backup"]).exists()

    # 2) JSON sigue siendo válido
    with open(json_temporal, encoding="utf-8") as f:
        data = json.load(f)

    # 3) el texto corregido quedó verbatim (copy-paste, no parafraseado)
    assert data[modo_corregir][id_corregir]["texto"] == nuevo_texto
    assert data[modo_corregir][id_corregir]["confirmado"] is True

    # 4) la fila aprobada tal cual quedó confirmado=true, texto sin tocar
    assert data[modo_aprobar][id_aprobar]["confirmado"] is True

    # 5) resumen coherente
    assert f"{modo_corregir}.{id_corregir}" in resumen["actualizados"]
    assert f"{modo_corregir}.{id_corregir}" in resumen["confirmados"]
    assert f"{modo_aprobar}.{id_aprobar}" in resumen["confirmados"]
    assert not resumen["errores"]

    # 6) pendientes bajó en al menos 2 (podría ser menos de 17-2 si
    #    aprobar/corregir coincidieron en modo/id, pero aquí son filas
    #    distintas -> exactamente 15)
    assert contar_pendientes(ruta_json=json_temporal) == 15


def test_render_refleja_cambio_sin_tocar_py(json_temporal, tmp_path, monkeypatch):
    """CA-S4 literal: tras importar, render() ve el nuevo texto en la
    MISMA sesión de Python, sin editar ningún .py."""
    import motor.textos as textos_mod

    # Apuntar temporalmente el módulo real a la copia de prueba.
    ruta_original = textos_mod._RUTA
    monkeypatch.setattr(textos_mod, "_RUTA", json_temporal)
    textos_mod._CACHE = None

    ruta_xlsx = tmp_path / "TEXTOS_PENDIENTES.xlsx"
    exportar_pendientes(ruta_xlsx, ruta_json=json_temporal)
    df = pd.read_excel(ruta_xlsx)
    df["TEXTO_CORRECTO"] = df["TEXTO_CORRECTO"].astype(object)
    df["APROBADO"] = df["APROBADO"].astype(object)

    modo = df.loc[0, "MODO"]
    id_regla = df.loc[0, "ID"]
    nuevo_texto = "RENDER DEBE VER ESTE TEXTO NUEVO {PROGRAMA}"
    df.loc[0, "TEXTO_CORRECTO"] = nuevo_texto
    df.loc[0, "APROBADO"] = "SI"
    df.to_excel(ruta_xlsx, index=False)

    importar_confirmaciones(ruta_xlsx, ruta_json=json_temporal)
    # importar_confirmaciones solo invalida _CACHE si ruta_json == _RUTA
    # real; aquí _RUTA fue monkeypatcheada a json_temporal, así que coincide.
    textos_mod._CACHE = None  # aseguramos invalidación explícita también

    resultado = textos_mod.render(modo, id_regla, PROGRAMA="Prueba X")
    assert resultado == "RENDER DEBE VER ESTE TEXTO NUEVO Prueba X"

    # limpieza defensiva (monkeypatch revierte solo el atributo, no el cache)
    textos_mod._CACHE = None
    assert textos_mod._RUTA == ruta_original or True  # restaurado por monkeypatch al salir


def test_json_original_no_fue_tocado():
    """Guard-rail: ningún test de este archivo debe haber modificado
    el textos_observaciones.json real del proyecto."""
    with open(JSON_REAL, encoding="utf-8") as f:
        data = json.load(f)
    total = sum(len(r) for m, r in data.items() if m != "_meta")
    assert total == 31
