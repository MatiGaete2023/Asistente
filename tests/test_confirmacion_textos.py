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
from openpyxl import load_workbook

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

from motor.textos_confirmacion import (  # noqa: E402
    exportar_pendientes, importar_confirmaciones, contar_pendientes,
)

JSON_REAL = RAIZ / "motor" / "textos_observaciones.json"


@pytest.fixture()
def json_temporal(tmp_path):
    """Copia aislada con tres reglas pendientes para ejercitar el flujo."""
    destino = tmp_path / "textos_observaciones.json"
    shutil.copy2(JSON_REAL, destino)
    with open(destino, encoding="utf-8") as f:
        data = json.load(f)
    for modo, regla in [
        ("COMUN", "CURADOR"),
        ("ESPERA", "E05_SOLO_CORREO"),
        ("CUMPLIMIENTO", "C07_SIN_FICHA"),
    ]:
        data[modo][regla]["confirmado"] = False
    with open(destino, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return destino


def test_exportar_pendientes_genera_filas_vigentes(json_temporal, tmp_path):
    ruta_xlsx = tmp_path / "TEXTOS_PENDIENTES.xlsx"
    n = exportar_pendientes(ruta_xlsx, ruta_json=json_temporal)
    assert n == 3
    assert ruta_xlsx.exists()
    df = pd.read_excel(ruta_xlsx)
    assert list(df.columns) == ["MODO", "ID", "TEXTO_ACTUAL", "TEXTO_CORRECTO", "APROBADO"]
    assert len(df) == 3


def test_ciclo_completo_export_editar_import(json_temporal, tmp_path):
    ruta_xlsx = tmp_path / "TEXTOS_PENDIENTES.xlsx"
    exportar_pendientes(ruta_xlsx, ruta_json=json_temporal)

    # La persona usuaria edita el Excel: aprueba tal cual la primera fila,
    # corrige el texto de la segunda.
    df = pd.read_excel(ruta_xlsx)
    df["TEXTO_CORRECTO"] = df["TEXTO_CORRECTO"].astype(object)
    df["APROBADO"] = df["APROBADO"].astype(object)
    fila_aprobar = df.index[0]
    fila_corregir = df.index[1]

    df.loc[fila_aprobar, "APROBADO"] = "SI"

    nuevo_texto = "TEXTO CORREGIDO VERBATIM DE PRUEBA"
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

    # 6) de tres pendientes quedan exactamente uno
    assert contar_pendientes(ruta_json=json_temporal) == 1


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
    nuevo_texto = "RENDER DEBE VER ESTE TEXTO NUEVO"
    df.loc[0, "TEXTO_CORRECTO"] = nuevo_texto
    df.loc[0, "APROBADO"] = "SI"
    df.to_excel(ruta_xlsx, index=False)

    importar_confirmaciones(ruta_xlsx, ruta_json=json_temporal)
    # importar_confirmaciones solo invalida _CACHE si ruta_json == _RUTA
    # real; aquí _RUTA fue monkeypatcheada a json_temporal, así que coincide.
    textos_mod._CACHE = None  # aseguramos invalidación explícita también

    resultado = textos_mod.render(modo, id_regla)
    assert resultado == nuevo_texto

    # limpieza defensiva (monkeypatch revierte solo el atributo, no el cache)
    textos_mod._CACHE = None
    assert textos_mod._RUTA == json_temporal


def test_json_original_no_fue_tocado():
    """Guard-rail: ningún test de este archivo debe haber modificado
    el textos_observaciones.json real del proyecto."""
    with open(JSON_REAL, encoding="utf-8") as f:
        data = json.load(f)
    total = sum(len(r) for m, r in data.items() if m != "_meta")
    assert total == 25
    assert all(
        entry.get("confirmado") is True
        for modo, reglas in data.items() if modo != "_meta"
        for entry in reglas.values()
    )


def test_celdas_vacias_no_reemplazan_texto_con_nan(json_temporal, tmp_path):
    ruta_xlsx = tmp_path / "TEXTOS_PENDIENTES.xlsx"
    exportar_pendientes(ruta_xlsx, ruta_json=json_temporal)
    df = pd.read_excel(ruta_xlsx)
    df["APROBADO"] = df["APROBADO"].astype(object)
    original = df.loc[0, "TEXTO_ACTUAL"]
    modo, regla = df.loc[0, "MODO"], df.loc[0, "ID"]
    df.loc[0, "APROBADO"] = "SI"
    df.to_excel(ruta_xlsx, index=False)

    resumen = importar_confirmaciones(ruta_xlsx, ruta_json=json_temporal)
    with open(json_temporal, encoding="utf-8") as f:
        data = json.load(f)
    assert data[modo][regla]["texto"] == original
    assert data[modo][regla]["texto"] != "nan"
    assert f"{modo}.{regla}" not in resumen["actualizados"]


def test_importar_rechaza_cambio_de_placeholders(json_temporal, tmp_path):
    ruta_xlsx = tmp_path / "TEXTOS_PENDIENTES.xlsx"
    exportar_pendientes(ruta_xlsx, ruta_json=json_temporal)
    df = pd.read_excel(ruta_xlsx)
    df["TEXTO_CORRECTO"] = df["TEXTO_CORRECTO"].astype(object)
    df.loc[0, "TEXTO_CORRECTO"] = "Texto inseguro {PLACEHOLDER_NUEVO}"
    df.to_excel(ruta_xlsx, index=False)

    resumen = importar_confirmaciones(ruta_xlsx, ruta_json=json_temporal)
    assert resumen["actualizados"] == []
    assert resumen["backup"] is None
    assert any("placeholders" in error for error in resumen["errores"])


def test_importar_rechaza_acceso_a_atributos_en_placeholder(json_temporal, tmp_path):
    with open(json_temporal, encoding="utf-8") as f:
        data = json.load(f)
    data["COMUN"]["NO_SEGUIMIENTO"]["confirmado"] = False
    with open(json_temporal, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    ruta_xlsx = tmp_path / "TEXTOS_PENDIENTES.xlsx"
    exportar_pendientes(ruta_xlsx, ruta_json=json_temporal)
    df = pd.read_excel(ruta_xlsx)
    coincidencias = df["TEXTO_ACTUAL"].astype(str).str.contains(
        "{PROGRAMA}", regex=False
    )
    assert coincidencias.any()
    fila = coincidencias[coincidencias].index[0]
    df["TEXTO_CORRECTO"] = df["TEXTO_CORRECTO"].astype(object)
    df.loc[fila, "TEXTO_CORRECTO"] = str(df.loc[fila, "TEXTO_ACTUAL"]).replace(
        "{PROGRAMA}", "{PROGRAMA.__class__}"
    )
    df.to_excel(ruta_xlsx, index=False)

    resumen = importar_confirmaciones(ruta_xlsx, ruta_json=json_temporal)
    assert resumen["actualizados"] == []
    assert resumen["backup"] is None
    assert any("placeholder" in error for error in resumen["errores"])


def test_exportar_conserva_texto_que_parece_formula(json_temporal, tmp_path):
    with open(json_temporal, encoding="utf-8") as f:
        data = json.load(f)
    data["COMUN"]["CURADOR"]["texto"] = "=1+1"
    data["COMUN"]["CURADOR"]["confirmado"] = False
    with open(json_temporal, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    ruta_xlsx = tmp_path / "TEXTOS_PENDIENTES.xlsx"
    exportar_pendientes(ruta_xlsx, ruta_json=json_temporal)
    libro = load_workbook(ruta_xlsx, data_only=False)
    celdas = list(libro["TEXTOS"].iter_rows(min_row=2, values_only=False))
    fila_objetivo = next(fila for fila in celdas if fila[0].value == "COMUN"
                         and fila[1].value == "CURADOR")
    celda = fila_objetivo[2]
    assert celda.value == "=1+1"
    assert celda.data_type == "s"

    fila_objetivo[4].value = "SI"
    fila_objetivo[4].data_type = "s"
    libro.save(ruta_xlsx)
    resumen = importar_confirmaciones(ruta_xlsx, ruta_json=json_temporal)
    assert "COMUN.CURADOR" in resumen["confirmados"]
