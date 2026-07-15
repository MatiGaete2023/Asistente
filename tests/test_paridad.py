#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tests/test_paridad.py

Compara render(modo, id_regla, **placeholders) contra un golden CONGELADO
en tests/goldens_textos.json (no contra el propio JSON leido en caliente
-- eso seria tautologico y nunca detectaria un sabotaje).

Si alguien edita motor/textos_observaciones.json sin pasar por el flujo
de confirmacion (tests/generar_goldens.py), este test FALLA.

Tras confirmar un texto deliberadamente (flujo S4 + aprobación humana):
    python tests/generar_goldens.py
Y este test vuelve a estar verde con el nuevo texto como referencia.
"""

import json
import sys
from pathlib import Path

import pytest

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

from motor.textos import (  # noqa: E402
    ErrorCatalogoTextos,
    render,
    _cargar,
    listar_no_confirmados,
)

RUTA_GOLDEN = Path(__file__).resolve().parent / "goldens_textos.json"


def _cargar_golden():
    if not RUTA_GOLDEN.exists():
        pytest.fail(
            "tests/goldens_textos.json no existe. "
            "Correr: python tests/generar_goldens.py"
        )
    with open(RUTA_GOLDEN, encoding="utf-8") as f:
        return json.load(f)


def _casos_golden():
    golden = _cargar_golden()
    placeholders = golden["_meta"]["placeholders_usados"]
    casos = []
    for modo, reglas in golden.items():
        if modo == "_meta":
            continue
        for id_regla, esperado in reglas.items():
            casos.append((modo, id_regla, esperado, placeholders))
    return casos


@pytest.mark.parametrize(
    "modo,id_regla,esperado,placeholders",
    _casos_golden(),
    ids=[f"{m}.{i}" for m, i, _, _ in _casos_golden()],
)
def test_render_coincide_con_golden(modo, id_regla, esperado, placeholders):
    obtenido = render(modo, id_regla, **placeholders)
    assert obtenido == esperado["texto_renderizado"], (
        f"[{modo}.{id_regla}] render() cambio respecto al golden.\n"
        f"  esperado : {esperado['texto_renderizado']!r}\n"
        f"  obtenido : {obtenido!r}\n"
        f"  Si el cambio es intencional (texto confirmado por una persona), "
        f"correr: python tests/generar_goldens.py"
    )


@pytest.mark.parametrize(
    "modo,id_regla,esperado,placeholders",
    _casos_golden(),
    ids=[f"{m}.{i}" for m, i, _, _ in _casos_golden()],
)
def test_texto_fuente_no_mutado(modo, id_regla, esperado, placeholders):
    """Chequeo redundante: compara directamente el string fuente del JSON
    (sin pasar por render/format_map), para aislar si el cambio esta en
    el texto crudo o en el motor de templating."""
    data = _cargar()
    texto_actual = data[modo][id_regla]["texto"]
    assert texto_actual == esperado["texto_fuente"], (
        f"[{modo}.{id_regla}] texto FUENTE cambio respecto al golden "
        f"(edicion directa del JSON sin regenerar goldens)."
    )


def test_ningun_id_huerfano_o_faltante_vs_golden():
    """Todo ID definido en el JSON debe tener golden, y viceversa.
    Detecta reglas agregadas/eliminadas sin actualizar el golden."""
    data = _cargar()
    golden = _cargar_golden()

    ids_json = set()
    for modo, reglas in data.items():
        if modo == "_meta":
            continue
        for id_regla in reglas:
            ids_json.add((modo, id_regla))

    ids_golden = set()
    for modo, reglas in golden.items():
        if modo == "_meta":
            continue
        for id_regla in reglas:
            ids_golden.add((modo, id_regla))

    faltan_en_golden = ids_json - ids_golden
    sobran_en_golden = ids_golden - ids_json

    assert not faltan_en_golden, (
        f"IDs nuevos en el JSON sin golden: {faltan_en_golden}. "
        f"Correr: python tests/generar_goldens.py"
    )
    assert not sobran_en_golden, (
        f"IDs en el golden que ya no existen en el JSON: {sobran_en_golden}. "
        f"Correr: python tests/generar_goldens.py"
    )


def test_conteo_total_textos_es_25():
    """Guard-rail explícito del inventario vigente."""
    data = _cargar()
    total = sum(len(reglas) for modo, reglas in data.items() if modo != "_meta")
    assert total == 25, f"Se esperaban 31 reglas totales, hay {total}"


def test_textos_no_confirmados_documentados():
    """Guard-rail informativo (no falla el build): imprime el estado actual
    de F-7 para que quede visible en cada corrida de pytest -v."""
    pendientes = listar_no_confirmados()
    print(f"\n[INFO] Textos con confirmado=False: {len(pendientes)}")
    for modo, id_regla in pendientes:
        print(f"        - {modo}.{id_regla}")
    assert pendientes == []


def test_render_falla_si_falta_placeholder():
    with pytest.raises(ErrorCatalogoTextos, match="FECHA_AUDIENCIA"):
        render("COMUN", "PROX_AUDIENCIA")


def test_render_falla_si_no_existe_regla():
    with pytest.raises(ErrorCatalogoTextos, match="Texto no definido"):
        render("COMUN", "NO_EXISTE")
