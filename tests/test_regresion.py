#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
tests/test_regresion.py

Dos modos de operación:

1) REGRESIÓN REAL (en una máquina autorizada, con tres Excel de producción):
   Definir variables de entorno antes de correr pytest:

     CSMP_EXCEL_ESPERA=C:\\ruta\\excel_espera.xlsx
     CSMP_EXCEL_CUMPLIMIENTO=C:\\ruta\\excel_cumplimiento.xlsx   (con Hoja2)
     CSMP_EXCEL_INFORMES=C:\\ruta\\excel_informes.xlsx

   Windows (PowerShell):
     $env:CSMP_EXCEL_ESPERA="C:\ruta\excel_espera.xlsx"
     $env:CSMP_EXCEL_CUMPLIMIENTO="C:\ruta\excel_cumplimiento.xlsx"
     $env:CSMP_EXCEL_INFORMES="C:\ruta\excel_informes.xlsx"
     pytest tests/ -v

   Con estas 3 variables presentes, este archivo exige los conteos
   conocidos de producción (ESPERA=16, CUMPLIMIENTO=100, INFORMES=333 —
   ver PLAN v8.14 sección 3, CA-GLOBAL). Si tu volumen real cambió,
   actualiza CONTEOS_ESPERADOS abajo a propósito.

2) SMOKE SINTÉTICO (este sandbox, sin los Excel reales):
   Si las variables de entorno NO están definidas, se usan los fixtures
   generados por tests/fixtures/generar_fixtures.py (datos inventados).
   Solo se validan invariantes estructurales (no crashea, no quedan
   'ERROR' ni '..' en las observaciones, cuenta de filas preservada).
   Esto NO reemplaza la regresión real — ver advertencia en el reporte.
"""

import os
import sys
from pathlib import Path

import pytest

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

from tests.runner import ejecutar_modo_aislado  # noqa: E402

FIXTURES_DIR = Path(__file__).resolve().parent / "fixtures"

# Conteos de producción conocidos (PLAN v8.14, CA-GLOBAL). Solo se exigen
# cuando se corre contra los Excel reales.
CONTEOS_ESPERADOS = {"ESPERA": 16, "CUMPLIMIENTO": 100, "INFORMES": 333}


def _rutas_reales():
    e = os.environ.get("CSMP_EXCEL_ESPERA")
    c = os.environ.get("CSMP_EXCEL_CUMPLIMIENTO")
    i = os.environ.get("CSMP_EXCEL_INFORMES")
    if e and c and i:
        return {"ESPERA": e, "CUMPLIMIENTO": c, "INFORMES": i}
    return None


def _rutas_fixture():
    return {
        "ESPERA": str(FIXTURES_DIR / "excel_espera.xlsx"),
        "CUMPLIMIENTO": str(FIXTURES_DIR / "excel_cumplimiento.xlsx"),
        "INFORMES": str(FIXTURES_DIR / "excel_informes.xlsx"),
    }


MODO_REAL = _rutas_reales() is not None
RUTAS = _rutas_reales() or _rutas_fixture()


@pytest.fixture(scope="module")
def ruta_salida(tmp_path_factory):
    return str(tmp_path_factory.mktemp("salida_regresion"))


@pytest.mark.parametrize("modo", ["ESPERA", "CUMPLIMIENTO", "INFORMES"])
def test_pipeline_no_crashea(modo, ruta_salida):
    excel = RUTAS[modo]
    if not Path(excel).exists():
        pytest.skip(f"No existe {excel} — generar con "
                    f"tests/fixtures/generar_fixtures.py o definir env var")
    r = ejecutar_modo_aislado(modo, excel, ruta_salida)
    assert r["ok"], f"[{modo}] pipeline falló: {r.get('error') or r['log']}"
    assert r["n_obs_error"] == 0, (
        f"[{modo}] {r['n_obs_error']} observaciones con 'ERROR' — "
        f"log: {r['log']}"
    )
    assert r["n_obs_doblepunto"] == 0, (
        f"[{modo}] {r['n_obs_doblepunto']} observaciones con '..' "
        f"(fragmento vacío mal unido)"
    )


@pytest.mark.skipif(not MODO_REAL, reason=(
    "Requiere Excel reales de producción (CSMP_EXCEL_ESPERA / "
    "CSMP_EXCEL_CUMPLIMIENTO / CSMP_EXCEL_INFORMES). "
    "Correr en una máquina autorizada, no en este entorno de desarrollo."
))
@pytest.mark.parametrize("modo", ["ESPERA", "CUMPLIMIENTO", "INFORMES"])
def test_conteo_filas_produccion(modo, ruta_salida):
    r = ejecutar_modo_aislado(modo, RUTAS[modo], ruta_salida)
    esperado = CONTEOS_ESPERADOS[modo]
    assert r["n_filas"] == esperado, (
        f"[{modo}] se esperaban {esperado} filas, hay {r['n_filas']}. "
        f"Si el volumen real cambió intencionalmente, actualizar "
        f"CONTEOS_ESPERADOS en este archivo."
    )


def test_advertencia_modo_sintetico(capsys):
    """No falla nunca — solo deja constancia visible en la salida de
    pytest -v de si esta corrida usó datos reales o sintéticos."""
    if MODO_REAL:
        print("\n[INFO] Regresión corrida contra Excel REALES de producción.")
    else:
        print(
            "\n[ADVERTENCIA] Regresión corrida con FIXTURES SINTÉTICOS "
            "(datos inventados, no de producción). Los conteos 16/100/333 "
            "y el diff=0 vs baseline v8.13 NO fueron verificados en esta "
            "corrida. Definir CSMP_EXCEL_ESPERA / CSMP_EXCEL_CUMPLIMIENTO / "
            "CSMP_EXCEL_INFORMES y re-correr en la máquina real antes de "
            "dar por buena la regresión completa."
        )
    assert True


def test_columnas_salida_v9_documentadas():
    import pandas as pd
    from motor.procesador import _insertar_columnas_salida

    salida = _insertar_columnas_salida(pd.DataFrame({"ORIGEN": [1]}))
    assert list(salida.columns) == [
        "ORIGEN", "FECHA_OBS", "OBSERVACION", "TT", "CC", "RES"
    ]


def test_observaciones_sin_marcadores_v9():
    import pandas as pd
    from motor.procesador import calcular_preview

    prohibidos = ("{", "}", "..", "ERROR", "None")
    for modo, ruta in RUTAS.items():
        if not Path(ruta).exists():
            continue
        salida, error = calcular_preview(ruta, modo)
        assert error is None, f"[{modo}] {error}"
        observaciones = salida["OBSERVACION"].fillna("").astype(str)
        for marcador in prohibidos:
            assert not observaciones.str.contains(marcador, regex=False).any(), (
                f"[{modo}] marcador prohibido {marcador!r}"
            )
        assert not observaciones.str.contains(
            r"\bnan\b", case=False, regex=True
        ).any(), f"[{modo}] marcador prohibido 'nan'"
