#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tests/runner.py

Ejecuta motor.procesador.procesar() en un subprocess Python COMPLETAMENTE
AISLADO (proceso nuevo, sys.path limpio).

Motivo (lección aprendida, sesión 2026-07-07): correr los 3 modos dentro
del mismo proceso de pytest, importando módulos repetidamente, dejó
sys.path contaminado entre corridas y produjo un falso "0 diffs" que no
reflejaba el comportamiento real de una ejecución fresca del programa.

Uso como librería:
    from tests.runner import ejecutar_modo_aislado
    r = ejecutar_modo_aislado("ESPERA", "/ruta/excel_espera.xlsx", "/tmp/salida")
    r["ok"], r["n_filas"], r["archivo_salida"], r["log"]

Uso como script (NO llamar directo — es invocado por sí mismo vía subprocess):
    python tests/runner.py _worker MODO EXCEL_PATH RUTA_SALIDA RESULTADO_JSON
"""

import glob
import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path

RAIZ_PROYECTO = Path(__file__).resolve().parent.parent


def ejecutar_modo_aislado(modo: str, excel_path: str, ruta_salida: str,
                          timeout: int = 120) -> dict:
    """
    Corre procesar(modo) en subprocess aislado. Retorna:
    {
      "ok": bool,
      "archivo_salida": str|None,
      "n_filas": int,
      "n_obs_vacias": int,
      "n_obs_error": int,
      "n_obs_doblepunto": int,
      "log": [str, ...],
      "error": str|None,
    }
    """
    Path(ruta_salida).mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        resultado_path = Path(tmp) / "resultado.json"
        cmd = [
            sys.executable, str(Path(__file__).resolve()),
            "_worker", modo, str(excel_path), str(ruta_salida),
            str(resultado_path),
        ]
        proc = subprocess.run(
            cmd, cwd=str(RAIZ_PROYECTO),
            capture_output=True, text=True, timeout=timeout,
        )
        if not resultado_path.exists():
            return {
                "ok": False, "archivo_salida": None, "n_filas": 0,
                "n_obs_vacias": 0, "n_obs_error": 0, "n_obs_doblepunto": 0,
                "log": [], "error": (
                    f"El worker no produjo resultado.json (returncode="
                    f"{proc.returncode}).\nSTDOUT:\n{proc.stdout}\n"
                    f"STDERR:\n{proc.stderr}"
                ),
            }
        with open(resultado_path, encoding="utf-8") as f:
            return json.load(f)


def _worker(modo, excel_path, ruta_salida, resultado_path):
    """Solo se ejecuta DENTRO del subprocess aislado."""
    sys.path.insert(0, str(RAIZ_PROYECTO))
    import queue as _queue

    import pandas as pd

    from motor.procesador import procesar

    q = _queue.Queue()
    config = {"ruta_salida_excel": ruta_salida}
    logs = []
    ok = True
    error = None

    try:
        procesar(excel_path, modo, config, q)
    except Exception as e:  # noqa: BLE001
        ok = False
        error = f"{type(e).__name__}: {e}"

    while not q.empty():
        kind, data = q.get()
        logs.append(f"[{kind}] {data}")
        if kind == "done" and str(data).startswith("❌"):
            ok = False

    patron = str(Path(ruta_salida) / f"RUS_{modo}_*.xlsx")
    candidatos = sorted(glob.glob(patron), key=os.path.getmtime)
    archivo_salida = candidatos[-1] if candidatos else None

    n_filas = n_vacias = n_error = n_doblepunto = 0
    if archivo_salida:
        df = pd.read_excel(archivo_salida)
        n_filas = len(df)
        if "OBSERVACION" in df.columns:
            obs = df["OBSERVACION"].fillna("").astype(str)
            n_vacias = int((obs.str.strip() == "").sum())
            n_error = int(obs.str.contains("ERROR", regex=False).sum())
            n_doblepunto = int(obs.str.contains("..", regex=False).sum())
    else:
        ok = False

    with open(resultado_path, "w", encoding="utf-8") as f:
        json.dump({
            "ok": ok, "archivo_salida": archivo_salida, "n_filas": n_filas,
            "n_obs_vacias": n_vacias, "n_obs_error": n_error,
            "n_obs_doblepunto": n_doblepunto, "log": logs, "error": error,
        }, f, ensure_ascii=False)


if __name__ == "__main__":
    if len(sys.argv) == 6 and sys.argv[1] == "_worker":
        _, _, modo_, excel_, salida_, resultado_ = sys.argv
        _worker(modo_, excel_, salida_, resultado_)
    else:
        print(__doc__)
        sys.exit(1)
