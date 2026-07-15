#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tests/generar_goldens.py

Congela el estado ACTUAL de motor/textos_observaciones.json en
tests/goldens_textos.json. Esto es lo que test_paridad.py usa como
referencia fija para detectar cambios NO intencionales al JSON.

NO se ejecuta automaticamente con pytest. Se ejecuta a mano, a proposito,
solo en dos casos:
  1. Al establecer una nueva versión aprobada del catálogo.
  2. Cada vez que la persona usuaria confirme/corrija un texto vía S4
     (importar TEXTOS_PENDIENTES.xlsx) y se quiera que el nuevo texto
     pase a ser el golden vigente.

Uso:
    python tests/generar_goldens.py

Placeholders sinteticos fijos (no son datos reales de ningun NNA):
    ver PLACEHOLDERS_SINTETICOS abajo.
"""

import json
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

from motor.textos import render, _cargar  # noqa: E402

RUTA_GOLDEN = Path(__file__).resolve().parent / "goldens_textos.json"

# Placeholders sinteticos: cubren TODOS los nombres de campo usados en
# cualquier texto del JSON. Un texto que solo usa un subconjunto los ignora.
PLACEHOLDERS_SINTETICOS = {
    "PNOMBRE":            "Juan",
    "FECHA_MAYORIA":       "10 de agosto de 2026",
    "DIAS_ESPERA":         "45",
    "PROGRAMA":            "PRM Esperanza",
    "FECHA_RESOLUCION":    "1 de junio de 2026",
    "FECHA_EGRESO_PROY":   "15 de julio de 2026",
    "FECHA_PROX_INFORME":  "20 de septiembre de 2026",
    "FECHA_INGRESO":       "5 de mayo de 2026",
    "FECHA_OIDO":          "12 de junio de 2026",
    "FECHA_AUDIENCIA":     "22 de julio de 2026",
    "FECHA_FICHA_INDIVIDUAL": "3 de marzo de 2026",
    "FECHA_EGRESO_PROYECTADO": "15 de julio de 2026",
    "FECHA_VENCIMIENTO":   "30 de agosto de 2026",
}


def generar():
    data = _cargar()
    golden = {
        "_meta": {
            "generado_por": "tests/generar_goldens.py",
            "advertencia": "NO editar a mano. Regenerar solo tras confirmar "
                            "un texto via flujo S4, con revision explicita.",
            "placeholders_usados": PLACEHOLDERS_SINTETICOS,
        }
    }
    total = 0
    for modo, reglas in data.items():
        if modo == "_meta":
            continue
        golden[modo] = {}
        for id_regla, entry in reglas.items():
            texto_fuente = entry["texto"]
            renderizado = render(modo, id_regla, **PLACEHOLDERS_SINTETICOS)
            golden[modo][id_regla] = {
                "texto_fuente": texto_fuente,
                "texto_renderizado": renderizado,
                "confirmado": bool(entry.get("confirmado", False)),
            }
            total += 1

    with open(RUTA_GOLDEN, "w", encoding="utf-8") as f:
        json.dump(golden, f, ensure_ascii=False, indent=2)

    print(f"OK: {total} goldens escritos en {RUTA_GOLDEN}")


if __name__ == "__main__":
    generar()
