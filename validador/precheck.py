# -*- coding: utf-8 -*-
"""
precheck.py — Validador pre-procesamiento CSMP Assistant v8.0

Función pública: validar_excel(df, modo, ruta_salida_reportes) → dict

Retorna:
{
    "puede_procesar": bool,
    "anomalias_bloqueantes": [...],
    "anomalias_advertencia": [...],
    "estadisticas": {...},
    "ruta_reporte_html": str | None,
    "modo": str,
    "timestamp": str,
}
"""

from datetime import datetime
from .reglas_validacion import (
    a1_columnas_minimas,
    a2_rit_duplicados,
    a3_derivacion_vacia,
    a4_fechas_futuras,
    a5_edades_inconsistentes,
    a6_edad_fuera_rango,
    a7_campos_vacios_masivos,
    a8_tribunal_no_reconocido,
)
from .reporte import imprimir_reporte, generar_html

# Reglas bloqueantes ejecutadas en orden (se detiene solo si A1 bloquea col mínimas,
# para evitar errores en cascada)
_BLOQUEANTES = [a1_columnas_minimas, a2_rit_duplicados, a3_derivacion_vacia]

# Reglas de advertencia (todas se ejecutan siempre)
_ADVERTENCIAS = [
    a4_fechas_futuras,
    a5_edades_inconsistentes,
    a6_edad_fuera_rango,
    a7_campos_vacios_masivos,
    a8_tribunal_no_reconocido,
]


def _estadisticas(df, modo: str) -> dict:
    stats = {
        "total_filas": len(df),
        "total_columnas": len(df.columns),
    }
    # Filas completamente vacías
    stats["filas_vacias"] = int(df.isnull().all(axis=1).sum())
    # % completitud global
    total_celdas = len(df) * len(df.columns)
    if total_celdas > 0:
        vacias = df.isnull().sum().sum() + (df == "").sum().sum()
        stats["completitud_global"] = f"{100 - (vacias / total_celdas * 100):.1f}%"
    else:
        stats["completitud_global"] = "N/A"
    return stats


def validar_excel(df, modo: str, ruta_salida_reportes: str = None) -> dict:
    """
    Valida DataFrame antes del procesamiento.

    Args:
        df: DataFrame cargado desde Excel.
        modo: "ESPERA" | "CUMPLIMIENTO" | "INFORMES"
        ruta_salida_reportes: directorio donde guardar reporte HTML.
                              Si None, no genera archivo HTML.

    Returns:
        dict con estructura documentada en módulo.
    """
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    modo = modo.upper()

    bloqueantes = []
    advertencias = []

    # ── Ejecutar bloqueantes ─────────────────────────────────────────────────
    a1 = a1_columnas_minimas(df, modo)
    if a1:
        bloqueantes.append(a1)
        # Si faltan columnas mínimas, las demás reglas pueden crashear.
        # Solo ejecutamos A7 y A8 que no dependen de columnas específicas.
        adv_seguras = [a7_campos_vacios_masivos, a8_tribunal_no_reconocido]
        for fn in adv_seguras:
            r = fn(df, modo)
            if r:
                advertencias.append(r)
    else:
        # A1 OK → ejecutar resto de bloqueantes
        for fn in [a2_rit_duplicados, a3_derivacion_vacia]:
            r = fn(df, modo)
            if r:
                bloqueantes.append(r)

        # Todas las advertencias
        for fn in _ADVERTENCIAS:
            r = fn(df, modo)
            if r:
                advertencias.append(r)

    puede_procesar = len(bloqueantes) == 0

    resultado = {
        "puede_procesar": puede_procesar,
        "anomalias_bloqueantes": bloqueantes,
        "anomalias_advertencia": advertencias,
        "estadisticas": _estadisticas(df, modo),
        "ruta_reporte_html": None,
        "modo": modo,
        "timestamp": ts,
    }

    # ── Reporte consola ──────────────────────────────────────────────────────
    imprimir_reporte(resultado)

    # ── Reporte HTML ─────────────────────────────────────────────────────────
    if ruta_salida_reportes:
        try:
            ruta_html = generar_html(resultado, ruta_salida_reportes)
            resultado["ruta_reporte_html"] = ruta_html
        except Exception as e:
            print(f"[VALIDADOR] ⚠ No se pudo generar HTML: {e}")

    return resultado
