# -*- coding: utf-8 -*-
"""
Generador de reportes de validación.
Produce: consola (texto) + archivo HTML autocontenido.
"""

from datetime import datetime
from html import escape
import os
import re
import uuid
from pathlib import Path
from motor.version import VERSION


# ─── Colores y etiquetas ──────────────────────────────────────────────────────

_COLORES = {
    "BLOQUEA":  {"bg": "#c0392b", "fg": "#ffffff", "badge": "🔴 BLOQUEA"},
    "ADVIERTE": {"bg": "#e67e22", "fg": "#ffffff", "badge": "🟡 ADVIERTE"},
    "OK":       {"bg": "#27ae60", "fg": "#ffffff", "badge": "✅ OK"},
}


# ─── Console ─────────────────────────────────────────────────────────────────

def imprimir_reporte(resultado: dict, *, incluir_detalle: bool = False) -> None:
    """Imprime resumen; los identificadores requieren opt-in explícito."""
    sep = "─" * 60
    puede = resultado["puede_procesar"]
    stats = resultado["estadisticas"]
    bloq  = resultado["anomalias_bloqueantes"]
    adv   = resultado["anomalias_advertencia"]

    print(f"\n{sep}")
    print(f"  VALIDACIÓN CSMP — {resultado.get('modo', '')}  "
          f"({stats.get('total_filas', 0)} filas)")
    print(sep)

    estado = "✅ PUEDE PROCESAR" if puede else "❌ BLOQUEADO — No se puede procesar"
    print(f"\n  Estado: {estado}")
    print(f"  Bloqueantes: {len(bloq)}  |  Advertencias: {len(adv)}\n")

    if bloq:
        print("  ── BLOQUEANTES ──────────────────────────")
        for a in bloq:
            print(f"  [{a['id']}] {a['descripcion']}  ({a['count']} afectados)")
            if incluir_detalle:
                for d in a["detalle"][:5]:
                    print(f"       • {d}")
                if len(a["detalle"]) > 5:
                    print(f"       ... y {len(a['detalle']) - 5} más")

    if adv:
        print("\n  ── ADVERTENCIAS ─────────────────────────")
        for a in adv:
            print(f"  [{a['id']}] {a['descripcion']}  ({a['count']} afectados)")
            if incluir_detalle:
                for d in a["detalle"][:3]:
                    print(f"       • {d}")
                if len(a["detalle"]) > 3:
                    print(f"       ... y {len(a['detalle']) - 3} más")

    if not bloq and not adv:
        print("  Sin anomalías detectadas. Excel limpio.")

    print(f"\n{sep}\n")


# ─── HTML ────────────────────────────────────────────────────────────────────

def _badge(severidad: str) -> str:
    c = _COLORES.get(severidad, _COLORES["ADVIERTE"])
    return (f'<span style="background:{c["bg"]};color:{c["fg"]};'
            f'padding:2px 8px;border-radius:4px;font-size:0.82em;'
            f'font-weight:bold">{c["badge"]}</span>')


def _seccion_anomalias(anomalias: list) -> str:
    if not anomalias:
        return ""
    html = ""
    for a in anomalias:
        color = _COLORES.get(a["severidad"], _COLORES["ADVIERTE"])
        items = "".join(f"<li>{escape(str(d), quote=True)}</li>" for d in a["detalle"])
        resto = ""
        if a["count"] > len(a["detalle"]):
            extra = a["count"] - len(a["detalle"])
            resto = f'<li style="color:#888">... y {extra} más (ver Excel completo)</li>'
        html += f"""
        <div class="anomalia" style="border-left:4px solid {color['bg']};
             background:#fafafa;margin:10px 0;padding:10px 14px;border-radius:4px">
          <div style="display:flex;align-items:center;gap:10px;margin-bottom:6px">
            {_badge(a['severidad'])}
            <strong>[{escape(str(a['id']), quote=True)}]</strong>
            <span>{escape(str(a['descripcion']), quote=True)}</span>
            <span style="margin-left:auto;color:#666;font-size:0.85em">
              {a['count']} afectado(s)
            </span>
          </div>
          <ul style="margin:4px 0 0 16px;padding:0;font-size:0.88em;color:#444">
            {items}{resto}
          </ul>
        </div>"""
    return html


def generar_html(resultado: dict, ruta_salida: str) -> str:
    """
    Genera reporte HTML en ruta_salida.
    Retorna ruta del archivo creado.
    """
    puede  = resultado["puede_procesar"]
    stats  = resultado["estadisticas"]
    bloq   = resultado["anomalias_bloqueantes"]
    adv    = resultado["anomalias_advertencia"]
    modo_original = str(resultado.get("modo", ""))
    modo   = escape(modo_original, quote=True)
    ts     = escape(str(resultado.get(
        "timestamp", datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    )), quote=True)

    estado_color = "#27ae60" if puede else "#c0392b"
    estado_texto = "✅ PUEDE PROCESAR" if puede else "❌ BLOQUEADO — Corrija anomalías antes de procesar"

    html_bloq = _seccion_anomalias(bloq) or '<p style="color:#27ae60">Sin bloqueantes ✓</p>'
    html_adv  = _seccion_anomalias(adv)  or '<p style="color:#27ae60">Sin advertencias ✓</p>'

    col_vals = "".join(
        f'<tr><td style="padding:3px 8px">{escape(str(k), quote=True)}</td>'
        f'<td style="padding:3px 8px;text-align:right"><strong>{escape(str(v), quote=True)}</strong></td></tr>'
        for k, v in stats.items()
    )

    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>Reporte Validación CSMP — {modo}</title>
<style>
  body {{font-family:'Segoe UI',Arial,sans-serif;background:#f0f2f5;
        color:#222;margin:0;padding:20px}}
  .card {{background:#fff;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,.08);
          padding:24px;margin-bottom:20px;max-width:860px;margin-left:auto;margin-right:auto}}
  h1 {{font-size:1.4em;margin:0 0 4px}}
  h2 {{font-size:1.05em;color:#555;border-bottom:1px solid #eee;
       padding-bottom:6px;margin-bottom:12px}}
  table {{border-collapse:collapse;font-size:0.9em}}
  td {{border-bottom:1px solid #eee}}
</style>
</head>
<body>

<div class="card">
  <h1>📋 Reporte de Validación — CSMP Assistant {VERSION}</h1>
  <p style="color:#888;font-size:0.88em;margin:0">
    Modo: <strong>{modo}</strong> &nbsp;|&nbsp;
    Generado: {ts} &nbsp;|&nbsp;
    Filas procesadas: <strong>{stats.get('total_filas', '?')}</strong>
  </p>
  <div style="margin-top:16px;padding:12px 18px;border-radius:6px;
              background:{estado_color};color:#fff;font-size:1.1em;font-weight:bold">
    {estado_texto}
  </div>
  <div style="margin-top:12px;display:flex;gap:20px;flex-wrap:wrap">
    <div style="padding:8px 16px;background:#fff3f3;border-radius:6px">
      🔴 Bloqueantes: <strong>{len(bloq)}</strong>
    </div>
    <div style="padding:8px 16px;background:#fff8f0;border-radius:6px">
      🟡 Advertencias: <strong>{len(adv)}</strong>
    </div>
    <div style="padding:8px 16px;background:#f0fff4;border-radius:6px">
      📊 Total filas: <strong>{stats.get('total_filas', '?')}</strong>
    </div>
  </div>
</div>

<div class="card">
  <h2>🔴 Anomalías Bloqueantes</h2>
  {html_bloq}
</div>

<div class="card">
  <h2>🟡 Advertencias</h2>
  {html_adv}
</div>

<div class="card">
  <h2>📊 Estadísticas del Excel</h2>
  <table><tbody>{col_vals}</tbody></table>
</div>

<p style="text-align:center;color:#aaa;font-size:0.8em;max-width:860px;margin:auto">
  CSMP Concepción — Generado automáticamente. Output supervisado por humano obligatorio.
</p>
</body>
</html>"""

    Path(ruta_salida).mkdir(parents=True, exist_ok=True)
    modo_archivo = re.sub(r"[^A-Za-z0-9_-]+", "_", modo_original).strip("_")
    modo_archivo = modo_archivo or "DESCONOCIDO"
    nombre = f"validacion_{modo_archivo}_{datetime.now():%Y%m%d_%H%M%S_%f}.html"
    ruta_archivo = Path(ruta_salida) / nombre
    temporal = ruta_archivo.with_name(
        f".{ruta_archivo.stem}.{uuid.uuid4().hex}.tmp.html"
    )
    try:
        with open(temporal, "w", encoding="utf-8") as f:
            f.write(html)
            f.flush()
            os.fsync(f.fileno())
        os.replace(temporal, ruta_archivo)
    except Exception:
        temporal.unlink(missing_ok=True)
        raise

    return str(ruta_archivo)
