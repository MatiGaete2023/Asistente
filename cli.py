#!/usr/bin/env python3
"""CLI operativo mínimo para flujos supervisados de CSMP RUS."""

import argparse
import queue
import sys
from pathlib import Path

import pandas as pd

from comunicaciones.generador_correos import GeneradorCorreos, crear_exportador_html
from motor.procesador import calcular_preview, procesar
from motor.utilidades import validar_archivo_excel
from resoluciones.generador_resoluciones import generar_resoluciones

_ROOT = Path(__file__).resolve().parent


def _leer_excel(path: str) -> pd.DataFrame:
    path_validado = validar_archivo_excel(path)
    engine = "xlrd" if path_validado.suffix.lower() == ".xls" else "openpyxl"
    df = pd.read_excel(path_validado, engine=engine)
    df.columns = [str(x).strip() for x in df.columns]
    return df.dropna(how="all").fillna("")


def _drenar(q: queue.Queue) -> tuple[list[str], str]:
    logs = []
    estado = ""
    while not q.empty():
        tipo, dato = q.get_nowait()
        if tipo == "log":
            logs.append(str(dato))
        elif tipo == "done":
            estado = str(dato)
    return logs, estado


def cmd_procesar(args) -> int:
    q: queue.Queue = queue.Queue()
    config = {"ruta_salida_excel": args.salida}
    procesar(args.entrada, args.modo, config, q)
    logs, estado = _drenar(q)
    for linea in logs:
        print(linea)
    if estado:
        print(estado)
    return 1 if estado.startswith("❌") else 0


def cmd_preview(args) -> int:
    df, error = calcular_preview(args.entrada, args.modo)
    if error:
        print(error, file=sys.stderr)
        return 1
    print(f"Preview {args.modo}: {len(df)} filas, {len(df.columns)} columnas")
    columnas = [c for c in ("RIT", "TRIBUNAL", "DERIVACION", "OBSERVACION") if c in df.columns]
    if columnas:
        print(df[columnas].head(args.filas).to_string(index=False))
    return 0


def cmd_correos(args) -> int:
    df = _leer_excel(args.entrada)
    catastro = str(_ROOT / "comunicaciones" / "catastro_programas.json")
    despachador = crear_exportador_html(args.dry_run_html) if args.dry_run_html else None
    gen = GeneradorCorreos({}, catastro, despachador=despachador)
    resultado = gen.procesar(df) if args.tipo == "informes" else gen.procesar_espera(df)
    for linea in resultado.get("detalle", []):
        print(linea)
    for aviso in resultado.get("errores", []) + resultado.get("grupos_sin_contacto", []):
        print(f"⚠️ {aviso}")
    creados = resultado.get("borradores_creados", 0)
    destino = "HTML" if args.dry_run_html else "Outlook"
    print(f"Borradores/exportaciones {destino}: {creados}")
    return 1 if creados == 0 and (resultado.get("errores") or resultado.get("grupos_sin_contacto")) else 0


def cmd_resoluciones(args) -> int:
    resultado = generar_resoluciones(args.entrada, args.salida)
    for error in resultado.get("errores", []):
        print(f"❌ {error}", file=sys.stderr)
    if resultado.get("archivo_generado"):
        print(resultado["archivo_generado"])
    print(
        f"Resoluciones: {resultado.get('total_resoluciones', 0)} | "
        f"omitidas: {resultado.get('omitidas', 0)} | fallidas: {resultado.get('fallidas', 0)}"
    )
    return 1 if resultado.get("errores") or resultado.get("fallidas") else 0


def construir_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="csmp-rus", description="CLI supervisado CSMP RUS")
    sub = parser.add_subparsers(dest="comando", required=True)

    p = sub.add_parser("procesar", help="Procesa un Excel con el motor RUS")
    p.add_argument("--modo", required=True, choices=["ESPERA", "CUMPLIMIENTO", "INFORMES"])
    p.add_argument("--entrada", required=True)
    p.add_argument("--salida", required=True)
    p.set_defaults(func=cmd_procesar)

    p = sub.add_parser("preview", help="Calcula vista previa sin escribir Excel")
    p.add_argument("--modo", required=True, choices=["ESPERA", "CUMPLIMIENTO", "INFORMES"])
    p.add_argument("--entrada", required=True)
    p.add_argument("--filas", type=int, default=5)
    p.set_defaults(func=cmd_preview)

    p = sub.add_parser("correos", help="Genera borradores Outlook o HTML dry-run")
    p.add_argument("--tipo", required=True, choices=["informes", "espera"])
    p.add_argument("--entrada", required=True)
    p.add_argument("--dry-run-html")
    p.set_defaults(func=cmd_correos)

    p = sub.add_parser("resoluciones", help="Genera Word consolidado de resoluciones")
    p.add_argument("--entrada", required=True)
    p.add_argument("--salida", required=True)
    p.set_defaults(func=cmd_resoluciones)
    return parser


def main(argv=None) -> int:
    parser = construir_parser()
    args = parser.parse_args(argv)
    try:
        return args.func(args)
    except Exception as exc:  # noqa: BLE001 - CLI muestra error limpio por defecto
        print(f"Error: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
