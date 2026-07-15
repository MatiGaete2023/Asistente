#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gui/app.py — v8.8
Integra en pestañas Tkinter:
  - Pestaña 1: Motor RUS (ESPERA / CUMPLIMIENTO / INFORMES)
  - Pestaña 2: Correos Outlook (Informes y Lista de espera)
  - Pestaña 3: Resoluciones .docx
"""

import sys
import os
from pathlib import Path

# Raíz del proyecto en sys.path — independiente del punto de entrada
_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import queue
import json
from datetime import datetime

try:
    import pandas as pd
except ImportError:
    _r = tk.Tk(); _r.withdraw()
    messagebox.showerror("Dependencia faltante",
        "No se encontró pandas.\nEjecuta: pip install pandas openpyxl xlrd python-dateutil")
    raise SystemExit(1)

try:
    from motor.procesador import procesar, calcular_preview
    from motor.mapeo_columnas import mapear_columnas
    from motor.textos_confirmacion import (
        exportar_pendientes, importar_confirmaciones, contar_pendientes
    )
except ImportError as e:
    _r = tk.Tk(); _r.withdraw()
    messagebox.showerror("Error de instalación",
        f"No se pudo cargar el motor:\n{e}\n\nEjecuta desde la carpeta raíz o usa main.py.")
    raise SystemExit(1)

CONFIG_FILE   = "config_rus.json"
CONTACTOS_FILE = "contactos.json"
VERSION       = "v8.14.0"
MAX_LOG_LINES = 500

_BASE = Path.home() / "CSMP_RUS"
DEFAULT_CONFIG = {
    "ruta_entrada_excel":  str(_BASE / "entrada" / "ENTRADA.xlsx"),
    "ruta_salida_excel":   str(_BASE / "salida"),
    "ruta_logs":           str(_BASE / "logs"),
    "cc_fijo":             "ucc_concepcion@pjud.cl",
    "dias_retencion_logs": 90
}
DEFAULT_CONTACTOS = {
    "LAJA":    ["abergman@pjud.cl", "ammillan@pjud.cl"],
    "MULCHEN": ["cavillanueva@pjud.cl", "ialister@pjud.cl"],
    "TOME":    ["nmojeda@pjud.cl", "csrain@pjud.cl"]
}

TAG_OK   = "ok"
TAG_ERR  = "err"
TAG_WARN = "warn"
TAG_INFO = "info"
TAG_DONE = "done_tag"


def _cargar_json(path, default):
    if not os.path.exists(path):
        try:
            os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
            with open(path, "w", encoding="utf-8") as f:
                json.dump(default, f, indent=4, ensure_ascii=False)
        except Exception:
            pass
        return default.copy()
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default.copy()


def _guardar_json(path, data):
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
    except Exception:
        pass


# ══════════════════════════════════════════════════════════════════════════════
# APP PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════

class RUSApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"Asistente de RUS — CSMP Concepción  {VERSION}")
        self.geometry("1050x800")
        self.minsize(860, 640)

        self.q          = queue.Queue()
        self.running    = False
        self.cfg        = _cargar_json(CONFIG_FILE,    DEFAULT_CONFIG)
        self.contactos  = _cargar_json(CONTACTOS_FILE, DEFAULT_CONTACTOS)

        self._build_ui()
        self.after(100, self._process_queue)

    # ── encabezado + pestañas ─────────────────────────────────────────────────

    def _build_ui(self):
        # Encabezado
        top = ttk.Frame(self, padding=(12, 10, 12, 4))
        top.pack(fill="x")
        ttk.Label(top, text="Asistente de RUS",
                  font=("Segoe UI", 20, "bold")).pack(anchor="w")
        ttk.Label(top,
                  text="Centro de Seguimiento de Medidas de Protección — Concepción",
                  font=("Segoe UI", 9), foreground="#555555").pack(anchor="w")
        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=12, pady=(0, 4))

        # Notebook
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=12, pady=(0, 6))

        self._tab_motor      = ttk.Frame(nb, padding=8)
        self._tab_correos    = ttk.Frame(nb, padding=8)
        self._tab_resoluciones = ttk.Frame(nb, padding=8)

        nb.add(self._tab_motor,        text="  Motor RUS  ")
        nb.add(self._tab_correos,      text="  Correos Outlook  ")
        nb.add(self._tab_resoluciones, text="  Resoluciones  ")

        self._build_tab_motor(self._tab_motor)
        self._build_tab_correos(self._tab_correos)
        self._build_tab_resoluciones(self._tab_resoluciones)

        # Barra de estado global
        status_bar = ttk.Frame(self)
        status_bar.pack(fill="x", padx=12, pady=(0, 4))
        self.status = tk.StringVar(value="Listo")
        ttk.Label(status_bar, textvariable=self.status,
                  font=("Segoe UI", 9), foreground="#444444").pack(side="left")
        ttk.Label(status_bar, text=VERSION,
                  font=("Segoe UI", 9), foreground="#aaaaaa").pack(side="right")

    # ── TAB 1: Motor RUS ──────────────────────────────────────────────────────

    def _build_tab_motor(self, parent):
        # Configuración
        cfg = ttk.LabelFrame(parent, text="Configuración", padding=8)
        cfg.pack(fill="x", pady=(0, 6))

        self.in_var  = tk.StringVar(value=self.cfg.get("ruta_entrada_excel", ""))
        self.out_var = tk.StringVar(value=self.cfg.get("ruta_salida_excel",  ""))

        for i, (lbl, var, cmd) in enumerate([
            ("Excel entrada:",  self.in_var,  self._pick_in),
            ("Carpeta salida:", self.out_var, self._pick_out),
        ]):
            ttk.Label(cfg, text=lbl, width=14, anchor="w").grid(row=i, column=0, sticky="w", pady=2)
            ttk.Entry(cfg, textvariable=var, width=76).grid(row=i, column=1, padx=5, sticky="ew")
            ttk.Button(cfg, text="Buscar…", width=8, command=cmd).grid(row=i, column=2)
        cfg.columnconfigure(1, weight=1)
        ttk.Button(cfg, text="💾  Guardar configuración",
                   command=self._save_cfg).grid(row=2, column=0, columnspan=3,
                                                sticky="ew", pady=(6, 0))

        # Botones proceso
        proc = ttk.LabelFrame(parent, text="Procesar", padding=8)
        proc.pack(fill="x", pady=(0, 6))
        self._motor_btns = []
        for i, (txt, modo) in enumerate([
            ("▶  ESPERA",       "ESPERA"),
            ("▶  CUMPLIMIENTO", "CUMPLIMIENTO"),
            ("▶  INFORMES",     "INFORMES"),
            ("📁  Abrir salida", None),
        ]):
            cmd = (lambda m=modo: self._start_motor(m)) if modo else self._open_out
            b = ttk.Button(proc, text=txt, command=cmd)
            b.grid(row=0, column=i, padx=4, pady=4, sticky="ew")
            self._motor_btns.append(b)

        # S3 (v8.14): vista previa en memoria, sin escribir archivo
        self._preview_btns = []
        for i, (txt, modo) in enumerate([
            ("👁  Vista previa ESPERA",       "ESPERA"),
            ("👁  Vista previa CUMPLIMIENTO", "CUMPLIMIENTO"),
            ("👁  Vista previa INFORMES",     "INFORMES"),
        ]):
            b = ttk.Button(proc, text=txt, command=lambda m=modo: self._start_preview(m))
            b.grid(row=1, column=i, padx=4, pady=(0, 4), sticky="ew")
            self._preview_btns.append(b)

        proc.columnconfigure((0,1,2,3), weight=1)

        # S4 (v8.14): confirmación de textos sin sesiones de chat
        conf = ttk.LabelFrame(parent, text="Confirmación de textos", padding=8)
        conf.pack(fill="x", pady=(0, 6))

        self._pendientes_var = tk.StringVar()
        self._actualizar_badge_pendientes()
        ttk.Label(conf, textvariable=self._pendientes_var,
                  font=("Segoe UI", 9)).grid(row=0, column=0, columnspan=2,
                                              sticky="w", pady=(0, 4))
        ttk.Button(conf, text="📤  Exportar pendientes…",
                   command=self._exportar_textos_pendientes).grid(
                       row=1, column=0, padx=4, sticky="ew")
        ttk.Button(conf, text="📥  Importar confirmaciones…",
                   command=self._importar_textos_confirmados).grid(
                       row=1, column=1, padx=4, sticky="ew")
        conf.columnconfigure((0, 1), weight=1)

        # Progress
        self.pb = ttk.Progressbar(parent, mode="determinate", maximum=100)
        self.pb.pack(fill="x", pady=(0, 4))

        # Bitácora
        self._log_frame = ttk.LabelFrame(parent, text="Bitácora", padding=8)
        self._log_frame.pack(fill="both", expand=True)
        self._build_log(self._log_frame)

    # ── TAB 2: Correos Outlook ────────────────────────────────────────────────

    def _build_tab_correos(self, parent):
        ttk.Label(parent,
                  text="Genera borradores en Outlook. Requiere Outlook instalado y abierto.",
                  font=("Segoe UI", 9), foreground="#666666").pack(anchor="w", pady=(0, 6))

        # Excel de entrada (compartido con motor o propio)
        inp = ttk.LabelFrame(parent, text="Excel de entrada", padding=8)
        inp.pack(fill="x", pady=(0, 6))
        self.correo_in_var = tk.StringVar(value=self.cfg.get("ruta_entrada_excel", ""))
        ttk.Label(inp, text="Archivo Excel:", width=14, anchor="w").grid(row=0, column=0, sticky="w")
        ttk.Entry(inp, textvariable=self.correo_in_var, width=76).grid(row=0, column=1, padx=5, sticky="ew")
        ttk.Button(inp, text="Buscar…", width=8,
                   command=lambda: self._pick_file(self.correo_in_var)).grid(row=0, column=2)
        inp.columnconfigure(1, weight=1)

        # Botones
        btns = ttk.LabelFrame(parent, text="Tipo de correos", padding=8)
        btns.pack(fill="x", pady=(0, 6))
        self._correo_btns = []

        acciones = [
            ("✉  Informes vencidos / por vencer",
             "Usa Excel de INFORMES (con columna FECHA VENCIMIENTO).\n"
             "Genera borradores: 1 por programa (vencidos) + 1 por tribunal + 1 por programa (por vencer).",
             self._run_correos_informes),
            ("✉  Lista de espera",
             "Usa Excel de ESPERA procesado por el motor (con columna OBSERVACION).\n"
             "Genera 1 borrador por programa+tribunal con NNA en espera.",
             self._run_correos_espera),
        ]
        for i, (txt, tooltip, cmd) in enumerate(acciones):
            frame = ttk.Frame(btns)
            frame.grid(row=i, column=0, sticky="ew", pady=3)
            b = ttk.Button(frame, text=txt, command=cmd, width=38)
            b.pack(side="left", padx=(0, 8))
            ttk.Label(frame, text=tooltip, font=("Segoe UI", 8),
                      foreground="#777777", wraplength=520, justify="left").pack(side="left")
            self._correo_btns.append(b)
        btns.columnconfigure(0, weight=1)

        # Log correos
        log_c = ttk.LabelFrame(parent, text="Resultado", padding=8)
        log_c.pack(fill="both", expand=True)
        self.txt_correos = self._make_text(log_c)

    # ── TAB 3: Resoluciones ───────────────────────────────────────────────────

    def _build_tab_resoluciones(self, parent):
        ttk.Label(parent,
                  text="Lee el Excel de salida del motor (con columna OBSERVACION) y genera un Word consolidado.",
                  font=("Segoe UI", 9), foreground="#666666").pack(anchor="w", pady=(0, 6))

        # Entradas
        inp = ttk.LabelFrame(parent, text="Archivos", padding=8)
        inp.pack(fill="x", pady=(0, 6))

        self.res_excel_var  = tk.StringVar(value="")
        self.res_salida_var = tk.StringVar(value=self.cfg.get("ruta_salida_excel", ""))

        for i, (lbl, var, es_archivo) in enumerate([
            ("Excel con OBSERVACION:", self.res_excel_var,  True),
            ("Carpeta salida .docx:",  self.res_salida_var, False),
        ]):
            ttk.Label(inp, text=lbl, width=22, anchor="w").grid(row=i, column=0, sticky="w", pady=2)
            ttk.Entry(inp, textvariable=var, width=68).grid(row=i, column=1, padx=5, sticky="ew")
            cmd = (lambda v=var: self._pick_file(v)) if es_archivo else (lambda v=var: self._pick_dir(v))
            ttk.Button(inp, text="Buscar…", width=8, command=cmd).grid(row=i, column=2)
        inp.columnconfigure(1, weight=1)

        # Botón
        act = ttk.Frame(parent)
        act.pack(fill="x", pady=(0, 6))
        self._res_btn = ttk.Button(act, text="⚙  Generar resoluciones .docx",
                                   command=self._run_resoluciones)
        self._res_btn.pack(side="left", padx=(0, 12))
        ttk.Label(act, text="Detecta automáticamente PC_IE, PC_INFO y NOMENCL según tribunal.",
                  font=("Segoe UI", 9), foreground="#777777").pack(side="left")

        # Log resoluciones
        log_r = ttk.LabelFrame(parent, text="Resultado", padding=8)
        log_r.pack(fill="both", expand=True)
        self.txt_resoluciones = self._make_text(log_r)

    # ── helpers de texto ──────────────────────────────────────────────────────

    def _build_log(self, parent):
        """Bitácora principal (pestaña Motor)."""
        toolbar = ttk.Frame(parent)
        toolbar.pack(fill="x", side="bottom", pady=(4, 0))
        ttk.Button(toolbar, text="🗑  Limpiar",
                   command=self._clear_log, width=10).pack(side="right")

        self.txt = tk.Text(parent, height=16, font=("Consolas", 9),
                           state="disabled", wrap="word",
                           background="#1e1e1e", foreground="#d4d4d4",
                           insertbackground="white", relief="flat")
        self.txt.pack(fill="both", expand=True, side="left")
        sb = ttk.Scrollbar(parent, command=self.txt.yview)
        sb.pack(fill="y", side="right")
        self.txt.config(yscrollcommand=sb.set)
        self.txt.tag_configure(TAG_OK,   foreground="#4ec9b0")
        self.txt.tag_configure(TAG_ERR,  foreground="#f44747", font=("Consolas", 9, "bold"))
        self.txt.tag_configure(TAG_WARN, foreground="#ce9178")
        self.txt.tag_configure(TAG_INFO, foreground="#569cd6")
        self.txt.tag_configure(TAG_DONE, foreground="#b5cea8", font=("Consolas", 9, "bold"))
        self._log("Sistema iniciado. Selecciona un Excel y presiona ▶ para comenzar.", TAG_INFO)

    def _make_text(self, parent):
        """Text widget de solo lectura para logs de correos/resoluciones."""
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True)
        t = tk.Text(frame, height=14, font=("Consolas", 9),
                    state="disabled", wrap="word",
                    background="#1e1e1e", foreground="#d4d4d4",
                    insertbackground="white", relief="flat")
        t.pack(fill="both", expand=True, side="left")
        sb = ttk.Scrollbar(frame, command=t.yview)
        sb.pack(fill="y", side="right")
        t.config(yscrollcommand=sb.set)
        t.tag_configure(TAG_OK,   foreground="#4ec9b0")
        t.tag_configure(TAG_ERR,  foreground="#f44747", font=("Consolas", 9, "bold"))
        t.tag_configure(TAG_WARN, foreground="#ce9178")
        t.tag_configure(TAG_INFO, foreground="#569cd6")
        t.tag_configure(TAG_DONE, foreground="#b5cea8", font=("Consolas", 9, "bold"))
        return t

    def _write(self, widget, msg, tag=None):
        """Escribe en cualquier Text widget con color."""
        t = datetime.now().strftime("%H:%M:%S")
        if tag is None:
            if msg.startswith(("✅","📁")):                    tag = TAG_OK
            elif msg.startswith("❌"):                          tag = TAG_ERR
            elif msg.startswith(("⚠️","⚠")):                  tag = TAG_WARN
            elif msg.startswith(("ℹ️","📋","📄","✓","📂")):    tag = TAG_INFO
        widget.config(state="normal")
        lines = int(widget.index("end-1c").split(".")[0])
        if lines > MAX_LOG_LINES:
            widget.delete("1.0", f"{lines - MAX_LOG_LINES}.0")
        start = widget.index("end")
        widget.insert("end", f"[{t}] {msg}\n")
        if tag:
            widget.tag_add(tag, start, widget.index("end"))
        widget.config(state="disabled")
        widget.see("end")

    def _log(self, msg, tag=None):
        self._write(self.txt, msg, tag)

    def _clear_log(self):
        self.txt.config(state="normal")
        self.txt.delete("1.0", "end")
        self.txt.config(state="disabled")

    # ── queue ──────────────────────────────────────────────────────────────────

    def _process_queue(self):
        try:
            while True:
                kind, data = self.q.get_nowait()
                if kind == "log":      self._log(data)
                elif kind == "progress": self.pb["value"] = data
                elif kind == "status":   self.status.set(data)
                elif kind == "done":     self._finish_motor(data)
                elif kind == "preview_ready": self._on_preview_ready(data)
        except queue.Empty:
            pass
        self.after(100, self._process_queue)

    # ── pickers ───────────────────────────────────────────────────────────────

    def _pick_in(self):
        f = filedialog.askopenfilename(title="Excel de entrada",
            filetypes=[("Excel","*.xlsx *.xls *.xlsm"),("Todos","*.*")])
        if f: self.in_var.set(f)

    def _pick_out(self):
        d = filedialog.askdirectory(title="Carpeta de salida")
        if d: self.out_var.set(d)

    def _pick_file(self, var):
        f = filedialog.askopenfilename(title="Seleccionar Excel",
            filetypes=[("Excel","*.xlsx *.xls *.xlsm"),("Todos","*.*")])
        if f: var.set(f)

    def _pick_dir(self, var):
        d = filedialog.askdirectory(title="Seleccionar carpeta")
        if d: var.set(d)

    def _save_cfg(self):
        self.cfg["ruta_entrada_excel"] = self.in_var.get().strip()
        self.cfg["ruta_salida_excel"]  = self.out_var.get().strip()
        _guardar_json(CONFIG_FILE, self.cfg)
        self._log("✅ Configuración guardada")

    def _open_out(self):
        p = self.out_var.get().strip()
        if p and Path(p).exists():
            os.startfile(p)
        else:
            messagebox.showwarning("Aviso", "La carpeta de salida no existe.")

    # ── TAB 1: Motor jobs ─────────────────────────────────────────────────────

    def _set_motor_btns(self, state):
        for b in self._motor_btns: b["state"] = state

    def _start_motor(self, modo):
        if self.running: return
        if not self.in_var.get().strip():
            messagebox.showerror("Falta configuración", "Selecciona el archivo Excel de entrada.")
            return
        self.running = True
        self.pb["value"] = 0
        self._set_motor_btns("disabled")
        self.status.set(f"Procesando {modo}…")
        threading.Thread(target=self._worker_motor, args=(modo,), daemon=True).start()

    def _worker_motor(self, modo):
        try:
            self.q.put(("progress", 5))
            path = self.in_var.get().strip()
            if modo == "CUMPLIMIENTO":
                self.q.put(("log", f"📂 Cargando archivo para {modo}…"))
                self.q.put(("progress", 20))
                procesar(path, modo, self.cfg, self.q)
            else:
                self.q.put(("log", f"📂 Cargando Excel para {modo}…"))
                engine = "xlrd" if path.endswith(".xls") else "openpyxl"
                df = pd.read_excel(path, engine=engine)
                df.columns = [str(x).strip() for x in df.columns]
                df = df.dropna(how="all").fillna("")
                self.q.put(("log", f"✓ {len(df)} filas cargadas"))
                self.q.put(("progress", 20))
                procesar(df, modo, self.cfg, self.q)
        except Exception as e:
            self.q.put(("done", f"❌ Error inesperado: {e}"))

    def _finish_motor(self, msg):
        self.running = False
        self._set_motor_btns("normal")
        self.pb["value"] = 100
        self.status.set("Listo")
        tag = TAG_ERR if msg.startswith("❌") else TAG_DONE
        self._log(msg.split("\n")[0], tag)
        messagebox.showinfo("Proceso finalizado", msg)

    # ── TAB 1: Vista previa (S3 — v8.14) ─────────────────────────────────────

    def _set_preview_btns(self, state):
        for b in self._preview_btns: b["state"] = state

    def _start_preview(self, modo):
        """Calcula observaciones EN MEMORIA (sin escribir archivo) y muestra
        una ventana con las primeras 30 filas antes de exportar."""
        if self.running:
            return
        path = self.in_var.get().strip()
        if not path:
            messagebox.showerror("Falta configuración", "Selecciona el archivo Excel de entrada.")
            return
        self.status.set(f"Generando vista previa {modo}…")
        self._set_preview_btns("disabled")
        threading.Thread(target=self._worker_preview, args=(modo, path), daemon=True).start()

    def _worker_preview(self, modo, path):
        try:
            df, err = calcular_preview(path, modo)
        except Exception as e:  # noqa: BLE001
            df, err = None, str(e)
        self.q.put(("preview_ready", (df, modo, err)))

    def _on_preview_ready(self, data):
        df, modo, err = data
        self.status.set("Listo")
        self._set_preview_btns("normal")
        if err:
            messagebox.showerror("Vista previa", f"No se pudo generar la vista previa:\n{err}")
            return
        self._mostrar_ventana_preview(df, modo)

    def _mostrar_ventana_preview(self, df, modo):
        """Ventana Treeview de solo lectura: NNA | Observación.
        No escribe ningún archivo — solo exporta si el usuario lo confirma."""
        col_nombre = mapear_columnas(df, modo).get("nombre")
        total = len(df)
        muestra = min(30, total)

        win = tk.Toplevel(self)
        win.title(f"Vista previa — {modo}")
        win.geometry("920x540")
        win.minsize(640, 360)

        ttk.Label(win, text=f"Vista previa {modo} — mostrando {muestra} de {total} filas",
                  font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=10, pady=(10, 4))
        ttk.Label(win, text="Nada se ha guardado todavía. Revisa y presiona Exportar si conforme.",
                  font=("Segoe UI", 9), foreground="#666666").pack(anchor="w", padx=10, pady=(0, 6))

        frame = ttk.Frame(win)
        frame.pack(fill="both", expand=True, padx=10, pady=(0, 6))

        tree = ttk.Treeview(frame, columns=("nna", "obs"), show="headings")
        tree.heading("nna", text="NNA")
        tree.heading("obs", text="Observación")
        tree.column("nna", width=200, anchor="w", stretch=False)
        tree.column("obs", width=680, anchor="w")
        tree.pack(fill="both", expand=True, side="left")
        sb = ttk.Scrollbar(frame, command=tree.yview)
        sb.pack(fill="y", side="right")
        tree.config(yscrollcommand=sb.set)

        for _, row in df.head(30).iterrows():
            nombre = str(row.get(col_nombre, "")).strip() if col_nombre else "(sin nombre)"
            obs = str(row.get("OBSERVACION", ""))
            tree.insert("", "end", values=(nombre, obs))

        btns = ttk.Frame(win)
        btns.pack(fill="x", padx=10, pady=(0, 10))
        ttk.Button(btns, text="Cerrar", command=win.destroy).pack(side="right", padx=4)
        ttk.Button(btns, text="💾  Exportar…",
                   command=lambda: self._exportar_desde_preview(modo, win)).pack(side="right", padx=4)

    def _exportar_desde_preview(self, modo, win):
        """El usuario aprobó la vista previa: cierra la ventana y dispara
        el pipeline REAL de guardado (idéntico al botón ▶ normal)."""
        win.destroy()
        self._start_motor(modo)

    # ── TAB 1: Confirmación de textos (S4 — v8.14) ───────────────────────────

    def _actualizar_badge_pendientes(self):
        try:
            n = contar_pendientes()
        except Exception:
            n = "?"
        self._pendientes_var.set(f"📋  Textos pendientes de confirmación: {n}")

    def _exportar_textos_pendientes(self):
        destino = filedialog.asksaveasfilename(
            title="Exportar textos pendientes",
            initialfile="TEXTOS_PENDIENTES.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")])
        if not destino:
            return
        try:
            n = exportar_pendientes(destino)
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Error al exportar", str(e))
            return
        self._log(f"✅ {n} textos pendientes exportados: {destino}", TAG_OK)
        messagebox.showinfo("Exportado", f"{n} textos pendientes exportados.\n\n{destino}")

    def _importar_textos_confirmados(self):
        origen = filedialog.askopenfilename(
            title="Importar confirmaciones",
            filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")])
        if not origen:
            return
        if not messagebox.askyesno(
                "Confirmar importación",
                "Esto va a modificar motor/textos_observaciones.json.\n"
                "Se creará un respaldo automático con timestamp antes de "
                "escribir.\n\n¿Continuar?"):
            return
        try:
            resumen = importar_confirmaciones(origen)
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Error al importar", str(e))
            return

        self._actualizar_badge_pendientes()
        if resumen["errores"]:
            self._log(f"⚠️ {len(resumen['errores'])} filas con error al importar",
                       TAG_WARN)
        self._log(
            f"✅ Confirmaciones importadas — "
            f"{len(resumen['confirmados'])} confirmados, "
            f"{len(resumen['actualizados'])} textos actualizados. "
            f"Backup: {resumen['backup']}", TAG_OK)
        messagebox.showinfo(
            "Importación completada",
            f"Actualizados: {len(resumen['actualizados'])}\n"
            f"Confirmados: {len(resumen['confirmados'])}\n"
            f"Sin cambios: {len(resumen['sin_cambios'])}\n"
            f"Errores: {len(resumen['errores'])}\n\n"
            f"Backup del JSON previo:\n{resumen['backup']}")

    # ── TAB 2: Correos jobs ───────────────────────────────────────────────────

    def _set_correo_btns(self, state):
        for b in self._correo_btns: b["state"] = state

    def _run_correos_informes(self):
        path = self.correo_in_var.get().strip()
        if not path:
            messagebox.showerror("Falta archivo", "Selecciona el Excel de entrada.")
            return
        self._set_correo_btns("disabled")
        self.status.set("Generando borradores de correos (informes)…")
        threading.Thread(target=self._worker_correos,
                         args=(path, "informes"), daemon=True).start()

    def _run_correos_espera(self):
        path = self.correo_in_var.get().strip()
        if not path:
            messagebox.showerror("Falta archivo", "Selecciona el Excel de entrada.")
            return
        self._set_correo_btns("disabled")
        self.status.set("Generando borradores de correos (espera)…")
        threading.Thread(target=self._worker_correos,
                         args=(path, "espera"), daemon=True).start()

    def _worker_correos(self, path, tipo):
        txt = self.txt_correos
        try:
            from comunicaciones.generador_correos import GeneradorCorreos
            self._write(txt, f"📂 Cargando {path}…", TAG_INFO)

            engine = "xlrd" if path.endswith(".xls") else "openpyxl"
            df = pd.read_excel(path, engine=engine)
            df.columns = [str(x).strip() for x in df.columns]
            df = df.dropna(how="all").fillna("")
            self._write(txt, f"✓ {len(df)} filas cargadas", TAG_INFO)

            catastro = str(_ROOT / "comunicaciones" / "catastro_programas.json")
            gen = GeneradorCorreos(self.cfg, catastro, self.contactos)

            if tipo == "informes":
                resultado = gen.procesar(df)
            else:
                resultado = gen.procesar_espera(df)

            creados = resultado.get("borradores_creados", 0)
            errores = resultado.get("errores", [])
            sin_ctc = resultado.get("grupos_sin_contacto", [])

            self._write(txt, f"✅ Borradores creados: {creados}", TAG_OK)
            for e in errores:
                self._write(txt, f"⚠️  {e}", TAG_WARN)
            for s in sin_ctc:
                self._write(txt, f"⚠️  Sin contacto: {s}", TAG_WARN)

            resumen = f"Correos ({tipo}) completado.\n{creados} borradores en Outlook Drafts."
            if errores:
                resumen += f"\n{len(errores)} advertencias."
            self.status.set("Listo")
            messagebox.showinfo("Correos generados", resumen)

        except ImportError:
            self._write(txt, "❌ pywin32 no instalado. Ejecuta: pip install pywin32", TAG_ERR)
            messagebox.showerror("Dependencia faltante",
                "Se requiere pywin32 y Outlook instalado.\npip install pywin32")
        except Exception as e:
            self._write(txt, f"❌ Error: {e}", TAG_ERR)
            messagebox.showerror("Error", str(e))
        finally:
            self._set_correo_btns("normal")
            self.status.set("Listo")

    # ── TAB 3: Resoluciones jobs ──────────────────────────────────────────────

    def _run_resoluciones(self):
        excel  = self.res_excel_var.get().strip()
        salida = self.res_salida_var.get().strip()
        if not excel:
            messagebox.showerror("Falta archivo", "Selecciona el Excel con columna OBSERVACION.")
            return
        if not salida:
            messagebox.showerror("Falta carpeta", "Selecciona la carpeta de salida.")
            return
        self._res_btn["state"] = "disabled"
        self.status.set("Generando resoluciones .docx…")
        threading.Thread(target=self._worker_resoluciones,
                         args=(excel, salida), daemon=True).start()

    def _worker_resoluciones(self, excel, salida):
        txt = self.txt_resoluciones
        try:
            from resoluciones.generador_resoluciones import generar_resoluciones
            self._write(txt, f"📂 Procesando {Path(excel).name}…", TAG_INFO)

            resultado = generar_resoluciones(excel, salida)

            arch     = resultado.get("archivo_generado")
            total    = resultado.get("total_resoluciones", 0)
            faltantes = resultado.get("faltantes", [])
            errores  = resultado.get("errores", [])

            if arch:
                self._write(txt, f"✅ Word generado: {Path(arch).name}", TAG_OK)
                self._write(txt, f"📁 {arch}", TAG_OK)
            self._write(txt, f"✓ Resoluciones generadas: {total}", TAG_INFO)

            for f in faltantes:
                self._write(txt,
                    f"⚠️  Sin plantilla — RIT {f.get('rit','?')} "
                    f"({f.get('tipo','?')} {f.get('tribunal','?')})", TAG_WARN)
            for e in errores:
                self._write(txt, f"⚠️  {e}", TAG_WARN)

            resumen = f"Resoluciones completadas.\n{total} generadas."
            if faltantes:
                resumen += f"\n{len(faltantes)} sin plantilla (ver bitácora)."
            messagebox.showinfo("Resoluciones generadas", resumen)

        except Exception as e:
            self._write(txt, f"❌ Error: {e}", TAG_ERR)
            messagebox.showerror("Error", str(e))
        finally:
            self._res_btn["state"] = "normal"
            self.status.set("Listo")


if __name__ == "__main__":
    RUSApp().mainloop()
