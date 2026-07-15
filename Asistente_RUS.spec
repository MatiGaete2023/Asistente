# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path

ROOT = Path(SPECPATH)

datas = [
    (str(ROOT / "motor" / "textos_observaciones.json"), "motor"),
    (str(ROOT / "comunicaciones" / "catastro_programas.json"), "comunicaciones"),
    (str(ROOT / "comunicaciones" / "aliases_programas.json"), "comunicaciones"),
]
datas.extend(
    (str(path), "comunicaciones/plantillas")
    for path in sorted((ROOT / "comunicaciones" / "plantillas").glob("*.html"))
)

a = Analysis(
    [str(ROOT / "main.py")],
    pathex=[str(ROOT)],
    binaries=[],
    datas=datas,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=1,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="Asistente_RUS",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
