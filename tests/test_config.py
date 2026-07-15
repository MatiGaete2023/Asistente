import json

import pytest

# gui.app importa tkinter a nivel de módulo; en entornos sin Tk (CI headless,
# sandboxes) este archivo se salta completo en vez de romper la colección.
pytest.importorskip("tkinter")

from gui.app import DEFAULT_CONFIG, _cargar_json, _guardar_json  # noqa: E402


def test_config_atomica_filtra_claves_no_permitidas(tmp_path):
    ruta = tmp_path / "config.json"
    data = {**DEFAULT_CONFIG, "cc_fijo": "atacante@example.invalid"}
    _guardar_json(ruta, data)

    cargada = _cargar_json(ruta, DEFAULT_CONFIG)
    assert "cc_fijo" not in cargada
    assert json.loads(ruta.read_text(encoding="utf-8"))["cc_fijo"]


def test_config_rechaza_tipos_invalidos(tmp_path):
    ruta = tmp_path / "config.json"
    ruta.write_text(json.dumps({"dias_retencion_logs": "noventa"}), encoding="utf-8")
    with pytest.raises(ValueError, match="entero"):
        _cargar_json(ruta, DEFAULT_CONFIG)


def test_cli_help_basico():
    from cli import construir_parser

    parser = construir_parser()
    ayuda = parser.format_help()
    assert "procesar" in ayuda
    assert "preview" in ayuda
    assert "correos" in ayuda
    assert "resoluciones" in ayuda
