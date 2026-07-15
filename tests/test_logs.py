from datetime import datetime

import pytest

from logs.log_manager import get_logger


def test_log_diario_y_rotacion(tmp_path):
    antiguo = tmp_path / "csmp_20000101.log"
    antiguo.write_text("antiguo\n", encoding="utf-8")

    logger = get_logger(str(tmp_path), nombre="prueba", dias_retencion=30)
    logger.info("evento=prueba filas=3")
    for handler in logger.handlers:
        handler.flush()

    actual = tmp_path / f"csmp_{datetime.now():%Y%m%d}.log"
    assert actual.exists()
    assert "evento=prueba filas=3" in actual.read_text(encoding="utf-8")
    assert not antiguo.exists()

    for handler in list(logger.handlers):
        handler.close()
        logger.removeHandler(handler)


def test_retencion_fuera_de_rango_se_rechaza(tmp_path):
    with pytest.raises(ValueError, match="dias_retencion"):
        get_logger(str(tmp_path), dias_retencion=0)
