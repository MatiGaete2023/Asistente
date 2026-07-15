from datetime import datetime, timedelta

import pandas as pd

from validador.precheck import validar_excel
from validador.reglas_validacion import a4_fechas_futuras


def test_columnas_requeridas_dependen_del_modo():
    base = pd.DataFrame([{"RIT": "R-1", "DERIVACION": "PIE", "TRIBUNAL": "LAJA", "NOMBRE": "Caso"}])
    for modo, esperado in [
        ("ESPERA", "TIEMPO_ESPERA"),
        ("CUMPLIMIENTO", "DIAS_CUMPLIMIENTO"),
        ("INFORMES", "FECHA_VENCIMIENTO"),
    ]:
        resultado = validar_excel(base, modo)
        assert not resultado["puede_procesar"]
        detalle = " ".join(resultado["anomalias_bloqueantes"][0]["detalle"])
        assert esperado in detalle


def test_a4_no_duplica_la_misma_columna_por_alias():
    futuro = datetime.now() + timedelta(days=5)
    df = pd.DataFrame([{
        "RIT": "R-1",
        "FEC.INGRESO EFECTIVO": futuro,
        "FEC. NACIMIENTO": futuro,
    }])
    resultado = a4_fechas_futuras(df, "CUMPLIMIENTO")
    assert resultado["count"] == 2
    assert len(resultado["detalle"]) == 2


def test_tribunal_desconocido_bloquea():
    df = pd.DataFrame([{
        "RIT": "R-1", "DERIVACION": "PIE", "TRIBUNAL": "OTRO", "NOMBRE": "Caso",
        "FECHA VENCIMIENTO": datetime.now(),
    }])
    resultado = validar_excel(df, "INFORMES")
    ids = {item["id"] for item in resultado["anomalias_bloqueantes"]}
    assert not resultado["puede_procesar"]
    assert "A8" in ids


def test_valor_operativo_invalido_bloquea():
    df = pd.DataFrame([{
        "RIT": "R-1", "DERIVACION": "PIE", "TRIBUNAL": "LAJA", "NOMBRE": "Caso",
        "FECHA VENCIMIENTO": "no-es-fecha",
    }])
    resultado = validar_excel(df, "INFORMES")
    ids = {item["id"] for item in resultado["anomalias_bloqueantes"]}
    assert "A10" in ids
