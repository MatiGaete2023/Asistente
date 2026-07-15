from motor.composicion import componer, prefijo
from motor.utilidades import contiene_token, es_derivacion_sin_seg


def test_componer_fragmentos_y_puntos():
    assert componer('Ana PIE: ', []) == ''
    assert componer('Ana PIE: ', ['Uno']) == 'Ana PIE: Uno.'
    assert componer('', ['Uno.', 'Dos', 'Tres..']) == 'Uno. Dos. Tres.'
    assert '..' not in componer('', ['Uno..', 'Dos'])


def test_prefijo():
    assert prefijo('Camila Pérez', 'pie norte') == 'Camila PIE: '
    assert prefijo('', 'rta sur') == 'RTA: '


def test_contiene_token():
    assert contiene_token('Residencia San Rafael', 'fae') is False
    assert contiene_token('FAE Familia Sur', 'fae') is True


def test_derivacion_sin_seguimiento_anclada():
    assert es_derivacion_sin_seg('OPD Laja') is True
    assert es_derivacion_sin_seg('PRODAM X') is False
    assert es_derivacion_sin_seg('Chile Crece Contigo B') is True
    assert es_derivacion_sin_seg('DCE Tomé') is False
    assert es_derivacion_sin_seg('Red Salud Bío Bío') is True
