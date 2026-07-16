# -*- coding: utf-8 -*-
"""Congelamiento global del reloj para la suite (fase M5 del plan v9.1).

Se activa en pytest_configure — ANTES de que pytest importe los módulos de
tests — porque varios de ellos calculan fechas relativas a nivel de módulo
(`HOY = datetime.now()`). Congelar recién en una fixture dejaría esos valores
con el reloj real y los motores con el congelado, rompiendo los bordes al
cruzar medianoche o al correr en otra fecha.

Nota: `tests/runner.py` ejecuta el motor en un subprocess aislado que NO
queda congelado; los tests que lo usan (regresión/smoke) solo verifican
invariantes estructurales, no fechas — sigue siendo válido.
"""
import time_machine

_FECHA_SUITE = "2026-07-15 10:00:00"
_traveller = None


def pytest_configure(config):
    global _traveller
    _traveller = time_machine.travel(_FECHA_SUITE, tick=False)
    _traveller.start()


def pytest_unconfigure(config):
    global _traveller
    if _traveller is not None:
        _traveller.stop()
        _traveller = None
