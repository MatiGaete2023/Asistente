# -*- coding: utf-8 -*-
from datetime import datetime
from .composicion import componer, prefijo, fecha_valida, Incidencias
from .utilidades import fecha_es, titulo_programa, es_derivacion_sin_seg, es_dce
from .textos import render
from .reglas_espera import _audiencia

def generar_observacion_informes(row, tribunal, cols, incidencias=None, fila_excel=None):
    if incidencias is None:
        incidencias = Incidencias()
    programa=str(row.get(cols.get('programa'), '')).strip(); nombre=str(row.get(cols.get('nombre'), '')).strip()
    pfx=prefijo(nombre, programa); prog=titulo_programa(programa)
    if es_derivacion_sin_seg(programa): return componer(pfx,[render('COMUN','NO_SEGUIMIENTO', PROGRAMA=prog)])
    fv=fecha_valida(row.get(cols.get('vencimiento'))) if cols.get('vencimiento') else None
    if not fv:
        incidencias.agregar(fila_excel, row.get(cols.get('rit'), '') if cols.get('rit') else '', 'I-01/I-02', 'sin fecha de vencimiento (G-05)'); return ''
    vd=fv.date() if hasattr(fv,'date') else fv; dias=(vd-datetime.now().date()).days
    if dias < 0: key='I01_VENCIDO_DCE' if es_dce(programa) else 'I01_VENCIDO_GENERAL'
    elif 0 <= dias <= 30: key='I02_POR_VENCER_DCE' if es_dce(programa) else 'I02_POR_VENCER_GENERAL'
    else: return ''
    fr=[render('INFORMES', key, PROGRAMA=prog, FECHA_VENCIMIENTO=fecha_es(fv))]
    audiencia = _audiencia(row, cols)
    if audiencia:
        fr.append(audiencia)
    return componer(pfx, fr)
