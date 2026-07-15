# -*- coding: utf-8 -*-
from datetime import datetime
from .composicion import componer, prefijo, fecha_valida, Incidencias
from .utilidades import (fecha_es, limpiar_nombre, normalizar, es_derivacion_sin_seg,
    get_int, calcular_edad_exacta, dias_para_mayoria, fecha_mayoria,
    titulo_programa, contiene_token)
from .textos import render
from .reglas_espera import _complementarias, _curador_oido, _audiencia

_RES=('rta','rtt','res','rfa','rva')
def _d(v): return v.date() if hasattr(v,'date') else v
def _pn(nombre):
    n=limpiar_nombre(nombre); return n.split()[0] if n else ''
def _rit(row, cols): return row.get(cols.get('rit'), '') if cols.get('rit') else ''

def generar_observacion_cumplimiento(row, tribunal, cols, fecha_hoja2=None, incidencias=None, fila_excel=None) -> str:
    if incidencias is None:
        incidencias = Incidencias()
    hoy=datetime.now().date()
    programa=str(row.get(cols.get('programa'), '')).strip(); nombre=str(row.get(cols.get('nombre'), '')).strip()
    pfx=prefijo(nombre, programa); pn=_pn(nombre); prog=titulo_programa(programa)
    if es_derivacion_sin_seg(programa): return componer(pfx,[render('COMUN','NO_SEGUIMIENTO', PROGRAMA=prog)])
    fn=fecha_valida(row.get(cols.get('nacimiento'))) if cols.get('nacimiento') else None
    if fn and (calcular_edad_exacta(fn) or 0) >= 18:
        return componer(pfx,[render('COMUN','MAYORIA_EDAD', PNOMBRE=pn, FECHA_MAYORIA=fecha_es(fecha_mayoria(fn)))] + _complementarias(row, cols))
    frags=[]; principal=False
    dm=dias_para_mayoria(fn) if fn else None
    if dm is not None and 1 <= dm <= 60:
        frags.append(render('COMUN','PROXIMA_MAYORIA', PNOMBRE=pn, FECHA_MAYORIA=fecha_es(fecha_mayoria(fn))))
    dc=get_int(row.get(cols.get('dias_cumpl'))) if cols.get('dias_cumpl') else None
    if cols.get('dias_cumpl') and dc is None:
        incidencias.agregar(fila_excel, _rit(row, cols), 'C-03',
                            'días de cumplimiento ausentes o inválidos (G-05)')
    fing=fecha_valida(row.get(cols.get('ingreso'))) if cols.get('ingreso') else None
    if dc is not None and 0 <= dc <= 30:
        if fing: frags.append(render('CUMPLIMIENTO','C03_INGRESO_RECIENTE', PROGRAMA=prog, FECHA_INGRESO=fecha_es(fing))); principal=True
        else: incidencias.agregar(fila_excel, _rit(row, cols), 'C-03', 'sin fecha de ingreso — regla omitida (G-05)')
    dpe=get_int(row.get(cols.get('dias_egresar'))) if cols.get('dias_egresar') else None
    if cols.get('dias_egresar') and dpe is None:
        incidencias.agregar(fila_excel, _rit(row, cols), 'C-04/C-05',
                            'días para egresar ausentes o inválidos (G-05)')
    fegr=fecha_valida(row.get(cols.get('egreso_proy'))) if cols.get('egreso_proy') else None
    vencida=bool(fegr and ((dpe is not None and dpe < 0) or (dc is not None and dc < 0)))
    if vencida: frags.append(render('CUMPLIMIENTO','C04_VENCIDA', FECHA_EGRESO_PROYECTADO=fecha_es(fegr))); principal=True
    elif ((dpe is not None and dpe < 0) or (dc is not None and dc < 0)) and not fegr:
        incidencias.agregar(fila_excel, _rit(row, cols), 'C-04', 'sin egreso proyectado (G-05)')
    c05=bool(fegr and dpe is not None and 0 <= dpe <= 45 and not vencida)
    if c05:
        frags.append(render('CUMPLIMIENTO','C05_VENCE_HOY' if dpe == 0 else 'C05_POR_VENCER', FECHA_EGRESO_PROYECTADO=fecha_es(fegr))); principal=True
    if fecha_hoja2 and not vencida and not c05:
        frags.append(render('CUMPLIMIENTO','C10_HOJA2', PROGRAMA=prog, FECHA_VENCIMIENTO=fecha_es(fecha_hoja2))); principal=True
    # Orden §9.2: curador (6) y oído (7) van ANTES de las fichas (8-9);
    # la audiencia (10) se agrega al final, después de C-07/C-08.
    frags += _curador_oido(row, cols)
    pnrm=normalizar(programa)
    fi=fecha_valida(row.get(cols.get('ficha_ind'))) if cols.get('ficha_ind') else None
    if pnrm.startswith(_RES) and cols.get('ficha_ind'):
        if not fi: frags.append(render('CUMPLIMIENTO','C07_SIN_FICHA'))
        else:
            dd=(hoy-_d(fi)).days
            if dd > 180: frags.append(render('CUMPLIMIENTO','C07_FICHA_ANTIGUA', FECHA_FICHA_INDIVIDUAL=fecha_es(fi)))
            elif 0 <= dd <= 30: frags.append(render('CUMPLIMIENTO','C07_FICHA_RECIENTE', FECHA_FICHA_INDIVIDUAL=fecha_es(fi)))
    ffae=fecha_valida(row.get(cols.get('ficha_fae'))) if cols.get('ficha_fae') else None
    if contiene_token(programa,'fae','fas') and fing and (hoy-_d(fing)).days > 120 and not ffae:
        frags.append(render('CUMPLIMIENTO','C08_FICHA_FAE', PNOMBRE=pn))
    aud=_audiencia(row, cols)
    if aud: frags.append(aud)
    if not principal:
        return componer(pfx, [render('CUMPLIMIENTO','C09_SIN_OBSERVACIONES')] if not frags else [render('CUMPLIMIENTO','C09_BASE_BREVE')] + frags)
    return componer(pfx, frags)
