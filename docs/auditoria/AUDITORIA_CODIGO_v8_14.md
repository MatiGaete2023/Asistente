# AUDITORÍA DE CÓDIGO — CSMP Assistant v8.14

**Fecha:** 15 de julio de 2026
**Base auditada:** `CSMP_Assistant_v8_14_COMPLETO.zip` (contenido copiado tal cual a la raíz de este repositorio — commit inicial).
**Estado del baseline verificado:** `pytest tests/ -q` → **73 passed, 3 skipped** (coincide con lo declarado en README y PROMPT_CONTINUACION).
**Alcance:** revisión senior de todo el código fuente, contrastada contra `docs/especificaciones/Catalogo_Reglas_CSMP_v2_20260714.md` (el "catálogo") y `docs/especificaciones/Mejoras_Correos_y_Otros_Modulos_20260714.md` (las "mejoras").

Este documento tiene tres secciones:

1. **Hallazgos propios** (H-01 … H-17): defectos y mejoras detectados en esta auditoría que **no** están en los dos documentos de especificación.
2. **Verificación de los hallazgos de los documentos MD**: confirmación contra el código real, con archivo:línea.
3. **Notas de riesgo para el implementador.**

Cada hallazgo indica la fase del plan (`docs/plan/PLAN_IMPLEMENTACION_v9.md`) que lo resuelve.

---

## 1. Hallazgos propios

### H-01 · CRÍTICO — Tribunal por defecto "LAJA" cuando no se reconoce

`motor/procesador.py:141` y `motor/procesador.py:182`:

```python
trib = detectar_tribunal(str(row.get(col_trib, ""))) or "LAJA"
```

Si la celda TRIBUNAL trae un valor no reconocible, la fila se procesa **como si fuera de Laja**. Con las reglas v9 esto es grave: una fila de Tomé mal escrita con 40 días de espera recibiría "proyecto de resolución" (rama Laja/Mulchén) cuando le corresponde solo correo (rama Tomé 30–59). El catálogo G-06 lo prohíbe expresamente: *"no hay valor predeterminado"*; sin tribunal reconocible se omiten las reglas tribunal-dependientes y la fila se marca en validación (A8).
**Resuelve:** Fase F6.

### H-02 · CRÍTICO — `validador/` y `logs/` son módulos huérfanos (nunca se ejecutan)

`grep -rn "validador\|precheck\|validar_excel\|log_manager"` fuera de sus propias carpetas: **cero referencias**. Ni `gui/app.py` ni `motor/procesador.py` llaman a `validador.precheck.validar_excel()` ni a `logs.log_manager.get_logger()`. Consecuencias:

- El catálogo asume que la regla A8 (tribunal no reconocido) "ya existe" y marca filas — existe pero **no corre nunca**.
- Las reglas bloqueantes A1–A3 (columnas mínimas, RIT duplicados, derivación vacía) no protegen nada.
- La config `ruta_logs` / `dias_retencion_logs` de `gui/app.py:53-58` promete una bitácora en disco que no se escribe.

**Resuelve:** Fase F9 (cablear validador), F10 (logging opcional).

### H-03 · CRÍTICO — Falso positivo en detección FAE por subcadena

`motor/reglas_cumplimiento.py:124`:

```python
if "fae" in programa_norm or "fas" in programa_norm:
```

`normalizar("Residencia San Rafael")` → `"residencia san rafael"`, que **contiene "fae"** (ra-**fae**-l). Cualquier programa cuyo nombre contenga "Rafael" (u otra palabra con esas letras) entra a la regla de ficha FAE. El catálogo C-08 dice "Programa **contiene FAE o FAS**" — debe interpretarse como token/palabra completa, no subcadena.
**Resuelve:** Fase F1 (helper `contiene_token()`), F4.

### H-04 · ALTO — `fecha_es(None)` inventa la fecha de hoy en observaciones

`motor/utilidades.py:38-47`: `fecha_es()` sin argumento (o con `None`) retorna la fecha **actual**. En `reglas_cumplimiento.py:87` (R4) y `:92` (R5), si `FEC.EGRESO PROYECTADO` está vacía se renderiza igual: *"La medida se visualiza vencida en RUS desde el 15 de julio de 2026"* — con una fecha **inventada** (hoy). Viola G-05 ("nunca fechas automáticas no justificadas"; la regla incompleta se omite). El catálogo v2 además exige explícitamente fecha válida como condición de C-04 y C-05.
**Resuelve:** Fases F1 (composer valida fechas antes de render) y F4.

### H-05 · ALTO — Violación de thread-safety de Tkinter en pestañas Correos y Resoluciones

`gui/app.py:644-689` (`_worker_correos`) y `gui/app.py:707-742` (`_worker_resoluciones`) corren en `threading.Thread` y desde ahí llaman **directamente** a `self._write(...)` (inserta en `tk.Text`), `self.status.set(...)` y `messagebox.*`. Tkinter no es thread-safe: esto produce crashes o congelamientos intermitentes difíciles de reproducir. La pestaña Motor lo hace bien (todo pasa por `self.q` + `_process_queue`); las otras dos no.
**Resuelve:** Fase F10.

### H-06 · ALTO — La próxima audiencia (T-03) se agrega incluso cuando opera el corte E-01

`motor/reglas_espera.py:44`, `motor/reglas_cumplimiento.py:60` y `motor/procesador.py:86` (`_obs_cruce`): el sufijo de audiencia `aud` se concatena también a la observación de derivación sin seguimiento. El catálogo T-03: *"no se aplica si opera la regla de no seguimiento"*.
**Resuelve:** Fases F3/F4 (motor de reglas nuevo).

### H-07 · ALTO — El cruce Hoja2 REEMPLAZA la observación completa en vez de componerse

`motor/procesador.py:183-189`: cuando hay match en Hoja2 (y no hay vencida/por vencer), `obs = _obs_cruce(...)` **descarta** todo lo ya calculado por `generar_observacion_cumplimiento` — se pierden próxima mayoría (C-02), ingreso reciente (C-03), curador (T-01), oído (T-02) y fichas (C-07/C-08) de esa fila. El catálogo §9.2 ubica C-10 como paso 5 de una composición acumulativa, con las complementarias después (pasos 6–10).
**Resuelve:** Fase F4 (el cruce pasa a evaluarse dentro del motor de reglas de cumplimiento).

### H-08 · MEDIO — E-01 (derivación sin seguimiento) no se aplica en INFORMES

`motor/reglas_informes.py`: no consulta `es_derivacion_sin_seg()` en ninguna parte. El catálogo E-01: *"Rige también en Cumplimiento e Informes"*, y §6: *"I-01 e I-02 sólo se evalúan si la derivación no quedó excluida por E-01"*. Hoy una fila OPD/DAM en INFORMES genera observación de correo al programa.
**Resuelve:** Fase F5.

### H-09 · MEDIO — Mayoría de edad corta curador/oído/audiencia (return temprano)

`motor/reglas_espera.py:47-55` y `motor/reglas_cumplimiento.py:63-71`: al detectar edad ≥ 18 se hace `return` inmediato conservando solo audiencia. El catálogo E-02/C-01: la mayoría de edad es observación principal que **permite acumular curador, oído y próxima audiencia**.
**Resuelve:** Fases F3/F4.

### H-10 · MEDIO — `requirements.txt` incompleto y sin marcadores de plataforma

- Falta `xlrd` (el código lo usa como engine para `.xls`: `procesador.py:226`, `gui/app.py:459,650`; el README sí lo menciona).
- Falta `pytest` como dependencia de desarrollo.
- `pywin32>=305` sin `; sys_platform == 'win32'` — rompe `pip install -r requirements.txt` en Linux/CI.

**Resuelve:** Fase F11.

### H-11 · MEDIO — Duplicación de la carga de CUMPLIMIENTO entre `procesar()` y `calcular_preview()`

`motor/procesador.py:233-264` vs `motor/procesador.py:295-319`: ~40 líneas duplicadas (detección de hoja, lectura de Hoja2, limpieza). Cualquier corrección en una copia deriva de la otra — exactamente la clase de bug que la vista previa S3 pretendía evitar.
**Resuelve:** Fase F6 (extraer `_cargar_libro_cumplimiento()` compartido).

### H-12 · BAJO — Versionado inconsistente y disperso

`VERSION = "v8.14.0"` vive en `gui/app.py:49`; los docstrings de módulos dicen v8.0/v8.1/v8.6/v8.8/v8.13; `main.py` dice v8.0. Centralizar (`motor/version.py` o constante en `__init__`) y subir a v9.0.0 al cerrar el plan.
**Resuelve:** Fase F10.

### H-13 · BAJO — Ineficiencias en el validador (relevantes al cablearlo)

`validador/reglas_validacion.py:165-179` (A4): `_col_match(df, ["RIT", "N° RIT"])` se llama **dentro** del loop por celda — O(filas × columnas × aliases). A4–A8 usan `iterrows`/loops por fila donde bastan operaciones vectorizadas de pandas. Hoy es invisible porque el módulo no corre (H-02); al cablearlo en F9, corregir.
**Resuelve:** Fase F9.

### H-14 · BAJO — La vista previa no bloquea el estado `running`

`gui/app.py:483-494`: `_start_preview` no marca `self.running`, así que se puede lanzar una vista previa y un procesamiento real simultáneos sobre el mismo archivo. Riesgo bajo (la preview no escribe), pero desordena la barra de estado.
**Resuelve:** Fase F10.

### H-15 · BAJO — INFORMES retorna tipos inconsistentes ("" vs None)

`motor/reglas_informes.py:21` retorna `""` sin fecha de vencimiento y `:38` retorna `None` para fechas lejanas. En el Excel ambos terminan como celda vacía, pero el contrato de la función es ambiguo y complica tests. Unificar a `""`.
**Resuelve:** Fase F5.

### H-16 · INFO — Sin `.gitignore`; `config_rus.json` se escribiría en la raíz del repo

La GUI crea `config_rus.json` en el directorio de trabajo (`gui/app.py:47,73-87`). PROMPT_CONTINUACION exige que nunca se distribuya. Resuelto ya en este commit inicial (`.gitignore` agregado).

### H-17 · INFO — `es_derivacion_sin_seg` incompleta y sin anclaje (confirmación de E-01)

`motor/utilidades.py:156-159`: solo reconoce OPD, DAM, RED SALUD, CONSULTA EXTERNA — faltan SALUD PRIVADA, HOSPITAL, UNIDAD DE SALUD MENTAL, CESFAM, COLEGIO y CHILE CRECE CONTIGO del catálogo E-01; y matchea por subcadena libre en vez de palabra completa anclada al inicio (corrección D1 del catálogo). Ej.: un programa hipotético "PRO**DAM**" dispararía el corte.
**Resuelve:** Fase F3 (y F1: helper compartido).

---

## 2. Verificación de los hallazgos de los documentos MD contra el código real

Todos los ítems aceptados de `Mejoras_Correos_y_Otros_Modulos_20260714.md` fueron verificados. Referencias exactas:

| Ítem MD | Verificado en | Confirmación |
|---|---|---|
| 1.1 correo tribunal vencidos | `comunicaciones/generador_correos.py:241-244` (llamada), `:374-405` (`_trib_borrador`), `plantillas/correo_tribunal_vencidos.html`, `contactos.json`, `gui/app.py:60-64,111` (`DEFAULT_CONTACTOS`/`self.contactos`) | ✅ Existe y se ejecuta para el grupo `vencidos`. Tras eliminarlo, `contactos.json`, `CONTACTOS_FILE` y `self.contactos` quedan sin otro uso → remover todo. |
| 1.2 `PATRON_ESPERA` | `comunicaciones/generador_correos.py:30` | ✅ El patrón actual no coincide ni con los textos v8.14 (`"...respecto de fecha estimada de ingreso"` — sin "la"/"efectivo", pero el patrón exige texto distinto) ni con los v9. Con el texto v9 aprobado, el patrón corregido del MD calza exactamente. |
| 1.3 DCE en plantillas | `plantillas/correo_programa_vencidos.html`, `correo_programa_por_vencer.html` (dicen "informes de avance" fijo) | ✅ No hay distinción DCE. `es_dce()` está en `motor/utilidades.py:150-153`. |
| 1.4/1.9 formato de programa duplicado | `generador_correos.py:35-54` (`_formatear_nombre_programa`, criterio posición 0-1 + ya-mayúscula) vs `motor/utilidades.py:87-104` (`titulo_programa`, criterio ≤4 letras) | ✅ Criterios distintos, confirmado. |
| 1.6 `DIAS_POR_VENCER` | `generador_correos.py:28` (`= 45`) | ✅ |
| 1.7 día 0 vencido | `generador_correos.py:230-232` (`d <= 0` → VENCIDO) | ✅ |
| 1.8/1.9 alias RUT/DERIVACIÓN | motor: `mapeo_columnas.py:44` (`["RUT","RUT MENOR"]`), `:24` (incluye "NOMBRE CENTRO") vs correos: `generador_correos.py:211` (`["RUT","RUT NNA","RUT LITIGANTE"]`), `:213` (sin "NOMBRE CENTRO") | ✅ Desalineados tal como describe el MD. |
| 1.10 tribunal crudo en asunto | `generador_correos.py:237,246,296` — agrupa y titula por `trib_raw` | ✅ Además: al normalizar el nombre hay que **agrupar por clave normalizada** (`_clave_trib`), no solo cambiar el texto, o dos variantes de escritura del mismo tribunal generan dos borradores (ver plan F7). |
| 2.1 `detectar_tipo()` | `resoluciones/generador_resoluciones.py:142-159` | ✅ Con matiz: los patrones PC_IE **sí** matchean los textos v8.14 vigentes (contienen "pidiendo cuenta a programa respecto al ie"), pero **ninguno** matchea los textos v9 del catálogo. El parche 2.1 es obligatorio en el mismo release que cambia los textos, o resoluciones cae a 0. PC_INFO ya no matchea nada hoy. NOMENCL se deja intacto (decisión del MD). |
| 2.2 aviso Hoja2 | `motor/procesador.py:213` (`"ℹ️ Sin Hoja2 — cruce desactivado"`, solo log) | ✅ |
| 2.6 código muerto validador | `validador/reglas_validacion.py:25-36` (`COLS_MINIMAS`, `COLS_CRITICAS_VACIAS`) | ✅ Sin referencias. |
| 2.12 índice Hoja2 `vd <= hoy` | `motor/procesador.py:56` | ✅ Confirmado sin cambio. |
| 2.13 supresión Hoja2 vs C-04/C-05 | `motor/procesador.py:184-186` (`not medida_vencida... and not r5_aplica...`) | ✅ El código ya suprime — conducta confirmada como correcta. Ojo: los umbrales de esos helpers cambian en v9 (C-04: `<0`; C-05: `0..45` + fecha válida) — actualizar `reglas_cumplimiento.py:138-145` en conjunto. |

Divergencias del código v8.14 respecto del **catálogo v2** (además de las anteriores): tramos de espera E-05 (hoy Mulchén ≥30 / Laja+Tomé ≥60 / DCE ≥30 **con texto de proyecto** — el Manual lo prohíbe para DCE), textos completos de casi todas las reglas, C-03 por fecha en vez de DÍAS DE CUMPLIMIENTO, C-04 sin rama día-0 ni chequeo de fecha, ausencia de rama "sin ficha individual" en C-07, prefijos residenciales (`pee`/`rppm` sobran, falta `rva` — `reglas_cumplimiento.py:21`), C-06/`proximo_informe()` a eliminar (`utilidades.py:239-259`, `reglas_cumplimiento.py:94-101`), T-02 excluye el día 0, columnas de salida §8 inexistentes. Todo esto está mapeado fase por fase en el plan.

---

## 3. Notas de riesgo para el implementador

1. **Los 73 tests de paridad van a fallar por diseño** al cargar los textos v9 (el propio catálogo §13 lo declara). No "arreglar" los tests para que pasen con textos viejos: regenerar goldens (`python tests/generar_goldens.py`) **solo después** de que el JSON v9 esté cargado y revisado. El guard `test_conteo_total_reglas_es_31` y el guard `<=17` pendientes de `test_paridad.py:120-140` deben actualizarse a los números v9 (ver plan, Apéndice A).
2. **Textos verbatim (restricción inviolable):** los textos del Apéndice A del plan provienen carácter por carácter del catálogo aprobado. No corregir ortografía, no "mejorar" redacción, no quitar/agregar puntuación. Las duplicaciones tipo doble "Medida revisada, a la espera de ingreso efectivo." cuando E-04+E-05 co-ocurren son **intencionales** (catálogo E-04: "Si también se cumple E-05, ambas se agregan").
3. **PROMPT_CONTINUACION.md contiene decisiones antiguas ya superadas** por el catálogo v2 (ej.: prefijos residenciales con PEE/RPPM, unión de fragmentos distinta por modo). Ante conflicto, **manda el catálogo v2 + el MD de mejoras**; PROMPT_CONTINUACION queda como contexto histórico.
4. **Outlook/pywin32 y Tkinter no existen en el sandbox de desarrollo.** Ningún test debe requerirlos (hoy ya es así). Mantener la separación: la lógica de armado de correos debe poder testearse sin `win32com` (ver F7).
5. **Los fixtures sintéticos (`tests/fixtures/generar_fixtures.py`) codifican las reglas viejas** (ramas Mulchén≥30/Laja-Tomé≥60, etc.). Hay que regenerarlos para las ramas v9 (ver F11); si no, el smoke test pierde cobertura de las ramas nuevas (Tomé 30/60, DCE≥30, día 0, etc.).
6. **Plantilla Word Tomé no existe** para proyectos de resolución (PC_IE): con E-05 v9, Tomé ≥60 días genera texto de proyecto → `detectar_tipo()` lo detectará → caerá al reporte de "faltantes" con `[SIN PLANTILLA]`, igual que hoy. Es comportamiento aceptado (el catálogo no aporta plantilla de Tomé); documentarlo, no inventar plantilla.
