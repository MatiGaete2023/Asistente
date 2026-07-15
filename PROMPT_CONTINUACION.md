# PROMPT DE CONTINUACIÓN — CSMP Assistant v8.14.0

Copia este bloque completo al inicio del próximo chat junto con los archivos indicados al final.

---

## CONTEXTO

Estoy trabajando en **CSMP Assistant** (Asistente de RUS), aplicación Python desktop Tkinter para el Centro de Seguimiento de Medidas de Protección de Concepción. Tres tribunales atendidos: Laja, Mulchén, Tomé.

**Versión actual:** v8.14.0 — S1-S6 del Plan de Ejecución Maestro ejecutados y verificados (micro-fixes, suite de tests permanente, vista previa GUI, flujo de confirmación de textos, consistencia visual, build PyInstaller validado).

**Estado:** Motor completo (3 modos) con 31 reglas de texto externalizadas a `motor/textos_observaciones.json`. GUI con 3 pestañas + vista previa en memoria + confirmación de textos sin sesiones de chat. Suite `tests/` permanente (73 passed, 3 skipped). Todos los bugs reportados de producción están corregidos.

## REGLAS DE TRABAJO

- Sigue mis `userPreferences` (español Chile coloquial, respuestas directas sin preámbulos, código completo no fragmentos, N1/N2/N3 epistemológicos).
- Antes de tocar código: leer el archivo real (`view` o `bash_tool`) — no asumir nada del snapshot que pueda haber cambiado.
- Antes de proponer optimizaciones: medir con benchmarks reales, no creer afirmaciones genéricas.
- **Antes de cualquier entrega: `pytest tests/ -v` en verde.** Si se sabotea o cambia un texto en `textos_observaciones.json` sin correr `tests/generar_goldens.py`, el test de paridad debe fallar — eso es correcto, no un bug.
- Cuando entregue un ZIP: **siempre completo** con todos los módulos, nunca parcial (causa `ModuleNotFoundError`). Verificar `textos_observaciones.json` presente y `config_rus.json` AUSENTE (se regenera portable).
- Para cualquier cambio en motor/reglas: correr `pytest tests/` y, si hay Excel reales disponibles, `test_regresion.py` con las env vars `CSMP_EXCEL_*`.
- Restricciones inviolables: sin envío automático de correos, sin acceso directo a RUS, CC fijo `ucc_concepcion@pjud.cl`, textos de observación siempre copy-paste verbatim (nunca parafraseados).

## DECISIONES YA TOMADAS — NO REVISAR

- **R6 curador**: detección por RUT real (`\d{6,8}-[\dkK]`) o curador institucional (`'Institución:' in valor`, agregado v8.14 S1) via `tiene_curador_real()`.
- **R8 ficha residencial**: solo prefijos RTA/RTT/RES/PEE/RFA/RPPM. Sin PRM (excluido explícitamente). No PIE/AFT/PAS/FAE.
- **Cruce Hoja2**: filtro `vd > hoy` estrictamente futuro. Match 5 campos.
- **Plantillas correo**: encabezado `SEÑORES: / {programa} / PRESENTE.` (PRESENTE con negrita + subrayado).
- **Firma Outlook**: `mail.Display(False)` antes de leer HTMLBody.
- **sys.path**: configurado al inicio de `gui/app.py`.
- **Textos de observación**: SOLO desde `motor/textos_observaciones.json` vía `render()`. Nunca hardcodear en `.py`. Nunca marcar `confirmado:true` por inferencia de una IA — solo Matías puede confirmar (flujo S4 en la GUI, o exportar/editar/importar el Excel).
- **ESPERA vs CUMPLIMIENTO, unión de fragmentos**: ESPERA usa `". ".join()` sin punto final en cada fragmento; CUMPLIMIENTO usa `" ".join()` con punto final en cada fragmento. Mezclar convenciones causa el bug de doble punto.
- **Vista previa (S3)**: `calcular_preview()` en `motor/procesador.py` es la ÚNICA función que debe usar la GUI para previsualizar — garantiza paridad byte-a-byte con el export real porque comparte el mismo código (`_calcular_simple`/`_calcular_cumplimiento`). No reimplementar la lógica de reglas en la GUI.

## PENDIENTE EXPLÍCITO (bloqueado en Matías, no en código)

### 1. Confirmar los 17 textos con `confirmado:false`
Usar el flujo S4 en la GUI (pestaña Motor → "Confirmación de textos" → "Exportar pendientes…"). Editar el Excel, importar. Después correr `python tests/generar_goldens.py` para que el test de paridad adopte el nuevo texto.

### 2. Regresión real contra los 4 Excel de producción
Este sandbox de desarrollo no tuvo acceso a los Excel reales — la validación se hizo con `tests/fixtures/generar_fixtures.py` (datos sintéticos, solo smoke test estructural). Correr en tu máquina:
```
CSMP_EXCEL_ESPERA=ruta\excel_espera.xlsx
CSMP_EXCEL_CUMPLIMIENTO=ruta\excel_cumplimiento.xlsx
CSMP_EXCEL_INFORMES=ruta\excel_informes.xlsx
pytest tests/ -v
```

### 3. Compilar el .exe real en Windows
```
pyinstaller --noconsole --onefile --add-data "motor/textos_observaciones.json;motor" --name "Asistente_RUS" main.py
```
(separador `;` en Windows). Se validó la lógica en Linux (el JSON se extrae correctamente en `_MEIPASS`) — falta el build real y probarlo contra un Excel de cada modo.

## OPCIONES PARA LA PRÓXIMA FASE

Elegir UNA y avisar si se quiere trabajar en algo distinto:

### Opción A — FASE 5: Reportes consolidados
4 reportes adicionales: R1 NNA en residencia, R2 órdenes de búsqueda, R3 trimestral, R4 semanal. Sin spec detallada — habría que definirla. Requiere plantillas (ver `Resumen_planteamiento_proyecto`).

### Opción B — Plantilla PC_INFO Laja
Falta la plantilla PC_INFO para Juzgado de Laja en `resoluciones/generador_resoluciones.py`. Bajo esfuerzo.

### Opción C — S3b (opcional, quedó fuera de v8.14)
Columna `REGLAS_APLICADAS` en el Excel de salida (checkbox default OFF en la GUI). Requiere que las funciones de reglas retornen `(texto, ids)` con wrapper retrocompatible.

### Opción D — Fase 6: auto-generar CALENDARIO desde logs del motor
Mencionada como ahorro de ~2h/mes en sesiones anteriores. Sin spec.

### Opción E — Otra cosa que surja en producción
Bugs reportados, screenshots, mensajes de error. Diagnóstico → fix → `pytest tests/` verde → ZIP completo.

## ARCHIVOS A SUBIR EN EL PRÓXIMO CHAT

**Obligatorios (en este orden):**
1. `SNAPSHOT.json` — estado consolidado del proyecto
2. `CHANGELOG.md` — historial de cambios
3. `README.md` — instrucciones operacionales

**Para tocar código del motor:**
4. `motor/utilidades.py`
5. `motor/procesador.py`
6. `motor/reglas_cumplimiento.py`
7. `motor/reglas_espera.py`
8. `motor/reglas_informes.py`
9. `motor/mapeo_columnas.py`
10. `motor/textos.py` + `motor/textos_observaciones.json` + `motor/textos_confirmacion.py`

**Para tocar tests:**
11. `tests/test_paridad.py`, `tests/goldens_textos.json`, `tests/runner.py`, `tests/test_regresion.py`

**Para tocar correos:**
12. `comunicaciones/generador_correos.py`
13. `comunicaciones/contactos_programas.py`
14. `comunicaciones/aliases_programas.json`
15. `comunicaciones/catastro_programas.json` (231 programas — grande, solo si toca)
16. `comunicaciones/plantillas/*.html`

**Para tocar resoluciones:**
17. `resoluciones/generador_resoluciones.py`

**Para tocar GUI:**
18. `gui/app.py`
19. `main.py`

**Archivos de prueba reales (subir solo cuando se necesite reproducir un bug o correr regresión real):**
- Excel ESPERA / CUMPLIMIENTO (con Hoja2) / INFORMES de producción.

## PRIMER MENSAJE SUGERIDO PARA EL PRÓXIMO CHAT

> "Carga el snapshot. Antes que nada: ¿ya confirmaste los 17 textos pendientes vía el flujo de la GUI? Si sí, pásame el `textos_observaciones.json` actualizado y corro `generar_goldens.py`. Si no, avísame cuál opción tomamos: A (Fase 5), B (plantilla PC_INFO Laja), C (columna REGLAS_APLICADAS), D (Fase 6 calendario), E (bug nuevo)."
