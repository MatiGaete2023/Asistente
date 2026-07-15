# AUDITORÍA TÉCNICA — Asistente_v9.0.1_revision3.zip

**Fecha:** 15 de julio de 2026 · **Motor de análisis:** Fable 5 · **Formato:** informe técnico Markdown

### 1. SPECIFICATIONS

- **STACK:** Python 3.12 · pandas / openpyxl / xlrd · python-docx · Tkinter (GUI escritorio Windows) · pywin32/Outlook COM · pytest · PyInstaller · GitHub Actions.
- **TARGET:** CSMP Assistant (Asistente de RUS) — generación supervisada de observaciones judiciales (ESPERA/CUMPLIMIENTO/INFORMES), borradores de correo Outlook y resoluciones .docx para el Centro de Seguimiento de Medidas de Protección de Concepción (tribunales Laja, Mulchén y Tomé).
- **CRITERIA:** Seguridad, Rendimiento, Mantenibilidad, Escalabilidad — **y conformidad con la especificación funcional cerrada** (`docs/especificaciones/Catalogo_Reglas_CSMP_v2` + `Mejoras_Correos_y_Otros_Modulos`), que en este proyecto es un criterio de primera clase: una "mejora de seguridad" que contradice una decisión cerrada del catálogo es un defecto, no una mejora.
- **Objeto auditado:** el ZIP `Asistente_v9.0.1_revision3` (revisión de otra IA sobre la implementación v9 del repo, HEAD `2259752`). Se auditaron ambos: la implementación base y la revisión.

**Contexto de línea base.** La implementación v9 mergeada (PR #1) llegó con la suite **roja** (4 tests de `test_confirmacion_textos` fallando), en un solo commit (el plan exigía uno por fase) y sin los tests de reglas del Apéndice C. La revisión3 corrigió parte de eso y agregó infraestructura valiosa, pero introdujo regresiones funcionales propias. El detalle sigue.

---

### 2. MODULE A — [CRITICAL SEVERITY] BUGS & SECURITY

#### A-1 · `os.fsync` sobre handle de solo lectura — el guardado fallaría SIEMPRE en Windows (revision3)

`motor/procesador.py` (también replicado en `resoluciones/`):

```python
wb.save(temporal)
with open(temporal, "rb") as f:
    os.fsync(f.fileno())      # Windows: _commit → FlushFileBuffers exige GENERIC_WRITE
os.replace(temporal, arch)
```

En Windows (la plataforma de producción), `FlushFileBuffers` sobre un descriptor abierto en `"rb"` falla con `OSError`; el `except` genérico lo captura, borra el temporal y reporta "❌ Error al guardar". Resultado: **ningún Excel ni Word se guardaría jamás en producción**, con la suite verde en Linux/CI. **Corregido:** guardado atómico `save → os.replace` sin fsync.

#### A-2 · Anti-inyección de fórmulas que corrompe datos legítimos (revision3)

`valor_excel_seguro` anteponía `'` a todo string que empezara con `=`, `+`, `-` o `@`. Dos defectos: (a) el placeholder institucional `---` (usado en RUT/fechas vacías) y cualquier valor con `-` inicial salía como `'---` **visible** en el Excel que se pega en RUS — el apóstrofe escrito programáticamente forma parte del valor, no es el prefijo de UI de Excel; (b) en openpyxl solo el prefijo `=` genera fórmula — `+`, `-`, `@` eran sobre-bloqueo. **Corregido:** `_neutralizar_formulas()` fuerza `data_type='s'` en las celdas afectadas — el texto queda **exacto** y nunca ejecutable. Test actualizado verifica `celda.value == original`.

#### A-3 · Hoja2 con duplicados → `ValueError` que aborta todo CUMPLIMIENTO (revision3, contra catálogo)

El catálogo C-10 es explícito: *"Si hay duplicados en Hoja2 se conserva la última fila encontrada"*. La revisión lanzaba `ValueError` ante claves duplicadas con fechas distintas y también cuando Hoja2 venía malformada (columnas faltantes), abortando el procesamiento completo del día por un defecto en la **segunda** hoja. **Corregido:** duplicados → gana la última fila; Hoja2 malformada → cruce desactivado + advertencia prominente en GUI (`warn_hoja2`, Mejoras §2.2), que ahora también se emite cuando el índice queda vacío.

#### A-4 · Resolución de contactos degradada a "solo match exacto" — correos silenciosamente no generados (revision3)

`CatastroContactos.resolver()` pasó a exigir igualdad normalizada exacta; el fuzzy quedó tras `permitir_fuzzy=True` que **ningún llamador usaba**. El propio ejemplo del docstring ("AFT EL CONQUISTADOR YUMBEL" vs "AFT - EL CONQUISTADOR **DE** YUMBEL") deja de resolver → "Sin contacto" → borrador no creado, sin ruido. Regresión funcional grave para el flujo diario. **Corregido:** `generador_correos` invoca con `permitir_fuzzy=True`; se **conservan** las mejoras reales de la revisión (ambigüedad → `None` en vez de "gana el más largo", claves ambiguas vetadas): ante duda el caso cae a "Sin contacto" y lo decide un humano.

#### A-5 · Lotes de correos "todo o nada" (revision3, contra decisión D9)

Cualquier fila con campo vacío, fecha inválida o tribunal no reconocido abortaba el lote **completo** (0 borradores de 333 por una celda mala), y dejaba código muerto (la re-clave por tribunal crudo tras el `return`). D9 del plan dice lo contrario: fila mala fuera con aviso, el resto sigue; tribunal no reconocido se agrupa por valor crudo. **Corregido:** `_excluir_filas_incompletas()` + omisión selectiva con detalle en `resultado["errores"]`. Tests de degradación suave incluidos.

#### A-6 · Bugs de composición heredados de la implementación base (HEAD, detectados en esta auditoría)

- `_complementarias(...)[:2]` y `comp[-1]` asumían lista de largo fijo: si curador u oído no disparaban, la **audiencia se insertaba antes de las fichas** (violación del orden §9.2) o, en INFORMES, se anexaba **curador/oído** (que no aplican en ese módulo). La revisión3 arregló INFORMES (`_audiencia()`) pero en CUMPLIMIENTO movió la audiencia *siempre* antes de las fichas. **Corregido:** `_curador_oido()` + fichas + `_audiencia()` al final; test de orden `curador < oído < ficha < audiencia`.
- **Crédito a revision3:** detectó y reparó que HEAD había eliminado el mapeo `ficha_ind` → **C-07 (ficha individual) jamás disparaba** en la v9 mergeada. Fix conservado con su test.

#### A-7 · Suite no coleccionable sin Tkinter (revision3)

`tests/test_config.py` importaba `gui.app` (→ `tkinter`) a nivel de módulo: en un entorno headless la **colección completa** de pytest muere (`1 error during collection`). **Corregido:** `pytest.importorskip("tkinter")`.

#### Seguridad — aportes de revision3 que se CONSERVARON

`validar_archivo_excel()` (extensión/tamaño/zip-bomb), escape HTML en tablas de correo, `_mail_valido()` (impide inyección de destinatarios vía `;`/saltos de línea en el catastro), CC institucional no configurable, `pythoncom.CoInitialize/CoUninitialize` + cierre de inspector en `finally`, guard de plantillas incompletas, config atómica validada en `LOCALAPPDATA`, catálogo de textos editable fuera del bundle PyInstaller, bitácora local sin datos personales, CI multi-OS + CodeQL + Dependabot + `requirements.lock`. No se detectaron secretos embebidos ni datos personales en fixtures (la revisión además limpió nombres reales — correcto).

**Rendimiento:** sin cuellos de botella nuevos; volúmenes (≤ ~350 filas) holgados para `df.apply`. El validador conserva loops por fila (H-13, pendiente menor).

---

### 3. MODULE B — [MEDIUM SEVERITY] REFACTOR & CLEAN CODE

1. **Violaciones de decisiones cerradas (spec-drift), todas revertidas:**
   - `ALIAS_RIT` ampliado a "N° RIT"/"N RIT"/… — la Mejora **2.5 lo rechazó explícitamente**. Revertido a `["RIT"]`.
   - `normalizar_match()` delegada a `normalizar()` (quita tildes y guiones) — **G-08 exige conservar la normalización actual** del cruce; cambiaba qué filas matchean (p. ej. RIT `P-123-2026` → `p 123 2026`). Restaurada la implementación histórica.
   - `resoluciones/`: la Mejora 2.1 autorizaba **solo** el parche de frases ("mismo mecanismo"). La revisión rediseñó el flujo: filas sin datos **omitidas** en vez de `COMPLETAR`, sin bloques `[SIN PLANTILLA]`, aborto total ante un error de fila, y tocó el fallback de `extraer_ncl()` (NOMENCL debía quedar **intacto**). Restaurado el flujo aprobado, conservando `validar_archivo_excel` y guardado atómico.
   - `es_dce(contacto["nombre"])` — la Mejora 1.3 fija el criterio sobre la **derivación de RUS** (mismo dato que el motor). Revertido a `es_dce(prog)`.
   - Alias inventados (`FECHA ACTUALIZACION FICHA INDIVIDUAL`) sin respaldo en spec. Eliminados.
   - Carga de `aliases_programas.json` con `raise` ante entrada malformada — el diseño original es "opcional, nunca bloquear". Vuelta tolerante (entrada mala se omite), conservando el veto de ambigüedades.
2. **DRY/KISS resueltos por revision3 y conservados:** `_leer_cumplimiento()` unifica la carga duplicada procesar/preview (H-11); `_audiencia()` extraída; excepciones estrechadas (`except Exception` → tipos concretos); docstrings sin nombres de personas reales.
3. **Estilo pendiente (no bloqueante):** `reglas_*.py` mantienen líneas multi-sentencia con `;` (herencia de la implementación base); nombres de un carácter (`_d`, `_pn`). Funcional y testeado; refactor cosmético queda a criterio del mantenedor — no se tocó para no inflar el diff de esta auditoría.
4. **`.gitignore` con `*.xlsx` global** ignoraba los fixtures de tests nuevos. Añadida excepción `!tests/fixtures/*.xlsx`.
5. **Incidencias G-05 ruidosas** (una por fila cuando faltaba la columna completa): ahora solo se registran si la columna existe y el valor es inválido.

**Tipado/nombres según estándar:** los añadidos de revision3 vienen con type hints correctos; no se detectaron anomalías adicionales de tipado que justifiquen cambio.

---

### 4. MODULE C — [PROPOSALS] ENHANCEMENTS & COMPLEMENTS

**C-1 · Antes/Después aplicados en este commit (los dos más críticos):**

*Guardado Excel (A-1/A-2):*
```python
# ANTES (revision3)
for r in dataframe_to_rows(df, index=False, header=True):
    ws.append([f"'{v}" if isinstance(v, str) and v.lstrip().startswith(("=","+","-","@")) else v
               for v in r])
wb.save(temporal)
with open(temporal, "rb") as f:
    os.fsync(f.fileno())          # OSError en Windows → guardado siempre fallido
os.replace(temporal, arch)

# DESPUÉS (v9.0.2)
for r in dataframe_to_rows(df, index=False, header=True):
    ws.append(r)
_neutralizar_formulas(ws)         # data_type='s' → texto EXACTO, nunca fórmula
wb.save(temporal)
os.replace(temporal, arch)        # atómico; sin fsync sobre handle de lectura
```

*Orden §9.2 en cumplimiento (A-6):*
```python
# ANTES (revision3)
frags += _complementarias(row, cols)   # audiencia queda ANTES de las fichas
... fichas C-07 / C-08 ...

# DESPUÉS (v9.0.2)
frags += _curador_oido(row, cols)      # T-01, T-02
... fichas C-07 / C-08 ...
aud = _audiencia(row, cols)            # T-03 SIEMPRE al final
if aud: frags.append(aud)
```

**C-2 · Cobertura completada (era el mayor incumplimiento del plan):** se agregaron los tests de borde del Apéndice C que ni la implementación ni la revisión escribieron — `test_reglas_espera.py` (21), `test_reglas_cumplimiento.py` (23), `test_reglas_informes.py` (8), `test_correos.py` (7), `test_salida_excel.py` (6), incluyendo los ejemplos §10 del catálogo byte a byte, los límites 29/30/59/60 de E-05, la supresión C-10 vs C-04/C-05 y la ausencia del default "LAJA". Suite final: **168 passed, 4 skipped** (antes: 102 + colección rota; HEAD: 63 + 4 FAILED).

**C-3 · Herramientas de terceros:** la revisión ya incorporó lo razonable (CI matriz ubuntu+windows con build PyInstaller verificado, CodeQL, Dependabot, lockfile, `defusedxml`, ruff/mypy/bandit como extras dev). Recomendaciones restantes, en orden de valor: (1) job de CI que ejecute `ruff check` y `bandit -r motor comunicaciones` que hoy están declarados pero no se corren; (2) `pytest-randomly` para detectar dependencias de orden; (3) vectorizar A4–A7 del validador con pandas si el volumen crece (hoy no es cuello de botella). No se recomienda agregar frameworks nuevos: la app es local, síncrona y de bajo volumen — más dependencias = más superficie.

**Anomalías no encontradas:** fugas de memoria (proceso batch de vida corta, sin referencias cíclicas relevantes); condiciones de carrera en la GUI (revision3 canalizó todos los workers por `queue` — corrige H-05 —, `RLock` en caché de textos e importación bloqueada durante ejecución); inyección SQL/OS-command (no hay SQL ni subprocesos con entrada de usuario fuera del runner de tests, que usa rutas propias).

---

### 5. Veredicto y estado del repositorio

| Ítem | Estado |
|---|---|
| Aportes de revision3 (seguridad, GUI 2.10, thread-safety, empaquetado, CI, `ficha_ind`) | **Integrados** |
| Regresiones de revision3 (A-1…A-5, A-7, spec-drift de Módulo B) | **Corregidas** |
| Bugs heredados de HEAD (orden §9.2, `comp[-1]`, suite roja) | **Corregidos** |
| Tests Apéndice C faltantes | **Escritos** (65 tests nuevos) |
| Versión | **v9.0.2** (motor/version.py, pyproject.toml) |

Pendientes conocidos (documentados, no bloqueantes): regresión real contra los 3 Excel de producción en la máquina del usuario (`CSMP_EXCEL_*`), build .exe real en Windows (CI ya lo compila como artefacto), fixtures sintéticos aún con ramas v8 (los bordes v9 quedaron cubiertos por tests unitarios directos), estilo compacto de `reglas_*.py`.
