# EVALUACIÓN DEL INFORME CODEX — Revisión senior 2026-07-15

**Evaluador:** auditoría independiente (Fable 5) contrastando cada afirmación contra el código real en `claude/zip-assistant-review-plan-7zfmyy` (HEAD `1c7f757`, v9.0.2).
**Veredicto global:** informe de calidad media-alta con diagnóstico general correcto ("los riesgos no son de exposición web sino de robustez local"), pero con **4 afirmaciones fácticas incorrectas**, varias propuestas de sobre-ingeniería impropias para una herramienta local de ~3.000 LOC con especificación funcional **cerrada**, y omisión del principal riesgo histórico real del proyecto: **las reescrituras grandes son las que han introducido regresiones** (v9.0.0 perdió el mapeo `ficha_ind`; revision3 rompió el guardado en Windows). Toda propuesta se evaluó también contra ese criterio.

Lo aceptado quedó operacionalizado en `docs/plan/PLAN_MEJORAS_V9_1.md` (plan ejecutable por Codex).

---

## 1. Correcciones fácticas al informe

| # | Afirmación de Codex | Realidad verificada |
|---|---|---|
| F-1 | "Un Excel grande puede **congelar Tkinter**" | Falso desde v9.0.2: TODAS las cargas corren en `threading.Thread` y la UI solo se toca vía `queue` (`gui/app.py:529,572,730,743,828`). El riesgo real es consumo de memoria del proceso, no congelamiento de la GUI. |
| F-2 | "La bitácora registra rutas y resultados [con posible PII]" | La bitácora operativa registra **solo métricas** por diseño documentado (`gui/app.py:420-423`: "nunca datos de personas/casos"); los eventos llevan conteos y tipos de error, no RIT/RUT/nombres. |
| F-3 | "Mensajes y strings de negocio dispersos… recomendación: catálogo central con IDs" | **Ya existe y es la pieza central del diseño**: `motor/textos_observaciones.json` con IDs, flag `confirmado` y trazabilidad al catálogo v2, protegido por goldens (`tests/test_paridad.py`). Las plantillas HTML son plantillas de correo, no "strings dispersos". |
| F-4 | "Aunque el HTML escapa contenido…" (en duda) | Confirmado que escapa: `validador/reporte.py` usa `html.escape(..., quote=True)` en detalle, ids, descripciones y estadísticas (líneas 85-131). |

---

## 2. Dictamen por hallazgo

### Hallazgos de alta prioridad de Codex

| Hallazgo | Dictamen | Fundamento |
|---|---|---|
| 1.1 Límite de filas/columnas al leer Excel | **ACEPTADO, reducido** (→ M1) | Legítimo como defensa; pero los Excel reales tienen ≤ ~350 filas y ya hay tope de 100 MB/200 MB descomprimido. Basta un prechequeo `openpyxl read_only` de dimensiones con umbral generoso (50.000 filas / 100 columnas) y rechazo temprano con mensaje claro. **No** es P0: es endurecimiento, no bug. Sin "límites configurables" (el catálogo §11 ya rechazó externalizar umbrales). |
| 1.2 PII en reportes HTML y bitácoras | **RECHAZADO el enmascarado; ACEPTADA la retención** (→ M3) | El reporte de validación **debe** mostrar RIT: su función es que el usuario identifique qué fila corregir; enmascararlo lo inutiliza. Es un archivo local en el equipo del funcionario cuyo trabajo es tratar esos datos. La bitácora ya es sin PII (F-2). Lo que sí falta: los `validacion_*.html` se acumulan sin retención — aplicar la misma rotación de 90 días de los logs, y documentar la matriz de salidas en SECURITY.md. |
| 1.3 Preflight de Outlook + dry-run | **ACEPTADO** (→ M2) | Dolor operativo real: Outlook cerrado/bloqueado hoy produce un error COM críptico por grupo. Preflight con mensaje accionable + exportación de borradores a `.html` cuando Outlook no está disponible (reutiliza el `despachador` inyectable que ya existe). |
| 1.4 Política central de saneamiento por destino | **RECHAZADO como refactor; ACEPTADO como documentación** (→ M3) | Cada salida ya tiene su defensa correcta y **testeada**: Excel `_neutralizar_formulas` (data_type='s'), HTML `html.escape`, correos `_mail_valido`, Word no ejecuta contenido. Un módulo `sanitizers.py` centralizado es indirección sin amenaza que la justifique. Se documenta la matriz destino→protección→test en SECURITY.md. |

### Hallazgos de prioridad media de Codex

| Hallazgo | Dictamen | Fundamento |
|---|---|---|
| 2.1 Inyección de reloj (`Clock`/`fecha_referencia`) | **RECHAZADO el refactor; ACEPTADA la meta vía tests** (→ M5) | Cambiar la firma de los tres motores y todos sus helpers es exactamente el tipo de cirugía que ha causado regresiones aquí. La meta (determinismo/reproducibilidad de pruebas) se logra con `time-machine` en dev-deps, **cero cambios en producción**. `FECHA_OBS` ya deja constancia de cuándo se procesó cada Excel. |
| 2.2 `except Exception` amplios | **ACEPTADO, acotado** (→ M7) | Los catch amplios de la GUI son deliberados (resiliencia); lo que falta es `logger.exception` técnico en esos puntos para no perder el traceback. Solo eso. |
| 2.3 Modo diagnóstico para `normalizar_match` | **RECHAZADO** | G-08 fijó la normalización del cruce como definitiva y el catálogo §11 rechazó validadores especulativos análogos. Si en producción aparecen cruces perdidos reales, se evalúa entonces con datos concretos. |
| 2.4 Advertir rutas UNC/nube | **RECHAZADO** | Herramienta local operada por el dueño del dato; una advertencia por ruta de red sería fricción sin modelo de amenaza. Nota operativa en SECURITY.md basta. |

### Estilo y refactorización

| Propuesta | Dictamen |
|---|---|
| Orden de imports (Ruff `I`) | **ACEPTADO** (→ M4, autofix de bajo riesgo). |
| Partir `RUSApp` en controladores por pestaña | **RECHAZADO por ahora**: refactor grande sobre el único módulo sin tests automatizables (Tkinter), riesgo alto/beneficio bajo. Reevaluar si la GUI crece. |
| Versiones divergentes en docstrings | **ACEPTADO** (→ M7): quitar versiones hardcodeadas de docstrings; la fuente es `motor/version.py`. |
| Dataclasses/TypedDict para resultados | **ACEPTADO opcional** (→ M7): valor real pero moderado; los contratos dict ya están fijados por tests. Solo si M1-M6 quedan verdes. |
| **Arquitectura en 4 capas (dominio/IO/aplicación/GUI)** | **RECHAZADO**. Sobre-ingeniería para este tamaño y contexto: las reglas ya son funciones casi puras con render externo; el riesgo dominante del proyecto es el spec-drift durante reescrituras, y una re-arquitectura total lo maximiza. La disciplina que protege este proyecto son el catálogo cerrado + los 168 tests, no las capas. |

### Complementos propuestos

| Propuesta | Dictamen |
|---|---|
| `pip-audit`/OSV en CI | **ACEPTADO** (→ M4): hay lockfile, costo mínimo, valor real. |
| Bandit en CI | **ACEPTADO** (→ M4): ya es dev-dep y el código pasa `-ll` hoy. |
| Ruff en CI | **ACEPTADO acotado** (→ M4): gate solo `F,B,E9` (bugs reales); las ~120 marcas de estilo heredado no bloquean. |
| gitleaks / secret scanning | **ACEPTADO** (→ M4, job simple). |
| pre-commit | **RECHAZADO**: el usuario final no desarrolla; CI cubre. |
| mypy incremental | **OPCIONAL P3** (→ M7, solo módulos nuevos). |
| pytest-cov con umbral | **ACEPTADO informativo** (→ M4): reporte sí; umbral duro no todavía. |
| Hypothesis (property-based) | **OPCIONAL P3**: útil para `get_date`/`normalizar`, no prioritario. |
| freezegun/time-machine | **ACEPTADO** (→ M5). |
| Modo CLI | **ACEPTADO** (→ M6): además habilita la regresión real en la máquina del usuario sin pytest/env vars. |
| Dry-run de correos (.html) | **ACEPTADO** (→ M2). `.eml` no: complejidad sin demanda. |
| Panel de diagnóstico | **ACEPTADO reducido** (→ M2: botón "Verificar entorno" con 5 chequeos). |
| Observabilidad JSONL | **RECHAZADO por ahora**: la bitácora actual (eventos clave-valor sin PII) es suficiente para un solo operador. |
| UI de gestión de catálogos | **RECHAZADO**; además compite con la herramienta futura **ya pactada** en el catálogo §8: el conteo trimestral TT/CC. Esa, y no esta, es la próxima funcionalidad legítima (queda como fase condicional en el plan, a confirmación del usuario). |
| SBOM CycloneDX + firma de binarios | **RECHAZADO/P4**: estándar corporativo sin cadena de distribución que lo exija (un ejecutable, un usuario). Se publica checksum SHA-256 del .exe en CI, que es la parte útil (→ M4). |

---

## 2-bis. Adenda — Auditoría de la implementación de Codex (PR #8, commit `7a10544`)

Mientras se redactaba esta evaluación, Codex publicó su propio plan (convergente con el de esta evaluación — se adoptó el suyo como canónico en `docs/plan/PLAN_MEJORAS_V9_1.md`) y ejecutó sus fases M1–M4 y M6 en un commit, mergeado por el usuario. Auditoría del resultado:

**Correcto y conservado:** M1 límites de dimensiones con `read_only=True` y `wb.close()` en `finally`; M2 preflight Outlook (`comunicaciones/outlook_preflight.py`, con `CoInitialize/CoUninitialize` correctos) + dry-run HTML (`exportar_borrador_html` con escritura atómica — aquí el `fsync` es válido porque el handle es de escritura) + checkbox y botón en GUI; M3 rotación de `validacion_*.html` por fecha del nombre (90 días) + matriz de saneamiento en SECURITY.md; M4 gates de CI (`ruff --select F,B,E9`, `bandit -ll`, cobertura informativa, checksum SHA-256 del .exe) con limpieza de los 8 avisos F/B pendientes; M6 CLI (`cli.py procesar|preview|correos --dry-run`) verificada de punta a punta en sandbox. Suite tras el merge: 172 passed.

**Defectos corregidos en esta auditoría (commit de cierre):**
1. **Regresión de versiones de Actions:** la rama de Codex nació antes del bump de Dependabot y su merge revirtió `ci.yml` a checkout@v4/setup-python@v5/upload-artifact@v4 (Dependabot habría reabierto los PRs #2–#5). Re-aplicado v7/v6/v7.
2. **`pip-audit` como gate bloqueante:** un CVE sin fix disponible habría bloqueado el CI del usuario. Marcado `continue-on-error: true` (reporta, no bloquea).
3. **M5 declarada pero no implementada:** `time-machine` quedó en dev-deps sin ningún uso, con restricción `<3` incompatible con la versión vigente (3.2.0) y sin entrada en el lock. Completado: `tests/conftest.py` congela el reloj de la suite en `pytest_configure` (antes de importar los módulos de tests, que calculan fechas a nivel de módulo), lock actualizado, restricción `<4`.

**Notas menores (sin acción):** `_verificar_outlook` corre en el hilo principal de Tk (un `Dispatch` lento congela la ventana unos segundos — aceptable para un botón de diagnóstico); el job de gitleaks quedó fuera (era opcional en ambos planes).

## 3. Conclusiones

1. **El diagnóstico base de Codex es correcto** (robustez local > exposición web) y validó que los controles v9.0.2 existen y funcionan. Ninguno de sus hallazgos es un bug activo: son endurecimientos.
2. **Se aceptan 11 propuestas (7 reducidas), se rechazan 12.** El criterio de corte fue uniforme: valor operativo demostrable para UN operador local con spec cerrada, y riesgo de regresión del cambio. Las priorizaciones P0 de Codex se degradaron: nada de lo hallado bloquea la operación actual.
3. **El rechazo más importante es la re-arquitectura en capas.** La historia del propio repo (dos reescrituras, dos tandas de regresiones críticas) demuestra que el mayor riesgo de este código es tocarlo a lo grande. La inversión correcta es CI más estricto y determinismo de pruebas — exactamente lo aceptado.
4. El plan ejecutable con fases, criterios de aceptación y exclusiones explícitas está en **`docs/plan/PLAN_MEJORAS_V9_1.md`**.
