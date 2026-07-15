# PLAN DE IMPLEMENTACIÓN — CSMP Assistant v9.0

**Fecha:** 15 de julio de 2026
**Autor del plan:** revisión y planificación senior (esta sesión). **Ejecutor previsto:** agente de codificación (Codex / Sonnet) siguiendo este documento fase por fase.
**Insumos normativos (mandan sobre cualquier otro documento del repo):**

1. `docs/especificaciones/Catalogo_Reglas_CSMP_v2_20260714.md` — "el catálogo" (reglas E/T/C/I/G + §8 columnas de salida + §13 criterios de prueba).
2. `docs/especificaciones/Mejoras_Correos_y_Otros_Modulos_20260714.md` — "las mejoras" (correos, validador, resoluciones, GUI).
3. `docs/auditoria/AUDITORIA_CODIGO_v8_14.md` — hallazgos H-01…H-17 de esta revisión.

**Punto de partida:** el código v8.14 tal cual está en la raíz de este repositorio (commit inicial). Suite verificada: `pytest tests/ -q` → 73 passed, 3 skipped.

---

## 0. Instrucciones para el agente ejecutor

- Ejecutar las fases **en orden** (F1 → F12). Cada fase termina con sus criterios de aceptación en verde y un commit propio con mensaje `F<N>: <resumen>`. No mezclar fases en un commit.
- **Regla inviolable de textos:** los textos de observación del Apéndice A se copian **carácter por carácter** al JSON. Nunca parafrasear, corregir ortografía ni ajustar puntuación. Si un texto parece redundante (p. ej. doble "Medida revisada, a la espera de ingreso efectivo." al co-ocurrir E-04+E-05), es intencional.
- **Conflictos entre documentos:** catálogo v2 + mejoras > este plan > PROMPT_CONTINUACION.md (histórico, contiene decisiones superadas). Si este plan contradijera el catálogo en el texto de una regla, manda el catálogo.
- **No inventar alcance:** todo lo listado en el catálogo §11 como "rechazado" queda fuera. No agregar reglas para ficha ambulatoria, no externalizar umbrales a config, no calcular carga/sin carga, no ampliar alias de RIT, no tocar NOMENCL en resoluciones.
- **Entorno:** Python ≥3.10, `pip install pandas openpyxl xlrd python-dateutil python-docx pytest`. No hay Tkinter con display ni Outlook/pywin32 en el sandbox: ningún test puede requerirlos (hoy ya es así; mantenerlo).
- Después de F2, `pytest` estará **rojo por diseño** hasta regenerar goldens en la misma fase (ver F2.5). En cualquier otra fase, la suite debe quedar verde antes de commitear.
- Los tests nuevos de reglas van en `tests/test_reglas_espera.py`, `tests/test_reglas_cumplimiento.py`, `tests/test_reglas_informes.py`, `tests/test_composicion.py`, `tests/test_correos.py`, `tests/test_resoluciones.py`, `tests/test_salida_excel.py` (crearlos donde se indique).

### Definición de HECHO global (verificar al final, antes del release)

1. `pytest tests/ -v` completamente verde en Linux, sin Outlook ni display.
2. Todos los criterios del catálogo §13 tienen test que los cubre (matriz del Apéndice C completa).
3. `grep -rn "proximo_informe\|R5B\|ficha_res\|_trib_borrador\|correo_tribunal_vencidos\|COLS_MINIMAS\|COLS_CRITICAS_VACIAS\|contactos_tribunales" --include="*.py" .` → sin resultados (código muerto realmente eliminado).
4. Ninguna observación de salida contiene `{`, `}`, `..`, `ERROR` ni "None".
5. `motor/textos_observaciones.json`: 25 textos, todos `confirmado: true`, idénticos al Apéndice A.
6. La GUI compila (`python -m py_compile gui/app.py`) y los módulos importan sin display (`python -c "import motor.procesador, comunicaciones.generador_correos, resoluciones.generador_resoluciones, validador.precheck"`).
7. README y CHANGELOG actualizados a v9.0.0.

---

## 1. Arquitectura objetivo (resumen de cambios estructurales)

```
motor/
  columnas_comunes.py    ← NUEVO (F1): alias de columnas + funciones compartidas motor↔correos
  composicion.py         ← NUEVO (F1): prefijo, unión de fragmentos, puntuación, incidencias
  textos.py              (igual; render() sobre el JSON v9)
  textos_observaciones.json  ← REESCRITO (F2): 25 textos v9, todos confirmados
  reglas_espera.py       ← REESCRITO (F3): E-01…E-06 + T-01…T-03
  reglas_cumplimiento.py ← REESCRITO (F4): C-01…C-10 (C-06 eliminada; Hoja2 dentro del motor)
  reglas_informes.py     ← REESCRITO (F5): corte E-01 + I-01/I-02
  utilidades.py          ← PODADO (F1/F4): sin proximo_informe/relativedelta; helpers nuevos
  mapeo_columnas.py      ← AJUSTADO (F1/F4): usa columnas_comunes; sin ficha_res/ficha_amb; + dias_cumpl ya existe
  procesador.py          ← AJUSTADO (F6): sin default LAJA, columnas FECHA_OBS/TT/CC/RES,
                            hoja VALIDACION, carga CUMPLIMIENTO unificada, aviso Hoja2
  version.py             ← NUEVO (F10): VERSION = "v9.0.0"
comunicaciones/
  generador_correos.py   ← AJUSTADO (F7): sin correo tribunal; DCE; 30 días; día 0;
                            alias/formato desde columnas_comunes; tribunal normalizado;
                            armado de borradores separado del envío (testeable)
  plantillas/            ← correo_tribunal_vencidos.html ELIMINADA; etiqueta {{ETIQUETA_INFORMES}}
resoluciones/
  generador_resoluciones.py ← PARCHE (F8): frases de detectar_tipo(); NOMENCL intacto; resto NO se toca
validador/
  reglas_validacion.py   ← PODADO+OPTIMIZADO (F9); precheck cableado al pipeline
gui/
  app.py                 ← AJUSTADO (F10): thread-safety, 2 campos en Correos, aviso Hoja2, versión
contactos.json           ← ELIMINADO (F7, junto con el correo a tribunal)
tests/                   ← ampliado (por fase) + goldens/fixtures regenerados (F2/F11)
```

Sin cambios: `motor/textos_confirmacion.py` (flujo S4 queda igual — mejoras 2.8/2.9 rechazadas), `logs/log_manager.py` (se cablea en F10, tarea opcional), `resoluciones/` fuera del parche 2.1 (mejora 2.3 rechazada), vista previa S3 (`calcular_preview`) que sigue siendo la única vía de preview.

---

## 2. Decisiones

### 2.1 Cerradas (vienen de los documentos — no re-litigar)

- E-05 por tramos: Laja/Mulchén ≥30 → proyecto+correo; Tomé 30–59 → solo correo; Tomé ≥60 → proyecto+correo; DCE ≥30 → solo correo **nunca proyecto**, cualquier tribunal.
- C-06 eliminada por completo (incluye `proximo_informe()` y `relativedelta`). Hoja2 (C-10) es fuente única y **se suprime** si C-04 o C-05 dispararon.
- Fichas: C-07 fusionada usa **solo** `FEC.ACT.F.INDIVIDUAL`; prefijos residenciales exclusivos RTA/RTT/RES/RFA/RVA; ficha FAE solo rama ausencia; ficha ambulatoria **nunca**.
- Sin tribunal por defecto (G-06). Sin marcadores `{}` en salida (G-05). Fechas largas en español (G-02). Criterio de sigla por largo ≤4 en `titulo_programa` se mantiene (G-03 rechazó cambiarlo).
- Columnas de salida §8: `FECHA_OBS` automática antes de OBSERVACION; `TT`/`CC`/`RES` vacías después.
- Correos: sin correo a tribunal (1.1); patrón espera corregido (1.2); DCE distinto (1.3); 30 días (1.6); día 0 = por vencer (1.7); unificación de alias/formato motor↔correos (1.9) **sin** incluir resoluciones (2.3); tribunal normalizado en asunto/cuerpo (1.10); saludo genérico se mantiene (1.5 rechazada).
- Resoluciones: solo parche de frases (2.1), NOMENCL intacto.
- Validador: eliminar `COLS_MINIMAS`/`COLS_CRITICAS_VACIAS` (2.6); no tocar A6 (2.7) ni A7 (2.4) ni alias RIT del motor (2.5).
- GUI: advertencia prominente sin Hoja2 (2.2); dos campos de Excel en Correos (2.10); vista previa sigue mostrando `head(30)` (2.11 rechazada).

### 2.2 Abiertas — resueltas en este plan con default (si el usuario no dice lo contrario, implementar el default)

| # | Punto | Default a implementar |
|---|---|---|
| D1 | Formato de `FECHA_OBS` | Texto `dd/mm/yyyy` con la fecha del día de procesamiento (formato corto: es una columna de gestión, no un texto de observación; G-02 aplica a observaciones). |
| D2 | Valor inicial de `TT`/`CC`/`RES` | Cadena vacía (no 0) — catálogo §13. |
| D3 | Cumplimiento donde solo dispara C-02 (próxima mayoría) | Tratarla como las complementarias para C-09: base breve "Medida revisada." + fragmento C-02. |
| D4 | C-01 (mayoría en cumplimiento): qué acumula | Curador, oído y próxima audiencia (idéntico a E-02). No fichas, no C-02…C-10. |
| D5 | Duplicaciones de texto al co-ocurrir E-04+E-05 o C-03+C-10 | Aceptarlas — los textos son verbatim y el catálogo dice que ambas se agregan. |
| D6 | Cómo se "registra en validación" una regla omitida por dato faltante (G-05) y el tribunal no reconocido (G-06) | El motor acumula incidencias `{fila_excel, rit, regla, motivo}`; el procesador las escribe en una **segunda hoja "VALIDACION"** del mismo Excel de salida y resume el conteo en la bitácora GUI. |
| D7 | Semántica de "prefijo residencial" | `programa_normalizado.startswith(("rta","rtt","res","rfa","rva"))` — prefijo de string (cubre "RESIDENCIA …" vía "res"), igual que hoy pero con la lista del catálogo. |
| D8 | Asunto de correos DCE | También cambia: "Informes diagnósticos vencidos — …" / "Informes diagnósticos por vencer — …". |
| D9 | Correos con tribunal no reconocido | Se agrupan por el valor crudo y se agrega advertencia al resultado (no se pierden filas). |
| D10 | INFORMES sin regla disparada (fecha lejana o sin fecha) | Celda OBSERVACION vacía (`""`), sin audiencia; si falta la fecha → además incidencia G-05. |
| D11 | Espera ≥30 no-DCE con tribunal no reconocido (E-05 inaplicable) | E-06 actúa como fallback genérico + incidencia G-06. E-06 se define operacionalmente como "no disparó ni E-04 ni E-05" (§9.1 lo llama "fallback"). |
| D12 | Columnas §8 ¿en los 3 modos? | Sí: ESPERA, CUMPLIMIENTO e INFORMES. |
| D13 | `logs/log_manager.py` huérfano | Cablearlo en F10 (bitácora en archivo con retención de 90 días que la config ya promete). Tarea opcional/P3: si el tiempo apremia, dejarlo y anotarlo en CHANGELOG como pendiente. |

---

## 3. Fases de ejecución

> Convención de estimación: S = <1 h, M = 1–3 h, L = 3–6 h de trabajo de agente.

---

### F1 · Fundaciones compartidas (M)

**Objetivo:** crear los módulos comunes de los que dependen todas las fases siguientes, sin cambiar aún ningún comportamiento observable.

**Archivos:** `motor/columnas_comunes.py` (nuevo), `motor/composicion.py` (nuevo), `motor/utilidades.py`, `motor/mapeo_columnas.py`, `tests/test_composicion.py` (nuevo).

**Tareas:**

1. **`motor/columnas_comunes.py`** — mover aquí (y re-exportar donde haga falta para no romper imports):
   - Listas de alias como constantes: `ALIAS_PROGRAMA` (DERIVACION, DERIVACIÓN, PROGRAMA, NOMBRE CENTRO), `ALIAS_TRIBUNAL`, `ALIAS_NOMBRE` (NOMBRE, NOMBRE COMPLETO, NOMBRE MENOR), `ALIAS_RUT` (**unión** de las listas de motor y correos: RUT, RUT MENOR, RUT NNA, RUT LITIGANTE — mejora 1.8/1.9), `ALIAS_RIT` (solo "RIT" — mejora 2.5 rechazó ampliar), `ALIAS_VENCIMIENTO`, `ALIAS_ESPERA`.
   - Funciones compartidas motor↔correos: `normalizar`, `detectar_tribunal`, `titulo_programa`, `es_dce`, `obtener_col(df, aliases)`. Pueden quedar implementadas en `utilidades.py` y re-exportadas aquí, o movidas aquí — a elección del ejecutor, pero **una sola implementación** de cada una.
   - `TRIBUNAL_DISPLAY = {"LAJA": "Laja", "MULCHEN": "Mulchén", "TOME": "Tomé"}` (para mejora 1.10).
   - **Nota de alcance:** `resoluciones/generador_resoluciones.py` NO se toca — conserva sus copias (mejora 2.3).
2. **Helpers nuevos en `utilidades.py`:**
   - `contiene_token(texto, *tokens) -> bool`: normaliza y compara por **palabra completa** (`re.search(r'(?<![a-z])fae(?![a-z])', norm)` o split por tokens). Corrige H-03. `es_dce()` puede mantener su criterio actual (subcadena "DCE"/"DIAGNOSTICO") — el catálogo no lo objetó — pero documentar la diferencia.
   - `es_derivacion_sin_seg(programa)` **reescrita** según E-01: lista completa `OPD, DAM, SALUD PRIVADA, HOSPITAL, UNIDAD DE SALUD MENTAL, CESFAM, RED SALUD, CONSULTA EXTERNA, COLEGIO, CHILE CRECE CONTIGO`, con match de **palabra completa anclada al inicio** del nombre normalizado (regex `^(?:opd|dam|salud privada|...)(?:\b|$)` sobre `normalizar(programa)`). Sin exclusión DCE (rechazada en §11).
3. **`motor/composicion.py`:**
   ```python
   def prefijo(nombre, programa) -> str          # mueve/envuelve prefijo_observacion (G-01, sin cambios)
   def componer(pfx, fragmentos) -> str          # G-04:
       # - descarta fragmentos vacíos/None
       # - garantiza que cada fragmento termine en exactamente un "."
       # - une con " " (los puntos de cada fragmento producen ". " entre frases)
       # - colapsa ".." accidentales y espacios dobles
       # - antepone pfx; si no hay fragmentos retorna ""
   class Incidencias:                            # D6 / G-05 / G-06
       def agregar(self, fila_excel, rit, regla, motivo)
       def como_dataframe(self) -> pd.DataFrame  # columnas FILA_EXCEL, RIT, REGLA, MOTIVO
   def fecha_valida(valor) -> datetime | None    # get_date + descarta None; las reglas SOLO
                                                 # renderizan fechas que pasaron por aquí (H-04)
   ```
4. `mapeo_columnas.py`: consumir los alias desde `columnas_comunes`. **Todavía no** eliminar `ficha_res`/`ficha_amb` (eso es F4, junto con las reglas que los usan).
5. Tests (`tests/test_composicion.py`): `componer` con 0/1/3 fragmentos, fragmento sin punto final, fragmento con punto, sin dobles puntos, prefijo con/sin nombre/sigla; `contiene_token("Residencia San Rafael", "fae") is False`; `contiene_token("FAE Familia Sur", "fae") is True`; `es_derivacion_sin_seg`: "OPD Laja" → True, "PRODAM X" → False, "Chile Crece Contigo B" → True, "DCE Tomé" → False, "Red Salud Bío Bío" → True.

**Aceptación:** suite completa previa sigue verde (73 passed) + tests nuevos verdes. Ningún cambio de comportamiento del motor todavía.

---

### F2 · Textos v9 (JSON + goldens) (M)

**Objetivo:** cargar la fuente única de verdad con los 25 textos aprobados del catálogo.

**Archivos:** `motor/textos_observaciones.json`, `tests/generar_goldens.py`, `tests/goldens_textos.json`, `tests/test_paridad.py`.

**Tareas:**

1. Reemplazar el contenido de `textos_observaciones.json` por la estructura del **Apéndice A** (4 secciones: `COMUN`, `ESPERA`, `CUMPLIMIENTO`, `INFORMES`; 25 entradas; todas `"confirmado": true`; conservar `_meta` actualizando su descripción). Los IDs antiguos (R0…R10, FALLBACK…) desaparecen.
2. `motor/textos.py` no requiere cambios de código (`render("COMUN", "CURADOR", ...)` ya funciona por diseño).
3. Ajustar `tests/test_paridad.py`:
   - `test_conteo_total_reglas_es_31` → renombrar y exigir **25**.
   - Guard de pendientes: ahora `assert len(pendientes) == 0` (todos los textos v9 vienen aprobados por el catálogo; si en el futuro se agrega una regla sin confirmar, debe fallar).
4. Ajustar `tests/generar_goldens.py` si asume IDs/modos antiguos (usa placeholders genéricos; verificar que el set de `placeholders_usados` cubra los nuevos: `PNOMBRE, PROGRAMA, FECHA_MAYORIA, FECHA_RESOLUCION, FECHA_INGRESO, FECHA_EGRESO_PROYECTADO, FECHA_FICHA_INDIVIDUAL, FECHA_VENCIMIENTO, FECHA_OIDO, FECHA_AUDIENCIA`).
5. Regenerar: `python tests/generar_goldens.py` y commitear el nuevo `goldens_textos.json`.

**Aceptación:** `pytest tests/test_paridad.py -v` verde con 25 reglas; `python -c "from motor.textos import listar_no_confirmados; assert not listar_no_confirmados()"`.
**Nota:** `tests/test_regresion.py` (smoke) puede quedar rojo transitoriamente porque las reglas aún usan IDs viejos → por eso F2 y F3–F5 pueden agruparse en una sola rama de trabajo; commits separados, pero correr la suite completa recién al final de F5 es aceptable. Indicarlo en el mensaje de commit.

---

### F3 · Motor de reglas ESPERA (L)

**Objetivo:** reescribir `generar_observacion_espera` según E-01…E-06 + T-01…T-03, usando el composer.

**Archivos:** `motor/reglas_espera.py`, `tests/test_reglas_espera.py` (nuevo).

**Firma nueva:**
```python
def generar_observacion_espera(row, tribunal, cols, incidencias=None, fila_excel=None) -> str
# tribunal: "LAJA" | "MULCHEN" | "TOME" | None  (None = no reconocido, SIN default)
```

**Lógica (orden §9.1 del catálogo — seguir el pseudocódigo del Apéndice B.1):**

1. **E-01 corte:** si `es_derivacion_sin_seg(programa)` → retornar `pfx + render("COMUN","NO_SEGUIMIENTO", PROGRAMA=titulo_programa(programa))` **sin** audiencia ni nada más (H-06).
2. **E-02 mayoría (principal):** edad exacta ≥18 **solo** desde FECHA DE NACIMIENTO (eliminar el fallback por columna EDAD y el texto `R1_FALLBACK` — catálogo E-02: el dato nunca falta y no hay texto alternativo). Si dispara: fragmentos = [E-02] + T-01 + T-02 + T-03 y componer (ya no hay return temprano que pierda curador/oído — H-09).
3. Si no hubo E-02: fragmentos en orden = **E-03** (1 ≤ días_para_mayoría ≤ 60) → **E-04** (fecha resolución válida y 0 ≤ días transcurridos ≤ 29) → **E-05** (T ESPERA ≥ 30, ramas del catálogo; requiere `tribunal` salvo rama DCE; si tribunal es None y no es DCE → omitir E-05 + incidencia G-06) → **E-06** si no disparó ni E-04 ni E-05 → **T-01** → **T-02** (0 ≤ días ≤ 45, ahora incluye día 0) → **T-03** (fecha ≥ hoy).
4. E-05 ramas: DCE (cualquier tribunal, ≥30) → `ESPERA.E05_SOLO_CORREO`; Laja/Mulchén ≥30 → `ESPERA.E05_PROYECTO_Y_CORREO`; Tomé 30–59 → `E05_SOLO_CORREO`; Tomé ≥60 → `E05_PROYECTO_Y_CORREO`.
5. T-03 deja de ser "sufijo" (`audiencia_suffix` con recortes de punto): es un fragmento normal `render("COMUN","PROX_AUDIENCIA", FECHA_AUDIENCIA=fecha_es(f))` que entra al composer. Eliminar `audiencia_suffix` cuando ya nadie lo use (fin de F5).
6. G-05 en cada regla: si el dato indispensable falta/es inválido (p. ej. FEC. RESOLUCIÓN ilegible para E-04), se omite **solo esa regla** y se registra incidencia.

**Tests mínimos (ver matriz Apéndice C.1):** límites 29/30 Laja y Mulchén; 29/30/59/60 Tomé; DCE 29/30 y verificación explícita de que el texto DCE **no contiene** "proyecto de resolución"; E-04 días 0/29/30; E-04+E-05 simultáneas (ambos fragmentos presentes, en ese orden); E-06 fallback (espera 10 sin resolución; espera 40 con tribunal None no-DCE → E-06 + incidencia); E-02 con curador sin RUT → ambos fragmentos; E-03 límites 0/1/60/61; T-02 día 0 y 46; T-03 con E-01 → ausente; prefijo `Camila PIE: `; sin `{`, sin `..`.

**Aceptación:** `pytest tests/test_reglas_espera.py tests/test_composicion.py -v` verde. Los ejemplos 10.1, 10.2 y 10.3 del catálogo reproducidos **byte a byte** como tests (armar filas sintéticas que los generen).

---

### F4 · Motor de reglas CUMPLIMIENTO (L)

**Objetivo:** reescribir cumplimiento según C-01…C-10; el cruce Hoja2 pasa a ser una regla dentro del motor (corrige H-07); eliminar C-06.

**Archivos:** `motor/reglas_cumplimiento.py`, `motor/utilidades.py` (podar), `motor/mapeo_columnas.py` (podar), `motor/procesador.py` (solo la parte del cruce), `tests/test_reglas_cumplimiento.py` (nuevo).

**Firma nueva:**
```python
def generar_observacion_cumplimiento(row, tribunal, cols, fecha_hoja2=None,
                                     incidencias=None, fila_excel=None) -> str
# fecha_hoja2: date | None — fecha de vencimiento si la fila matcheó en el índice Hoja2
```
El procesador calcula la clave y busca en el índice (como hoy) pero **pasa la fecha al motor** en vez de reemplazar la observación después.

**Lógica (orden §9.2 — pseudocódigo Apéndice B.2):**

1. **E-01 corte** (idéntico a espera, texto COMUN, sin audiencia).
2. **C-01 mayoría:** igual a E-02; acumula T-01/T-02/T-03 (D4).
3. Fragmentos en orden: **C-02** (próx. mayoría 1–60) → **C-03** (condición: columna `DIAS DE CUMPLIMIENTO` con `0 ≤ valor ≤ 30`; la fecha de ingreso es solo para el texto — si la fecha falta, omitir + incidencia G-05) → **C-04** (`dias_para_egresar < 0` **o** `dias_cumplimiento < 0`, **y** fecha de egreso proyectado válida; sin fecha → omitir + incidencia — corrige H-04) → **C-05** (fecha válida y `0 ≤ dias_para_egresar ≤ 45`; día 0 usa `C05_VENCE_HOY`, 1–45 usa `C05_POR_VENCER`) → **C-10** (si `fecha_hoja2` y **no** dispararon C-04 ni C-05) → **T-01** → **T-02** → **C-07** → **C-08** → **T-03**.
4. **C-07 fusionada:** aplica si `programa_norm.startswith(("rta","rtt","res","rfa","rva"))` (D7 — sale PEE/RPPM, entra RVA) **y** existe la columna `FEC.ACT.F.INDIVIDUAL`. Tres ramas: sin fecha → `C07_SIN_FICHA`; > 180 días → `C07_FICHA_ANTIGUA`; 0–30 días → `C07_FICHA_RECIENTE`; 31–180 → nada. Eliminar por completo la regla de `ficha_res` (columna inexistente en los Excel reales) y su mapeo.
5. **C-08:** `contiene_token(programa, "fae", "fas")` (corrige H-03) + columna de ficha FAE existe + >120 días desde ingreso efectivo + sin fecha de ficha → `C08_FICHA_FAE`.
6. **C-09:** si ningún principal disparó (principales = C-01, C-03, C-04, C-05, C-10): sin fragmentos → `C09_SIN_OBSERVACIONES`; solo complementarias (C-02, T-01, T-02, C-07, C-08, T-03 — D3) → `C09_BASE_BREVE` como primer fragmento + las complementarias.
7. Actualizar los helpers de supresión usados por el procesador a los umbrales v9 y renombrarlos: `medida_vencida_para_fila` → `dias < 0 or dias_cumpl < 0` (con fecha válida), `c05_aplica_para_fila` → `0 ≤ dias ≤ 45` (con fecha válida). Mejor aún: eliminar los helpers y decidir la supresión **dentro** del motor (ya tiene todo el contexto) — preferido.
8. **Podas:** `proximo_informe()` + import `relativedelta` en `utilidades.py`; claves `ficha_res` y `ficha_amb` en `mapeo_columnas.py`; `_PREFIJOS_RESIDENCIAL` viejo; `r5_aplica_para_fila`/`medida_vencida_para_fila` si se optó por decidir dentro del motor.

**Tests mínimos (Apéndice C.2):** C-03 con DIAS DE CUMPLIMIENTO 0/30/31 (y verificación de que **no** se usa el delta de fechas: fila con días=15 pero fecha de ingreso de hace 90 días → dispara igual); C-04 con -1/0 y con dias_cumpl -1; C-04 sin fecha egreso → omitida + incidencia (y **no** aparece la fecha de hoy en el texto); C-05 día 0 texto "para el día de hoy"; C-10 suprimida con C-04, suprimida con C-05, presente sola, conviviendo con C-03 (ambas presentes); C-07 tres ramas + límites 30/31/180/181 + PEE ya no aplica + RVA aplica; C-08 con "San Rafael" NO aplica; C-09 ambas variantes; ejemplo 10.4 y 10.5 del catálogo byte a byte.

**Aceptación:** tests de la fase verdes; `grep -n "proximo_informe\|relativedelta\|ficha_res" motor/*.py` vacío.

---

### F5 · Motor de reglas INFORMES (M)

**Objetivo:** corte E-01 (H-08) + límites I-01/I-02 + variante DCE en vencidos.

**Archivos:** `motor/reglas_informes.py`, `tests/test_reglas_informes.py` (nuevo).

**Lógica (Apéndice B.3):**

1. E-01 corte → texto COMUN, sin audiencia.
2. Sin fecha de vencimiento válida → `""` + incidencia G-05 (tipo de retorno unificado a `str` — H-15).
3. `dias = (fecha_venc - hoy).days`: `dias < 0` → I-01 (DCE → `I01_VENCIDO_DCE`, si no → `I01_VENCIDO_GENERAL`); `0 ≤ dias ≤ 30` → I-02 (DCE/general); `> 30` → `""`.
4. T-03 audiencia como fragmento si hay observación principal (D10). La detección DCE usa `es_dce()` (no el `"DCE" in upper` local de hoy).

**Tests (C.3):** días -1/0/30/31; DCE vencido usa "informe diagnóstico"; OPD en INFORMES → texto no seguimiento (nuevo comportamiento); sin fecha → `""` + incidencia; ejemplo 10.6 byte a byte.

**Aceptación:** tests de fase verdes + **suite completa verde** (aquí terminan los cambios de reglas; correr `pytest tests/ -v` y arreglar cualquier residuo de F2–F5, regenerando fixtures sintéticos si el smoke los necesita ya — ver F11.1, puede adelantarse).

---

### F6 · Procesador: salida §8, validación, sin default LAJA (M)

**Objetivo:** columnas de salida nuevas, hoja VALIDACION, eliminación del default de tribunal, desduplicación de la carga.

**Archivos:** `motor/procesador.py`, `tests/test_salida_excel.py` (nuevo).

**Tareas:**

1. Eliminar `or "LAJA"` en `_calcular_simple` y `_calcular_cumplimiento` (H-01): pasar `tribunal=None` al motor cuando no se reconoce (las reglas ya lo manejan desde F3/F4) y registrar incidencia G-06 por fila.
2. `Incidencias` viaja por todo el cálculo: `_calcular_simple`/`_calcular_cumplimiento` retornan `(df_r, incidencias)`. `calcular_preview` ignora las incidencias para la ventana (o las muestra como contador — opcional) pero **debe seguir siendo byte-idéntico** en OBSERVACION.
3. Columnas de salida (D1/D2/D12) en `_calcular_*`: al inicio, dropear si existen `["OBSERVACION","FECHA_OBS","TT","CC","RES"]`; al final, insertar `FECHA_OBS` (valor `datetime.now().strftime("%d/%m/%Y")`) inmediatamente **antes** de OBSERVACION y `TT`, `CC`, `RES` (vacías) inmediatamente **después**.
4. `_guardar_excel`: escribir la hoja `VALIDACION` (segunda hoja) cuando haya incidencias, con columnas `FILA_EXCEL, RIT, REGLA, MOTIVO`; y loguear `⚠️ N incidencias de validación — ver hoja VALIDACION`.
5. Extraer `_cargar_libro_cumplimiento(path) -> (df_h1, df_h2, log_msgs)` y usarla desde `procesar()` **y** `calcular_preview()` (H-11).
6. Aviso Hoja2 (mejora 2.2, lado backend): cuando el modo es CUMPLIMIENTO y `df_h2 is None` **o** el índice queda vacío, además del log, encolar un mensaje nuevo `("warn_hoja2", texto)` para que la GUI lo muestre prominente (F10). Texto sugerido: `"Sin Hoja2 utilizable: HOY no se generará NINGUNA observación de próximo informe (C-10). Verifica el archivo."`.

**Tests (`tests/test_salida_excel.py`):** usar `runner.ejecutar_modo_aislado` o llamar `_calcular_*` directo con DataFrames sintéticos: orden de columnas `... FECHA_OBS, OBSERVACION, TT, CC, RES ...`; `TT/CC/RES` vacías; `FECHA_OBS` = hoy dd/mm/yyyy; reproceso de un df que ya traía OBSERVACION/TT no duplica columnas; fila con tribunal "JUZGADO DE COYHAIQUE" espera 40 no-DCE → OBSERVACION con E-06 y una incidencia G-06; hoja VALIDACION presente en el xlsx guardado cuando hay incidencias.

**Aceptación:** tests de fase + suite completa verdes.

---

### F7 · Correos (L)

**Objetivo:** ítems 1.1, 1.2, 1.3, 1.6, 1.7, 1.9, 1.10 de las mejoras + testeabilidad sin Outlook.

**Archivos:** `comunicaciones/generador_correos.py`, `comunicaciones/plantillas/*.html`, `gui/app.py` (solo lo tocante a contactos), `contactos.json` (eliminar), `tests/test_correos.py` (nuevo).

**Tareas:**

1. **1.1 — Eliminar correo a tribunal:** borrar `_trib_borrador`, su llamada en `procesar()`, `plantillas/correo_tribunal_vencidos.html`, el parámetro `contactos_tribunales` del constructor, `contactos.json`, y en `gui/app.py`: `CONTACTOS_FILE`, `DEFAULT_CONTACTOS`, `self.contactos` y su paso al generador.
2. **1.2 —** `PATRON_ESPERA = "se remite correo electronico al programa consultando respecto de la fecha estimada de ingreso efectivo"` (coincide con las tres ramas E-05 v9 tras `_normalizar_texto`).
3. **1.9 —** borrar `_formatear_nombre_programa`, `_normalizar_texto` propio y `_detectar_clave_tribunal`; usar `titulo_programa`, `normalizar`, `detectar_tribunal`, `es_dce` y los `ALIAS_*` de `motor/columnas_comunes.py` (también en `_col`).
4. **1.6/1.7 —** `DIAS_POR_VENCER = 30`; clasificación `d < 0` → VENCIDO, `0 ≤ d ≤ 30` → POR_VENCER.
5. **1.3/D8 —** plantillas `correo_programa_vencidos.html` y `correo_programa_por_vencer.html`: reemplazar el literal "informes de avance" por `{{ETIQUETA_INFORMES}}`; el generador pasa `"informes diagnósticos"` si `es_dce(programa)`, si no `"informes de avance"`; mismo criterio en el asunto.
6. **1.10/D9 —** agrupar por `(programa, clave_tribunal_normalizada)` y mostrar `TRIBUNAL_DISPLAY[clave]` en asunto y cuerpo; clave no reconocida → agrupar por crudo + advertencia en `resultado["errores"]`.
7. **Testeabilidad:** separar armado de envío — `procesar()`/`procesar_espera()` construyen una lista de borradores `[{"para", "cc", "asunto", "cuerpo_html", "n_registros"}, ...]` y un paso final los entrega a `self._despachar(borrador)` (por defecto `_crear_borrador_outlook`, inyectable en el constructor para tests: `GeneradorCorreos(..., despachador=fake)`).

**Tests (`tests/test_correos.py`, con despachador fake, sin win32com):** día -1 → vencido, 0 → por vencer, 30 → por vencer, 31 → sin correo; grupo DCE usa "informes diagnósticos" en asunto y cuerpo; grupo no-DCE usa "informes de avance"; asunto contiene "Mulchén" aunque la celda diga "JUZGADO DE LETRAS Y GARANTIA DE MULCHEN"; dos variantes de escritura del mismo tribunal caen en **un** borrador; `procesar_espera` matchea filas cuya OBSERVACION viene de E05_SOLO_CORREO y E05_PROYECTO_Y_CORREO; sin observación de espera → error informativo; ya no existe tipo B (buscar "TRIBUNAL" en detalle → ausente).

**Aceptación:** tests de fase + suite verdes; `grep -rn "correo_tribunal_vencidos\|contactos_tribunales\|_trib_borrador" .` sin resultados en código (solo en docs).

---

### F8 · Resoluciones — parche `detectar_tipo()` (S)

**Objetivo:** ítem 2.1 de las mejoras, exactamente como está especificado. **Nada más de este módulo se toca** (2.3 rechazada).

**Archivos:** `resoluciones/generador_resoluciones.py`, `tests/test_resoluciones.py` (nuevo).

**Tareas:**

1. Reemplazar patrones:
   ```python
   _PATRON_PC_IE   = ["proyecto de resolucion pidiendo cuenta al programa respecto del ingreso efectivo"]
   _PATRON_PC_INFO = ["que se encuentra vencido en rus desde el"]
   _PATRON_NOMENCL = [...]  # SIN CAMBIOS — dormant
   ```
2. Ojo con el orden de chequeo existente (PC_INFO primero): mantenerlo.

**Tests:** `detectar_tipo(render E05_PROYECTO_Y_CORREO con prefijo) == "PC_IE"`; `detectar_tipo(E05_SOLO_CORREO) is None` (DCE/Tomé-bajo jamás genera resolución); `detectar_tipo(I01_VENCIDO_GENERAL) == "PC_INFO"`; `detectar_tipo(I01_VENCIDO_DCE) == "PC_INFO"`  ← **verificar contra el flujo real**: los proyectos por informe vencido se generan desde el Excel de INFORMES; `detectar_tipo(C04_VENCIDA texto) is None` (dice "vencida", no "vencido… desde el" — no debe colisionar; test explícito); `detectar_tipo("aplica nomenclaturas") == "NOMENCL"` (dormant intacto).

**Aceptación:** tests verdes. Documentar en CHANGELOG que Tomé sin plantilla sigue cayendo a "faltantes" (nota de riesgo 6 de la auditoría).

---

### F9 · Validador: cablear + podar + optimizar (M)

**Objetivo:** que el validador por fin corra (H-02), sin su código muerto (2.6), con costo razonable (H-13).

**Archivos:** `validador/reglas_validacion.py`, `validador/precheck.py`, `motor/procesador.py`, `gui/app.py` (mensajes), `tests/test_validador.py` (nuevo).

**Tareas:**

1. Eliminar `COLS_MINIMAS` y `COLS_CRITICAS_VACIAS` (2.6).
2. Optimizar A4: resolver `_col_match` **fuera** de los loops; vectorizar A4–A8 donde sea directo (p. ej. `df[col].map(get_date)` + máscaras).
3. Cablear: en `procesar()` (los 3 modos), tras cargar el DataFrame y antes de calcular, llamar `validar_excel(df, modo, ruta_salida_reportes=config["ruta_salida_excel"])`:
   - `puede_procesar == False` → log de cada bloqueante + `q.put(("done", "❌ {modo} cancelado por validación: ..."))` y abortar.
   - Advertencias → una línea de log por regla (`⚠️ A8: 3 filas con tribunal no reconocido`) + ruta del HTML si se generó.
   - `calcular_preview` NO ejecuta el precheck (la vista previa debe seguir siendo instantánea y sin archivos).
4. Alinear alias del validador con `columnas_comunes` (importar las mismas listas).

**Tests:** df sin columna TRIBUNAL → bloquea; df con RIT duplicado (mismo RIT+NNA+programa) → bloquea; tribunal desconocido → advierte A8 con la fila correcta; df limpio → `puede_procesar=True` sin advertencias.

**Aceptación:** tests verdes; correr `runner.ejecutar_modo_aislado` sobre los fixtures y verificar que el flujo completo sigue OK con el precheck activado.

---

### F10 · GUI (M)

**Objetivo:** mejoras 2.2 y 2.10 + hallazgos H-05, H-12, H-14, D13.

**Archivos:** `gui/app.py`, `motor/version.py` (nuevo), `main.py`.

**Tareas:**

1. **H-05 — thread-safety:** los workers de Correos y Resoluciones dejan de tocar widgets: encolan `("log_correos", msg)` / `("log_res", msg)` / `("status", ...)` / `("done_correos", resumen)` etc., y `_process_queue` (hilo principal) escribe en el widget que corresponda y muestra los `messagebox`. Patrón idéntico al de la pestaña Motor.
2. **2.2 — aviso Hoja2:** manejar el mensaje `("warn_hoja2", texto)` de F6: `messagebox.showwarning("Hoja2 ausente", texto)` + línea en bitácora con tag de warning. Solo en CUMPLIMIENTO.
3. **2.10 — dos campos de entrada en Correos:** `self.correo_informes_var` (Excel de INFORMES) y `self.correo_espera_var` (Excel de ESPERA **procesado**, salida del motor), cada uno con su botón "Buscar…" junto al botón de acción correspondiente. Cada acción valida su propio campo.
4. **H-14:** `_start_preview` marca/verifica `self.running` igual que `_start_motor`.
5. **H-12 — versión:** crear `motor/version.py` con `VERSION = "v9.0.0"`; `gui/app.py` la importa; actualizar título de ventana y barra de estado.
6. **D13 (opcional/P3):** en el arranque de `RUSApp`, `get_logger(self.cfg["ruta_logs"])` y duplicar a archivo lo que entra por `("log", ...)`. Si se omite, registrar como pendiente en CHANGELOG.

**Aceptación:** `python -m py_compile gui/app.py`; revisión manual del diff confirmando que **ningún** `messagebox`/`.set(`/`_write(` queda dentro de funciones `_worker_*`; suite verde.

---

### F11 · Fixtures, regresión y cierre de pruebas (M)

**Objetivo:** que la cobertura declarada en el catálogo §13 exista de verdad y el smoke sintético ejercite las ramas v9.

**Archivos:** `tests/fixtures/generar_fixtures.py`, `tests/fixtures/*.xlsx` (regenerar), `tests/test_regresion.py`.

**Tareas:**

1. Reescribir los datos sintéticos para cubrir las ramas v9: Espera (Laja 29/30, Mulchén 30, Tomé 29/30/59/60, DCE 30 en Tomé y en Laja, E-04 reciente, E-04+E-05, mayoría, próxima mayoría, OPD, tribunal desconocido, curador institucional, oído día 0); Cumplimiento (dias_cumpl 0/30/31, vencida -1, vence hoy 0, por vencer 45, Hoja2 match con y sin C-04/C-05, RTA sin ficha, RES ficha 181d, RVA ficha 15d, PEE (ya no aplica), FAE sin ficha 121d, "Residencia San Rafael", fila limpia → sin observaciones); Informes (vencido -1, día 0, 30, 31, DCE vencido, OPD). Ejecutar `python tests/fixtures/generar_fixtures.py` y commitear los .xlsx.
2. `tests/test_regresion.py`: mantener el mecanismo de env vars para los Excel reales; en el modo smoke agregar invariantes nuevos: columnas `FECHA_OBS/TT/CC/RES` presentes y en orden, sin `{` en OBSERVACION, `TT/CC/RES` vacías. Actualizar el comentario de `CONTEOS_ESPERADOS` aclarando que los conteos 16/100/333 son pre-v9 y deben re-medirse en la máquina del usuario.
3. Verificación cruzada final de la matriz del Apéndice C: cada casilla debe apuntar a un test existente (agregar los que falten).

**Aceptación:** `pytest tests/ -v` — todo verde; conteo de tests significativamente mayor al baseline (referencia: ≥120).

---

### F12 · Documentación y release (S)

**Archivos:** `README.md`, `CHANGELOG.md`, `requirements.txt`, `PROMPT_CONTINUACION.md`.

**Tareas:**

1. `requirements.txt` (H-10):
   ```
   pandas>=1.5.0
   openpyxl>=3.1.0
   xlrd>=2.0.1
   python-docx>=0.8.11
   python-dateutil>=2.8.2
   pywin32>=305; sys_platform == 'win32'
   ```
   (+ nota de dev: `pip install pytest`).
2. README: actualizar secciones de reglas (resumen E/C/I v9 con referencia al catálogo), columnas de salida `FECHA_OBS/TT/CC/RES` (§8 y su aritmética TT−CC), eliminación del correo a tribunal, dos campos en la pestaña Correos, validador activo.
3. CHANGELOG: entrada `v9.0.0` enumerando: catálogo v2 completo, 14 mejoras aceptadas, hallazgos H-01…H-15 corregidos, y los pendientes conocidos (plantilla Word PC_IE de Tomé inexistente → faltantes; herramienta de conteo trimestral TT/CC agendada, catálogo §8; D13 si quedó fuera).
4. `PROMPT_CONTINUACION.md`: anteponer un aviso de 3 líneas: "HISTÓRICO v8.14 — decisiones superadas por docs/especificaciones/…; no usar como fuente de verdad".
5. Verificar la **Definición de HECHO global** (sección 0) punto por punto. Build Windows: mismas instrucciones PyInstaller del PROMPT_CONTINUACION (sin cambios; queda en manos del usuario).

---

## 4. Matriz de trazabilidad

### 4.1 Catálogo v2 → fases

| Regla/ítem | Fase | | Regla/ítem | Fase |
|---|---|---|---|---|
| E-01 (lista+anclaje) | F1, F3 | | C-04, C-05 (día 0, fecha válida) | F4 |
| E-02/E-03 | F3 | | C-06 eliminada | F4 |
| E-04 | F3 | | C-07 fusionada / C-08 FAE | F4 |
| E-05 rediseñada | F2 (textos), F3 | | C-09 cierres | F4 |
| E-06 fallback | F3 | | C-10 Hoja2 (compone+suprime) | F4, F6 |
| T-01/T-02/T-03 | F1, F3, F4, F5 | | I-01/I-02 (+corte E-01) | F5 |
| G-01…G-05 (composer) | F1 | | §8 columnas salida | F6 |
| G-06 sin default | F6 | | §13 criterios de prueba | F3–F5, F11 (Apéndice C) |
| G-07/G-08 | F4 | | Textos aprobados | F2 (Apéndice A) |

### 4.2 Mejoras → fases: 1.1→F7 · 1.2→F7 · 1.3→F7 · 1.4/1.8/1.9→F1+F7 · 1.6→F7 · 1.7→F7 · 1.10→F7 · 2.1→F8 · 2.2→F6+F10 · 2.6→F9 · 2.10→F10 · 2.12→sin cambio · 2.13→F4/F6 (conducta confirmada).

### 4.3 Auditoría → fases: H-01→F6 · H-02→F9(+F10 D13) · H-03→F1/F4 · H-04→F1/F4 · H-05→F10 · H-06→F3/F4 · H-07→F4 · H-08→F5 · H-09→F3/F4 · H-10→F12 · H-11→F6 · H-12→F10 · H-13→F9 · H-14→F10 · H-15→F5 · H-16→hecho · H-17→F1/F3.

---

## Apéndice A · `motor/textos_observaciones.json` v9 — contenido definitivo

Los textos van **sin** el prefijo `{PRIMER_NOMBRE} {SIGLA}: ` (el composer lo antepone — G-01). Placeholders en ASCII (sin tildes). Todos con `"confirmado": true` y una `"nota": "Catálogo v2 2026-07-14 — <ID catálogo>"`.

**Sección `COMUN`** (compartidos por los tres modos):

| ID JSON | ID catálogo | Texto exacto |
|---|---|---|
| `NO_SEGUIMIENTO` | E-01 | `La derivación {PROGRAMA} no se encuentra sujeta a seguimiento por este Centro.` |
| `MAYORIA_EDAD` | E-02 / C-01 | `Se hace presente que {PNOMBRE} alcanzó la mayoría de edad el {FECHA_MAYORIA}, se sugiere egresar la medida.` |
| `PROXIMA_MAYORIA` | E-03 / C-02 | `Se hace presente que {PNOMBRE} alcanzará la mayoría de edad el {FECHA_MAYORIA}.` |
| `CURADOR` | T-01 | `No registra curador asociado en RUS, se sugiere asociar curador ad litem informáticamente.` |
| `OIDO` | T-02 | `Oído con fecha {FECHA_OIDO}.` |
| `PROX_AUDIENCIA` | T-03 | `Se cita a audiencia para el día {FECHA_AUDIENCIA}.` |

**Sección `ESPERA`:**

| ID JSON | ID catálogo | Texto exacto |
|---|---|---|
| `E04_RESOLUCION_RECIENTE` | E-04 | `Medida revisada, a la espera de ingreso efectivo. Se hace presente que el Tribunal ordenó el ingreso efectivo al programa {PROGRAMA} con fecha {FECHA_RESOLUCION}. (Observación administrativa, no requiere acción/respuesta del Tribunal).` |
| `E05_SOLO_CORREO` | E-05 (DCE; Tomé 30–59) | `Medida revisada, a la espera de ingreso efectivo. Se remite correo electrónico al programa consultando respecto de la fecha estimada de ingreso efectivo.` |
| `E05_PROYECTO_Y_CORREO` | E-05 (Laja/Mulchén ≥30; Tomé ≥60) | `Medida revisada, a la espera de ingreso efectivo. Se remite proyecto de resolución pidiendo cuenta al programa respecto del ingreso efectivo. Igualmente, se remite correo electrónico al programa consultando respecto de la fecha estimada de ingreso efectivo.` |
| `E06_SIN_RESOLUCION` | E-06 | `Medida revisada, a la espera de ingreso efectivo al programa {PROGRAMA}.` |

**Sección `CUMPLIMIENTO`:**

| ID JSON | ID catálogo | Texto exacto |
|---|---|---|
| `C03_INGRESO_RECIENTE` | C-03 | `Medida revisada. Se hace presente que el ingreso efectivo al programa {PROGRAMA} se registra con fecha {FECHA_INGRESO}.` |
| `C04_VENCIDA` | C-04 | `La medida se visualiza vencida en RUS desde el {FECHA_EGRESO_PROYECTADO}.` |
| `C05_VENCE_HOY` | C-05 día 0 | `Se hace presente que la medida se visualiza con vencimiento para el día de hoy, {FECHA_EGRESO_PROYECTADO}.` |
| `C05_POR_VENCER` | C-05 días 1–45 | `Se hace presente que la medida se visualiza próxima a vencer en RUS el {FECHA_EGRESO_PROYECTADO}.` |
| `C07_SIN_FICHA` | C-07 sin fecha | `No registra ficha individual en RUS, se sugiere confeccionar.` |
| `C07_FICHA_ANTIGUA` | C-07 >180d | `Atendido que la ficha individual registra como última actualización el {FECHA_FICHA_INDIVIDUAL}, superando los 180 días, se sugiere actualizar.` |
| `C07_FICHA_RECIENTE` | C-07 0–30d | `Se hace presente que la ficha individual fue actualizada con fecha {FECHA_FICHA_INDIVIDUAL}.` |
| `C08_FICHA_FAE` | C-08 | `Se hace presente que {PNOMBRE} no tiene ficha FAE, se sugiere confeccionar.` |
| `C09_SIN_OBSERVACIONES` | C-09 | `Medida revisada, sin observaciones.` |
| `C09_BASE_BREVE` | C-09 | `Medida revisada.` |
| `C10_HOJA2` | C-10 | `Medida revisada, se hace presente que el programa {PROGRAMA} deberá remitir informe de avance a más tardar el {FECHA_VENCIMIENTO}.` |

**Sección `INFORMES`:**

| ID JSON | ID catálogo | Texto exacto |
|---|---|---|
| `I01_VENCIDO_DCE` | I-01 DCE | `Se remite correo electrónico al programa {PROGRAMA} a fin de requerir el informe diagnóstico que se encuentra vencido en RUS desde el {FECHA_VENCIMIENTO}.` |
| `I01_VENCIDO_GENERAL` | I-01 otros | `Se remite correo electrónico al programa {PROGRAMA} a fin de requerir el informe de avance que se encuentra vencido en RUS desde el {FECHA_VENCIMIENTO}.` |
| `I02_POR_VENCER_DCE` | I-02 DCE | `Se remite correo electrónico al programa {PROGRAMA} a fin de señalar que el informe diagnóstico ordenado en autos debe ser remitido a más tardar el {FECHA_VENCIMIENTO}.` |
| `I02_POR_VENCER_GENERAL` | I-02 otros | `Se remite correo electrónico al programa {PROGRAMA} a fin de señalar que el próximo informe de avance vence el {FECHA_VENCIMIENTO}.` |

Total: **25 textos** (6+4+11+4). Placeholders permitidos: `PNOMBRE, PROGRAMA, FECHA_MAYORIA, FECHA_RESOLUCION, FECHA_INGRESO, FECHA_EGRESO_PROYECTADO, FECHA_FICHA_INDIVIDUAL, FECHA_VENCIMIENTO, FECHA_OIDO, FECHA_AUDIENCIA`. Todas las fechas se insertan ya formateadas con `fecha_es()` (formato largo español, G-02). `PROGRAMA` se inserta ya formateado con `titulo_programa()` (G-03).

---

## Apéndice B · Pseudocódigo de los motores de reglas

### B.1 ESPERA

```
def generar_observacion_espera(row, tribunal, cols, incidencias, fila):
    programa = row[programa]; pfx = prefijo(nombre, programa)
    if es_derivacion_sin_seg(programa):
        return componer(pfx, [render(COMUN.NO_SEGUIMIENTO, PROGRAMA=titulo_programa(programa))])   # corte total

    frags = []
    fnac = fecha_valida(row[nacimiento])
    if fnac and edad_exacta(fnac) >= 18:                                   # E-02 (sin fallback EDAD)
        frags.append(render(COMUN.MAYORIA_EDAD, PNOMBRE, FECHA_MAYORIA))
        frags += complementarias_basicas(row, cols, incidencias)           # T-01, T-02, T-03
        return componer(pfx, frags)

    if fnac and 1 <= dias_para_mayoria(fnac) <= 60:                        # E-03
        frags.append(render(COMUN.PROXIMA_MAYORIA, ...))

    principal = False
    fres = fecha_valida(row[resolucion])
    if fres and 0 <= (hoy - fres).days <= 29:                              # E-04
        frags.append(render(ESPERA.E04_RESOLUCION_RECIENTE, ...)); principal = True

    espera = get_int(row[espera]) or 0
    if espera >= 30:                                                       # E-05
        if es_dce(programa):
            frags.append(render(ESPERA.E05_SOLO_CORREO)); principal = True
        elif tribunal in ("LAJA", "MULCHEN"):
            frags.append(render(ESPERA.E05_PROYECTO_Y_CORREO)); principal = True
        elif tribunal == "TOME":
            frags.append(render(E05_PROYECTO_Y_CORREO if espera >= 60 else E05_SOLO_CORREO)); principal = True
        else:
            incidencias.agregar(fila, rit, "E-05", "tribunal no reconocido — regla omitida (G-06)")

    if not principal:                                                      # E-06 (fallback §9.1)
        frags.append(render(ESPERA.E06_SIN_RESOLUCION, PROGRAMA=...))

    frags += complementarias_basicas(row, cols, incidencias)               # T-01, T-02, T-03
    return componer(pfx, frags)

def complementarias_basicas(row, cols, incidencias):
    out = []
    if cols["curador"] and not tiene_curador_real(row[curador]): out.append(render(COMUN.CURADOR))       # T-01
    foido = fecha_valida(row[oido])
    if foido and 0 <= (hoy - foido).days <= 45: out.append(render(COMUN.OIDO, FECHA_OIDO=fecha_es(foido)))  # T-02
    faud = fecha_valida(row[prox_aud])
    if faud and faud >= hoy: out.append(render(COMUN.PROX_AUDIENCIA, FECHA_AUDIENCIA=fecha_es(faud)))       # T-03
    return out
```

### B.2 CUMPLIMIENTO

```
def generar_observacion_cumplimiento(row, tribunal, cols, fecha_hoja2, incidencias, fila):
    if es_derivacion_sin_seg(programa): return componer(pfx, [NO_SEGUIMIENTO])

    if mayor_de_edad: return componer(pfx, [MAYORIA_EDAD] + complementarias_basicas(...))    # C-01 (D4)

    frags, principal = [], False
    if 1 <= dias_para_mayoria <= 60: frags.append(PROXIMA_MAYORIA)                           # C-02 (complementaria p/ C-09, D3)

    dias_cumpl = get_int(row[dias_cumpl])
    fing = fecha_valida(row[ingreso])
    if dias_cumpl is not None and 0 <= dias_cumpl <= 30:                                     # C-03 (condición por columna)
        if fing: frags.append(C03_INGRESO_RECIENTE); principal = True
        else: incidencias.agregar(..., "C-03", "sin fecha de ingreso — regla omitida (G-05)")

    fegr = fecha_valida(row[egreso_proy]); dpe = get_int(row[dias_egresar])
    vencida = fegr and ((dpe is not None and dpe < 0) or (dias_cumpl is not None and dias_cumpl < 0))
    if vencida: frags.append(C04_VENCIDA); principal = True                                  # C-04
    if (dpe is not None and dpe < 0) and not fegr: incidencias.agregar(..., "C-04", "sin egreso proyectado (G-05)")

    c05 = fegr and dpe is not None and 0 <= dpe <= 45 and not vencida
    if c05: frags.append(C05_VENCE_HOY if dpe == 0 else C05_POR_VENCER); principal = True    # C-05

    if fecha_hoja2 and not vencida and not c05:                                              # C-10 (suprime con C-04/C-05)
        frags.append(C10_HOJA2); principal = True

    frags += [T-01, T-02]                                                                    # curador, oído
    frags += fichas(row, cols)                                                               # C-07 (3 ramas), C-08 (token FAE/FAS)
    frags += [T-03]

    if not principal:                                                                        # C-09
        return componer(pfx, [C09_SIN_OBSERVACIONES]) if not frags else componer(pfx, [C09_BASE_BREVE] + frags)
    return componer(pfx, frags)
```

### B.3 INFORMES

```
def generar_observacion_informes(row, tribunal, cols, incidencias, fila):
    if es_derivacion_sin_seg(programa): return componer(pfx, [NO_SEGUIMIENTO])
    fv = fecha_valida(row[vencimiento])
    if not fv: incidencias.agregar(..., "I-01/I-02", "sin fecha de vencimiento (G-05)"); return ""
    dias = (fv - hoy).days
    if dias < 0:      texto = I01_VENCIDO_DCE if es_dce(programa) else I01_VENCIDO_GENERAL
    elif dias <= 30:  texto = I02_POR_VENCER_DCE if es_dce(programa) else I02_POR_VENCER_GENERAL
    else:             return ""
    return componer(pfx, [render(texto, PROGRAMA=titulo_programa(programa), FECHA_VENCIMIENTO=fecha_es(fv))]
                          + ([PROX_AUDIENCIA] si aplica T-03))
```

---

## Apéndice C · Matriz de pruebas de borde (catálogo §13)

Cada fila debe existir como test automatizado al cierre de F11. "días" = valor de la variable de la condición.

**C.1 ESPERA**

| Caso | Valores | Esperado |
|---|---|---|
| E-03 bordes | días a mayoría 0 / 1 / 60 / 61 | no / sí / sí / no |
| E-02 borde | 17a364d / 18a0d | no / sí (+T-01/T-02/T-03 acumulan) |
| E-04 bordes | 0 / 29 / 30 días desde resolución | sí / sí / no |
| E-05 Laja | espera 29 / 30 | no / proyecto+correo |
| E-05 Mulchén | espera 29 / 30 | no / proyecto+correo |
| E-05 Tomé | espera 29 / 30 / 59 / 60 | no / correo / correo / proyecto+correo |
| E-05 DCE (Laja y Tomé) | espera 29 / 30 | no / correo; texto **sin** "proyecto de resolución" |
| E-04 + E-05 | resolución 10d + espera 35 Laja | ambos fragmentos, orden E-04→E-05 |
| E-06 | espera 10 sin resolución | texto E-06 |
| E-06 fallback D11 | espera 40, tribunal None, no DCE | E-06 + incidencia G-06 |
| T-01 | sin RUT / con RUT / "Institución:" / sin columna | sí / no / no / no |
| T-02 | oído hoy / 45 / 46 / futuro | sí / sí / no / no |
| T-03 | audiencia ayer / hoy / +300 | no / sí / sí |
| E-01 | "OPD X" con audiencia futura | solo texto no-seguimiento, sin audiencia |
| Ejemplos catálogo | 10.1 / 10.2 / 10.3 | byte a byte |

**C.2 CUMPLIMIENTO**

| Caso | Valores | Esperado |
|---|---|---|
| C-03 bordes | DIAS CUMPL. 0 / 30 / 31 | sí / sí / no |
| C-03 fuente | días=15 pero fecha ingreso hace 90d | sí (manda la columna) |
| C-03 sin fecha | días=10, sin FEC.INGRESO | omitida + incidencia |
| C-04 bordes | dias_egresar −1 / 0; dias_cumpl −1 | sí / no (→C-05 hoy) / sí |
| C-04 sin fecha | dias_egresar −5 sin egreso proy. | omitida + incidencia; texto sin fecha de hoy inventada |
| C-05 bordes | 0 / 1 / 45 / 46 | "día de hoy" / por vencer / por vencer / no |
| C-10 supresión | match + C-04; match + C-05; match solo; match + C-03 | no / no / sí / sí (ambas) |
| Hoja2 índice | vencimiento hoy / mañana; duplicados | excluido / incluido; gana última fila |
| C-07 ramas | RTA sin fecha / 181d / 180d / 31d / 30d / 0d | confeccionar / actualizar / nada / nada / reciente / reciente |
| C-07 prefijos | RVA / PEE / RPPM / PIE / "RESIDENCIA …" | sí / **no** / **no** / no / sí |
| C-08 | FAE sin ficha 121d / 120d / con ficha / FAS / "Residencia San Rafael" | sí / no / no / sí / **no** |
| C-09 | nada dispara / solo curador / solo C-02 | "sin observaciones" / base breve + curador / base breve + C-02 (D3) |
| C-01 | ≥18 + curador sin RUT | mayoría + curador (H-09) |
| Ejemplos catálogo | 10.4 / 10.5 | byte a byte |

**C.3 INFORMES**

| Caso | Valores | Esperado |
|---|---|---|
| I-01/I-02 bordes | días −1 / 0 / 30 / 31 | vencido / por vencer / por vencer / nada |
| DCE | vencido y por vencer | "informe diagnóstico" en ambos |
| E-01 | "OPD X" con fecha vencida | texto no-seguimiento (no I-01) |
| Sin fecha | vencimiento vacío | `""` + incidencia |
| Ejemplo catálogo | 10.6 | byte a byte |

**C.4 Transversales:** ninguna salida con `{`/`}`/`..`/`ERROR`; un solo punto final; prefijo `{PrimerNombre} {SIGLA}: `; fechas en español formato largo; columnas `FECHA_OBS, OBSERVACION, TT, CC, RES` en orden y con los valores de D1/D2; tribunal no reconocido nunca hereda LAJA; correos: bordes −1/0/30/31, DCE en asunto/cuerpo, tribunal normalizado, un borrador por variante de escritura; resoluciones: PC_IE/PC_INFO/no-colisión C-04/NOMENCL dormant.
