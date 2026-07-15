# CHANGELOG — CSMP Assistant

---

## [8.0.5] — 2026-05-15 — Correos espera + formato correos + negrita resoluciones

### Añadido
- **Correos lista de espera (tipo D)**: `GeneradorCorreos.procesar_espera(df)`
  - Detecta filas en Excel ESPERA donde OBSERVACION contiene "se remite correo electrónico al programa consultando respecto de fecha estimada de ingreso"
  - Genera 1 borrador Outlook por PROGRAMA+TRIBUNAL
  - Tabla: RIT, Tribunal, RUT NNA, Nombre NNA, Días en espera
  - Columna T ESPERA tomada directamente del Excel
- **Plantilla `correo_programa_espera.html`**: nueva plantilla para lista de espera
- **Botón GUI**: "✉ Generar borradores Lista de Espera" en pestaña Correos Outlook

### Modificado
- **Todas las plantillas de correo**: eliminada mención al director. Nuevo formato:
  ```
  SRES. NOMBRE PROGRAMA
  PRESENTE.
  ```
- **`_formatear_nombre_programa()`**: nuevo helper
  - Siglas (≤4 chars, solo letras, posición 0-1) → MAYÚSCULAS
  - Resto → capitalize()
  - Guiones/rayas aislados eliminados
  - Ejemplo: `"AFT - MULCHEN"` → `"AFT Mulchen"`
- **Resoluciones .docx**: datos extraídos del Excel ahora en **negrita**
  - Nombre NNA → negrita + Title Case (`"JUAN PEREZ"` → `"Juan Perez"`)
  - Programa/derivación → negrita + formato sigla (`"AFT - MULCHEN"` → `"AFT Mulchen"`)
  - Fecha resolución, RUT, duración → negrita
  - Nomenclatura (NOMENCL) → negrita + cursiva
- **`_titulo_nombre()`** y **`_titulo_programa()`**: nuevos helpers en `generador_resoluciones.py`

---

## [8.0.4] — 2026-05-15 — Fase 4 Completa

### Añadido
- Módulo `resoluciones/`: generador .docx (PC_IE, PC_INFO, NOMENCL)
- Formato: Arial 12, justificado, interlineado 1.5
- 1 Word consolidado, 1 RIT por página
- Excel faltantes para casos sin plantilla (Tomé, PC_INFO Laja)
- Reintegración validador y correos (omitidos en entrega anterior)

---

## [8.0.3] — 2026-05-15 — Fase 3 Completa
- Correos Outlook Drafts, matching fuzzy 3 niveles, firma preservada

## [8.0.2] — 2026-05-15 — Fase 2 Completa
- Logs persistentes, rotación 90 días

## [8.0.1] — 2026-05-15 — Fase 1 Completa
- Validador 8 anomalías A1-A8, reporte HTML

## [8.0.0] — 2026-05-15 — Fase 0 Completa
- Refactor modular desde RUS_Engine_v7.3.py monolítico

## [7.3] — base
- Motor monolítico 26 reglas

---

## [8.1.0] — 2026-05-15 — Cruce Hoja2, formatos globales, PROXS. AUDS., limpiar_nombre

### Añadido
- **Cruce Hoja2 en CUMPLIMIENTO**: si el archivo fuente tiene 2 pestañas,
  la segunda (informes por vencer) se cruza con la primera por 5 campos
  (RIT + RUT + NOMBRE + DERIVACIÓN + TRIBUNAL).
  Si hay match y R5 no aplica → observación de informe usa fecha de Hoja2.
  Sin Hoja2 → motor funciona igual que antes.
- **PROXS. AUDS.**: si la columna existe y tiene valor en cualquier modo
  (ESPERA / CUMPLIMIENTO / INFORMES), se agrega al final:
  "Se cita a audiencia para el día DD de mes de AAAA."
- **Aliases Hoja2**: mapeo_columnas detecta RUT MENOR / NOMBRE MENOR /
  NOMBRE CENTRO como equivalentes de RUT / NOMBRE / DERIVACIÓN.

### Modificado
- **`limpiar_nombre()`** (nueva función en utilidades.py):
  elimina paréntesis y su contenido (`"JUAN () PEREZ (TUT)"` → `"Juan Perez"`),
  colapsa espacios, Title Case.
- **`prefijo_observacion()`**: usa `limpiar_nombre()` — todos los prefijos
  ya no incluyen paréntesis ni contenido entre paréntesis.
- **`titulo_programa()`**: formatea derivaciones en observaciones:
  siglas ≤4 chars → MAYÚSCULAS, resto → capitalize(), colapsa espacios/guiones.
  `"AFT  MULCHEN"` → `"AFT Mulchen"`.
- **Todas las fechas** en observaciones de los 3 modos → `fecha_es()`
  (formato: "13 de mayo de 2026").
- **`normalizar_match()`** (nueva función): normalización específica para
  cruce entre hojas (lowercase + colapso whitespace, sin eliminar guiones).
- **`procesador.py`**: modo CUMPLIMIENTO acepta path de archivo (para leer
  Hoja2) o DataFrame directo (sin cruce). ESPERA/INFORMES sin cambios.
- **`gui/app.py`**: botón CUMPLIMIENTO pasa path al procesador en vez de
  DataFrame precargado.

---

## [8.2.0] — 2026-05-15 — Fixes R2 mayoría, cruce Hoja2, titulo_programa

### Corregido
- **R1/R2 (mayoría de edad)**: edad calculada desde `FEC. NACIMIENTO` exacta,
  no desde columna `EDAD` (que puede ser stale). R1 corta si edad_real ≥ 18.
- **R2 (próximo mayoría)**: solo dispara si ≤60 días para los 18 años.
  Texto corregido: `"Se hace presente que [Nombre] alcanzará la mayoría de edad el [fecha]"`.
  Acumulable con otras reglas.
- **Cruce Hoja2 — medida vencida**: cruce bloqueado si `DÍAS PARA EGRESAR ≤ 0`.
  Antes aplicaba cruce incorrectamente sobre medidas ya vencidas.
- **Cruce Hoja2 — fecha pasada**: al construir el índice se descartan todas las
  filas de Hoja2 donde `FECHA VENCIMIENTO < hoy`. El índice solo contiene fechas futuras.
- **`titulo_programa()`**: preposiciones cortas (de/del/los/las/y/en/el/la/por/con/a)
  → lowercase. `"Provincia DEL Biobio LOS Angeles"` → `"Provincia del Biobio los Angeles"`.

---

## [8.2.0 — Roadmap] — 2026-05-15 — Fase 5 postergada, Fase 6 próxima

### Decisión
- **Fase 5 (Reportes consolidados)**: postergada por decisión del usuario. Sin ETA.
- **Fase 6 (GUI expandida)**: designada como próxima fase.

### Alcance Fase 6
Integrar comunicaciones, resoluciones, validador y logs en pestañas Tkinter.
gui/app.py pasa de 1 pestaña (Motor RUS) a 5 pestañas completas.
Motor v8.2.0 no se toca en esta fase.

---

## [8.3.0] — 2026-05-15 — Limpieza, refactor visual

### Eliminado (código muerto)
- `dias_desde()` en utilidades.py — nunca llamada
- `sanitizar()` en utilidades.py — importada pero nunca invocada

### Refactorizado
- `_audiencia_suffix()`: antes triplicada en reglas_espera/cumplimiento/informes.
  Centralizada en utilidades.py como `audiencia_suffix()`, importada desde allí.
- `_procesar_espera()` + `_procesar_informes()` en procesador.py:
  colapsadas en `_procesar_simple(df, modo, fn_regla, config, q)`.
- `sanitizar` eliminado del import de procesador.py.
- Import `_audiencia_suffix` inline en procesador.py eliminado.

### Visual GUI (gui/app.py)
- Título: "Asistente de RUS — CSMP Concepción"
- `minsize(820, 600)` — ventana no se puede encoger hasta ilegible
- Progressbar: pasos reales (5 → 20 → 90 → 100) en vez de salto directo
- Bitácora oscura (fondo #1e1e1e) con colores por tipo de mensaje:
    ✅ / 📁  → verde agua  (éxito)
    ❌       → rojo negrita (error)
    ⚠️       → naranja (advertencia)
    ℹ️ / ✓  → azul (información)
    done     → verde claro negrita
- Subtítulo institucional bajo el título principal
- Versión visible en barra de estado inferior derecha
- Separador visual entre encabezado y configuración
- Selector de Excel acepta .xlsx, .xls y .xlsm

---

## [8.4.0] — 2026-05-16 — Corrección 4 bugs auditados

### BUG-01 — get_date(): parsing determinista de fechas
Antes: `pd.to_datetime(s, dayfirst=True)` en todos los casos.
Fechas ISO `YYYY-MM-DD` con día ≤ 12 se invertían (ej: `2026-01-05` → 5 mayo).
Ahora: detecta formato con regex antes de parsear.
- `YYYY-MM-DD` → `format="%Y-%m-%d"` (sin ambigüedad)
- `DD/MM/YYYY` → `format="%d/%m/%Y"` (formato RUS)
- Otros → heurística pandas como fallback
Elimina además el `UserWarning` en consola.

### BUG-02 — audiencia_suffix(): puntuación correcta
Antes: sufijo `" Se cita..."` sin punto inicial → `"...ingreso Se cita..."`.
Ahora: sufijo `". Se cita..."` y todos los fragmentos en `obs[]` terminan
en punto. Se elimina el punto final del fragmento antes de unir con el sufijo
para evitar doble punto.
Afectaba: reglas_espera, reglas_cumplimiento, reglas_informes.

### BUG-03 — normalizar_match(): guard pd.isna()
Antes: `str(pd.NA)` → `'<NA>'` → clave `'<na>'` en índice → nunca matcheaba.
Ahora: `pd.isna(txt)` al inicio → retorna `''` para todos los nulos de pandas.

### BUG-04 — normalizar_match(): limpia () antes de comparar (CRÍTICO)
Antes: `'JORGE CONTRERAS ()'` generaba clave distinta a `'JORGE CONTRERAS'`.
12% de NNA en CUMPLIMIENTO tenían `()` en nombre; Hoja2 no los tiene.
El cruce fallaba silenciosamente para esos casos.
Ahora: `re.sub(r'\(.*?\)', '', s)` antes de lowercase → claves iguales.
Verificado: NNA con `()` ahora cruzan correctamente con Hoja2.

---

## [8.5.0] — 2026-05-16 — Corrección 5 bugs segunda auditoría

### BUG-A — GUI: botones bloqueados si procesador retorna sin done
`_procesar_simple()` y `_procesar_cumplimiento()` ahora garantizan que
`("done",...)` se envíe en todos los paths de retorno: columnas faltantes,
error al guardar, y excepción en carga de archivo.
Antes, un Excel con columnas incorrectas dejaba la GUI bloqueada
hasta cerrarla.

### BUG-B — audiencia_suffix(): filtra fechas de audiencia pasadas
Si `PROXS. AUDS.` contiene una fecha anterior a hoy, el sufijo
no se genera. Antes aparecía en la observación aunque la audiencia
ya hubiera ocurrido.

### BUG-C — limpiar_nombre(): elimina residuos de paréntesis mal formados
Añadido `re.sub(r'[()]', '', s)` tras la limpieza de pares `()`.
'ANA ((doble))' → 'Ana'  |  'JUAN )huerfano(' → 'Juan Huerfano'.

### BUG-D — GUI: barra de progreso retrocedía de 100 a 90
Eliminado `q.put(("progress", 90))` del `_worker`. `procesar()` pone
`done` internamente; luego `_finish_job` establece la barra en 100.
El 90 llegaba después del 100 y bajaba la barra visualmente.

### BUG-E — prefijo_observacion(): fallback cuando nombre queda vacío
Si el nombre queda vacío tras limpiar paréntesis (ej: nombre = '()'),
el prefijo ahora devuelve las 3 letras del programa ('AFT: ')
en vez de string vacío.

### Mejoras adicionales GUI
- Límite de 500 líneas en la bitácora (evita crecimiento indefinido)
- Botón "🗑 Limpiar" en la bitácora
- Guard de dependencias al inicio: mensaje de error claro si pandas falta

---

## [8.6.0] — 2026-05-16 — Optimizaciones de rendimiento

Basado en auditoría externa. Claims incorporados tras verificación con benchmarks reales.

### CLAIM-1 — relativedelta a nivel de módulo
`from dateutil.relativedelta import relativedelta` movido al inicio de
`utilidades.py`. Antes se importaba dentro de `proximo_informe()` en
cada llamada. Python cachea módulos, por lo que el impacto real es mínimo
(~0.2 µs/llamada), pero es buena práctica y elimina la importación en caliente.

### CLAIM-2 — Regex pre-compiladas
8 expresiones regulares ahora pre-compiladas a nivel de módulo:
`_PAREN_RE`, `_PAREN_SUELTO_RE`, `_ESPACIOS_RE`, `_GUIONES_RE`,
`_NO_LETRAS_RE`, `_SEP_PROG_RE` (nuevas) + `_ISO_RE`, `_LATAM_RE` (ya existían).
Benchmark: 1.5x speedup en `limpiar_nombre` con datasets grandes.
La afirmación "10x más rápido" de la auditoría es incorrecta — el speedup
real medido es 1.5x.

### CLAIM-3 — Escritura Excel con dataframe_to_rows
`_guardar_excel` reemplaza escritura iterativa (`for _, row in df.iterrows()`)
por `dataframe_to_rows(df, index=False, header=True)`.
Benchmark: 3x más rápido (42.8 ms → 14.4 ms en 525 filas).

### CLAIM-4 — Índice Hoja2 con to_dict('records')
`_construir_indice_hoja2` reemplaza `iterrows` por `to_dict('records')`.
Benchmark: 8.6x más rápido (421 ms → 49 ms en ~10.000 filas).
Este era el cuello de botella real, no el import de relativedelta.

### CLAIM-5 — Ruta dinámica en DEFAULT_CONFIG
Reemplazada ruta estática `C:\Users\cmgaete\Desktop\Matias\...` por
`Path.home() / "CSMP_RUS"`. Funciona en cualquier equipo Windows/Mac/Linux.
La propuesta de la auditoría (`Path.home() / "Desktop"`) fue rechazada:
en equipos con OneDrive el Desktop está en ruta distinta.

### Claims rechazados (no incorporados)
- CLAIM-6: Eliminar try/except en pd.isna() — rechazado porque pd.isna([])
  devuelve array, causando ValueError. El try/except es necesario.
- CLAIM-7: PyInstaller con gui/app.py — corregido a main.py en instrucciones.

---

## [8.7.0] — 2026-05-18 — Fix crítico ModuleNotFoundError + GUI bloqueada

### Causa raíz
Los ZIPs parciales anteriores (v8.4–v8.6) contenían solo los archivos
modificados. Al descomprimir en carpeta nueva faltaban: comunicaciones/,
resoluciones/, validador/, logs/ → "módulo de correos perdido".

Adicionalmente, `from motor.procesador import procesar` estaba dentro de
`_worker()`. Si sys.path no incluía la raíz (ejecución directa de app.py
sin pasar por main.py), lanzaba ModuleNotFoundError en el thread, que no
era capturado → done nunca llegaba → GUI bloqueada permanentemente.

### Fixes
1. `gui/app.py`: agrega `_ROOT` (carpeta raíz) a `sys.path` al inicio
   del módulo, antes de cualquier import. Funciona independientemente
   de si se ejecuta via `main.py` o directamente.
2. `gui/app.py`: `from motor.procesador import procesar` movido a nivel
   de módulo (no dentro de `_worker`). Error de import ahora se detecta
   al iniciar la app con mensaje claro, no al presionar un botón.
3. ZIP ahora es COMPLETO (todo el proyecto, no solo archivos modificados).

---

## [8.8.0] — 2026-05-18 — GUI expandida: Correos y Resoluciones integrados

### Nuevo: 3 pestañas en la GUI

**Pestaña "Motor RUS"** (sin cambios funcionales)
- ESPERA / CUMPLIMIENTO / INFORMES + Abrir salida
- Bitácora con colores + botón limpiar

**Pestaña "Correos Outlook"** (nuevo)
- Botón "✉ Informes vencidos / por vencer": lee Excel de INFORMES,
  genera borradores Outlook (1 por programa vencidos + 1 por tribunal + 1 por programa por vencer).
- Botón "✉ Lista de espera": lee Excel de ESPERA procesado por el motor
  (requiere columna OBSERVACION), genera 1 borrador por programa+tribunal.
- Selector de Excel propio (independiente del Motor).
- Requiere pywin32 y Outlook instalado — muestra error claro si falta.

**Pestaña "Resoluciones"** (nuevo)
- Botón "⚙ Generar resoluciones .docx": lee Excel con columna OBSERVACION,
  detecta automáticamente PC_IE / PC_INFO / NOMENCL según tribunal,
  genera 1 Word consolidado (1 RIT por página).
- Selectores independientes para Excel entrada y carpeta salida.
- Muestra advertencias por filas sin plantilla disponible.

### Otros cambios
- `_write()`: método centralizado para escribir en cualquier Text widget con colores.
- Todos los workers en threads separados — GUI nunca se congela.
- `_set_correo_btns` y `_res_btn` se deshabilitan durante procesamiento.

---

## [8.9.0] — 2026-05-18 — Fix correos Outlook: GetInspector + matching contactos

### BUG: Error COM -2147467259 en todos los borradores
`GetInspector` fuerza apertura de ventana de composición en Outlook.
Si Outlook está minimizado, sin foco o en modo protegido, devuelve
`E_UNEXPECTED (-2147467259)` — todos los borradores fallaban.

Fix en `comunicaciones/generador_correos.py`:
`GetInspector` se intenta primero. Si falla, fallback a leer `HTMLBody`
directamente (también contiene la firma en la mayoría de configuraciones).
Si tampoco hay firma, se guarda el borrador sin ella.

### BUG: Sin contacto para programas con nombre levemente distinto
El matching de 2 niveles (exact + substring) no alcanzaba variantes como:
- `'AFT  EL CONQUISTADOR YUMBEL'` vs `'AFT - EL CONQUISTADOR DE YUMBEL'`
- `'PIE - PROGRAMA... ESPECIALIZADA'` vs `'PIE - PROGRAMA... IESPECIALIZADA'`

Fix en `comunicaciones/contactos_programas.py`: agregado nivel 3 de matching
(token-set): si ≥2 tokens significativos (>3 chars) coinciden con ratio ≥45%,
se considera match. Elimina paréntesis y contenido antes de comparar.

Resultado con Excel real (333 filas):
- Vencidos: 4/4 grupos resuelven contacto (antes 3/4)
- Por vencer: 32/35 grupos resuelven contacto
- 3 sin contacto son genuinamente ausentes del catastro:
  DCE - KAMULAYAN, FAE - AADD SENAME CONCEPCIÓN, RFA - RESIDENCIA FAMILIAR AADD CASTELLÓN
  → Deben agregarse manualmente a catastro_programas.json

---

## [8.9.1] — 2026-05-18 — Aliases de programas para nombres distintos en RUS

### Problema
3 programas en el catastro tienen nombre distinto al que exporta RUS:
- RUS: `DCE - KAMULAYAN` → catastro: `DCE - KAMULYAN`
- RUS: `FAE - AADD SENAME CONCEPCIÓN (PROV. CONCEPCIÓN-ARAUCO)` → catastro: `FAE AADD FAMILIA DE ACOGIDA DE ADMINISTRACIÓN DIRECTA DE CONCEPCIÓN`
- RUS: `RFA - RESIDENCIA FAMILIAR AADD CASTELLÓN` → catastro: `RFA – CASTELLÓN`

El token-set matching (nivel 3) no los resuelve porque las diferencias son
demasiado grandes. Subir el umbral generaría falsos positivos.

### Solución
Nuevo archivo `comunicaciones/aliases_programas.json` con mapeo explícito:
alias (nombre RUS) → nombre_catastro.

`CatastroContactos._cargar()` levanta el archivo al iniciar si existe,
registrando cada alias como entrada adicional en el índice.
Si el archivo no existe, la carga continúa sin error.

Para agregar nuevos aliases en el futuro, editar directamente
`comunicaciones/aliases_programas.json` sin tocar código.

---

## [8.10.0] — 2026-05-18 — Fix formato correos: encabezado y firma

### Encabezado de correo (3 plantillas de programa)
Antes: `SRES. {{NOMBRE_PROGRAMA}} / PRESENTE.` (PRESENTE solo en negrita)
Ahora:
```
SEÑORES:
{{NOMBRE_PROGRAMA}}
PRESENTE.   ← negrita + subrayado
```
Afecta: correo_programa_vencidos.html, correo_programa_por_vencer.html,
correo_programa_espera.html.

### Firma del usuario ausente
`GetInspector` no cargaba la firma si Outlook estaba minimizado o sin foco.
Fix: usar `mail.Display(False)` antes de leer `HTMLBody`. Este método
abre el compositor en segundo plano y fuerza la inyección de la firma
sin mostrar ninguna ventana. Luego `inspector.Close(0)` cierra sin guardar.
El fallback anterior (leer HTMLBody sin Display) se mantiene si Display falla.

### Nombre NNA en tabla: elimina paréntesis
`_titulo_nombre()` ahora elimina `()` y su contenido antes de formatear.
`'Martina Elena Rivas González ()'` → `'Martina Elena Rivas González'`

---

## [8.11.0] — 2026-05-18 — Fix R6 curador y R8 ficha residencial

### BUG-R6: curador disparaba en todas las filas cuando columna no existía
Causa: `cols["curador"]` devuelve `None` si la columna no se llama exactamente
"CURADOR" en el Excel (ej: "CURADOR AD LITEM"). `row.get(None, "")` devuelve `""`
→ la condición `not curador` era `True` → R6 disparaba para todas las filas.

Fixes:
1. `motor/mapeo_columnas.py`: agregar aliases para CURADOR:
   "CURADOR AD LITEM", "CURADOR AD-LITEM", "CUR. AD LITEM", "CURADOR/A AD LITEM", etc.
2. `motor/reglas_cumplimiento.py`: R6 solo evalúa si `cols.get("curador")` no es None.
   Ausencia de columna ≠ ausencia de curador.

### BUG-R8: ficha residencial disparaba en programas ambulatorios (PIE, AFT, PAS, FAE)
Causa: R8 solo verificaba `cols.get("ficha_res")` (existencia de la columna),
sin verificar si el programa es de tipo residencial.
Si el Excel incluye la columna FEC.ACT.F.RESIDENCIAL, R8 aplicaba a todos.

Fix: R8 ahora verifica prefijo del programa antes de disparar.
Solo aplica a residencias: RTA, RTT, PRM, RES, PEE, RFA, RPPM.
PIE, AFT, PAS, FAE, OPD, DAM, DCE → no disparan R8.

---

## [8.11.1] — 2026-05-19 — Fix R6 curador: detección por RUT real

### Problema
R6 (curador ausente) no disparaba cuando la columna existe pero contiene
un placeholder de RUS indicando ausencia: '---', 'NO POSEE', 'Sin designar', etc.
La condición anterior solo detectaba "" y "SIN CURADOR".

### Fix
Nueva función `tiene_curador_real(val)` en `motor/utilidades.py`:
detecta si el valor contiene un RUT chileno (patrón `\d{6,8}-[\dkK]`).
Un curador real siempre viene con su RUT en el formato de RUS.
Cualquier valor sin RUT (vacío, '---', 'NO POSEE', etc.) = sin curador.

R6 ahora usa: `not tiene_curador_real(curador)` en vez de `not curador or "SIN CURADOR" in ...`

Comportamiento resultante:
- '' → dispara ✓
- '---' → dispara ✓
- 'NO POSEE' → dispara ✓
- 'Sin designar' → dispara ✓
- '(17955927-7)SEBASTIÁN BUSTAMANTE...' → NO dispara ✓
- Sin columna en Excel → NO dispara ✓

---

## [8.11.2] — 2026-05-19 — Fix cruce Hoja2: fechas vencidas marcadas como por vencer

### Problema
El índice de cruce Hoja2 usaba `vd < hoy` (estrictamente menor).
Informes con fecha == hoy entraban al índice y generaban observación
"deberá remitir informe... el 19 de mayo de 2026" cuando ese informe
ya estaba vencido en ese mismo día.

### Fix
`motor/procesador.py`, `_construir_indice_hoja2()`:
`if vd < hoy` → `if vd <= hoy`

Solo entran al índice fechas estrictamente futuras (> hoy).
Resultado: índice pasó de 122 a 116 entradas con el archivo de prueba
(6 informes con vencimiento = hoy correctamente excluidos).

---

## [8.11.3] — 2026-05-20 — Fix textos R1, R4 y filtros R8, R9

### R1 — Mayor de edad: texto con fecha y sugerencia de egreso
Antes: `"Mayor de edad"` (corte sin más información).
Ahora: `"Se hace presente que {Nombre} alcanzó la mayoría de edad el {fecha}, se sugiere egresar la medida."`
Calcula la fecha desde FEC.NACIMIENTO con `fecha_mayoria()`.

### R4 — Medida vencida: texto corto y directo
Antes: `"Medida vencida. Se hace presente que tanto la vigencia de la medida como el próximo informe de avance ya vencieron. Egreso proyectado: {fecha}."`
Ahora: `"La medida se visualiza vencida en RUS desde el {fecha}."`

### R9 — Ficha individual: solo aplica a residencias
Antes: solo verificaba si la columna FEC.ACT.F.INDIVIDUAL existía.
Disparaba para AFT/PIE/PAS/FAE/DCE/PRM también.
Ahora: filtra por prefijo igual que R8 — solo RTA/RTT/RES/PEE/RFA/RPPM.
Excluidos explícitamente: AFT, PIE, PAS, PRM, FAE, DCE.

### R8 — Ficha residencial: quitar PRM
Por consistencia con R9. PRM (Programa de Reparación de Maltrato) NO es residencial.
Lista final R8: RTA, RTT, RES, PEE, RFA, RPPM.

---

## [8.12.0] — 2026-06-05 — Auditoría completa: fechas, curador en espera, encabezado Mulchén

### BUG CRÍTICO — Fechas en formato inglés (resoluciones)
`fecha_numerica()` y `fecha_en_palabras()` usaban `pd.to_datetime(fecha)` sin
`dayfirst=True`. Una fecha 3/6/2025 se interpretaba como 6 de marzo (formato US)
en vez de 3 de junio. Nuevo parser `_parse_fecha()` determinista:
- ISO (YYYY-MM-DD) → formato explícito
- LATAM (DD/MM/YYYY o DD-MM-YYYY) → dayfirst, normaliza guiones a barras
- datetime/Timestamp → directo
Si no parsea, retorna "COMPLETAR" en vez de fallar.
También reforzado `_fila_fecha()` en generador_correos.py con dayfirst=True.

### BUG CRÍTICO — R7 curador no disparaba en ESPERA
`reglas_espera.py` usaba la lógica vieja (`not curador or "SIN CURADOR"`),
mientras cumplimiento ya usaba `tiene_curador_real()` (detección por RUT).
Por eso en espera NO se marcaba la ausencia de curador cuando RUS ponía
'---', 'NO POSEE', 'Sin designar'. Ahora ambos modos usan la misma función.

### Encabezado institucional Mulchén eliminado (resoluciones)
`_header_mulchen()` ya no agrega las 4 líneas:
- "Juzgado de Letras y Garantía de Mulchén"
- "Calle Villagra N° 047, Mulchén, fono..."
- "Correo electrónico jlyg_mulchen@pjud.cl"
- "ATENCIÓN DE PÚBLICO VIRTUAL: conecta.pjud.cl"
El documento comienza directo con "Mulchén, {fecha}."

### Mejoras adicionales
- R1/R2 en ESPERA alineadas con CUMPLIMIENTO: edad exacta desde FEC.NACIMIENTO
  y fecha de mayoría calculada (antes usaba solo el campo EDAD entero).
- R1 ESPERA ahora tiene texto completo: "Se hace presente que {Nombre} alcanzó
  la mayoría de edad el {fecha}, se sugiere egresar la medida." (antes solo "Mayor de edad").
- `_titulo_nombre()` en resoluciones limpia paréntesis: 'JUAN PEREZ ()' → 'Juan Perez'.
- Tildes corregidas en números a palabras: veintidós, veintitrés, veintiséis.

### Regresión verificada
- ESPERA / CUMPLIMIENTO / INFORMES: los 3 modos procesan OK con XLS reales.
- Cruce Hoja2: intacto.
- Correos: matching catastro + 3 aliases + encabezado SEÑORES/PRESENTE intactos.
- Resoluciones: detectar_tipo captura observaciones reales del motor.
- Todos los módulos importan sin error.

---

## [8.12.1] — 2026-06-09 — Textos literales en ESPERA + fix doble punto + optimizaciones

### BUG — Textos de ESPERA no literales (reportado en producción)
La reescritura v8.12 alteró los textos del manual:
1. Agregó puntos finales a los fragmentos R3-R8 (el esquema original v7.1
   une fragmentos SIN punto con ". ").
2. El fallback usaba el programa formateado (titulo_programa) en vez del
   programa crudo del Excel: "...ingreso efectivo a {programa}".
Restaurados los textos LITERALES v7.1 como constantes de módulo
(_T_R0, _T_R4_DCE, _T_R5_MULCHEN, _T_R6_LAJA_TOME, _T_R7_CURADOR, _T_R8_OIDO).
Únicos cambios autorizados que se mantienen: R1 con fecha y sugerencia
(texto entregado textual por el usuario), R2 alineada con cumplimiento,
R7 detección curador por RUT.

### BUG — Doble punto con audiencia en R1 (espera y cumplimiento)
R1 retorna texto terminado en "." y audiencia_suffix empieza con ". "
→ "...egresar la medida.. Se cita a audiencia...". Fix: si hay sufijo
de audiencia, se elimina el punto final del texto antes de concatenar.

### Optimizaciones de rendimiento
- limpiar_nombre(): se llamaba hasta 4 veces por fila en cumplimiento
  (R1 x2, R2, R10) y 3 en espera. Ahora 1 vez al inicio (pnombre).
- R10: eliminado programa_norm.upper() redundante — normalizar() ya
  devuelve lowercase; se compara "fae"/"fas" directamente.
- Textos literales de espera como constantes de módulo (no se
  reconstruyen los strings en cada fila).

### Validación
- Paridad textual: 12/12 reglas producen el texto literal exacto.
- Regresión 3 modos con XLS reales: OK (1.1 s total).
- Outputs: 0 doble punto, 0 errores, 0 observaciones vacías.

---

## [8.13.0] — 2026-07-07 — Externalizacion de textos + fix ESPERA + hardening

### Causa raiz resuelta: deriva textual entre sesiones
Textos de observacion (ESPERA/CUMPLIMIENTO/INFORMES/cruce Hoja2) migrados
desde strings hardcodeados en .py a `motor/textos_observaciones.json`,
extraidos verbatim del codigo fuente (no retipeados de memoria).
Nuevo `motor/textos.py`: `render(modo, id_regla, **datos)` con SafeDict
(placeholder faltante -> `{CAMPO}` visible, nunca crash ni texto vacio).
`listar_no_confirmados()`: utilidad de auditoria, expone que textos
quedan como baseline heredado v7.1/v8.x sin reconfirmar caracter-por-caracter.

### FIX — ESPERA generaba texto rechazado por el usuario (DCE < 30 dias)
Reportado 2026-06-11 con archivo real `RUS_ESPERA_20260610_085639.xlsx`
(16/16 filas eran el caso disputado). Texto viejo: "Medida en espera desde
hace N dias". Nuevo comportamiento: si `FEC. RESOLUCION` esta presente
(columna ahora mapeada, antes ausente — hallazgo F2 de auditoria previa),
usa el texto literal entregado por el usuario:
"Medida revisada, a la espera de ingreso efectivo. Se hace presente que
se ordena ingreso a {PROGRAMA} con fecha {FECHA_RESOLUCION}"
Aplica tanto al caso DCE<30d como al fallback generico (Mulchen/Laja/Tome
bajo umbral) cuando existe fecha de resolucion — generalizacion explicita,
no eran casos disputados individualmente pero comparten la misma condicion
semantica ("nada severo disparo, pero hay resolucion de ingreso").
Sin FEC. RESOLUCION disponible: se preserva el texto v7.1 como residual
(marcado `confirmado:false` en el JSON).

Validado con subprocess aislados (no contaminacion sys.path) contra el
archivo real: 15/16 filas cambiaron al texto correcto, 1/16 sin cambio
(caso DCE>=30d, correctamente fuera de alcance), 0 doble punto, 0 errores,
0 observaciones vacias.

### Fix bug propio detectado en validacion: doble punto
Primera version del texto nuevo llevaba punto final; el esquema de union
de ESPERA (". ".join sin punto en fragmentos) generaba "...2026.. No
registra...". Corregido: texto sin punto final, igual al ejemplo literal
original del usuario.

### Regresion verificada (subprocess aislados vs baseline v8.12.1 real)
- CUMPLIMIENTO: 0 filas distintas / 100 (cc46_cga_amblistcump__63_.xls)
- INFORMES: 0 filas distintas / 333 (RUS_INFORMES_20260518_093752.xlsx)
- Confirma: la migracion a JSON no altero ningun comportamiento existente,
  solo el mecanismo de almacenamiento del texto.

### Hardening (sin fabricar texto nuevo)
- `mapeo_columnas.py`: alias para `FEC. RESOLUCION` (6 variantes).
- `comunicaciones/generador_correos.py`: `except:` desnudo -> `except Exception:`.
- `config_rus.json` (con path `C:\Users\cmgaete\...` hardcodeado) eliminado
  del paquete. `gui/app.py` ya regenera un default portable
  (`Path.home()/CSMP_RUS`) desde v8.6 — el archivo shippeado lo pisaba.

### Pendiente — NO ejecutado en este turno (requiere confirmacion humana)
17 textos quedan `confirmado:false` (ver `listar_no_confirmados()` en
`motor/textos.py`). Ninguno fue tocado ni reformulado por inferencia:
se preserva exactamente el comportamiento anterior hasta que el usuario
entregue el texto literal. Lista completa en `motor/textos_observaciones.json`,
campo `nota` de cada entrada.

---

## [8.14.0] — 2026-07-10 — S1-S6: micro-fixes, tests permanentes, vista previa, confirmación de textos

Ejecutado contra `CSMP_Assistant_v8_13_COMPLETO.zip` (VERSION="v8.13.0" verificado
antes de tocar nada). Orden de ejecución: S1 → S2 → CA-S1/S2 → S3 → S4 →
CA-S3/S4 → S5 → S6 → CA-GLOBAL, según Plan de Ejecución Maestro v8.14.

### S1 — Micro-fixes (validados contra corpus, no solo "aplicados")
- `motor/utilidades.py` `tiene_curador_real()`: agrega `or ('Institución:' in str(val or ''))`
  junto al regex RUT existente (`\d{6,8}-[\dkK]`, sin tocar). Cubre curador
  institucional (ej. CAJ Biobío) sin RUT individual. Validado contra corpus
  de 30 valores sintéticos: 30/30 veredictos correctos, 0 cambios en
  placeholders (`---`, `NO POSEE`, `Sin designar`, vacío).
- `logs/log_manager.py:33`: `except:` desnudo → `except Exception:`. Verificado
  `grep -rP 'except\s*:\s*$'` → 0 en todo el proyecto.
- `motor/utilidades.py` `get_date()`: soporta serial Excel numérico
  (`30000 < val < 60000` → fecha vía epoch 1899-12-30). `get_date(45810.0)`
  → `2025-06-02`.

### S2 — Suite de tests permanente (`tests/`)
- `tests/goldens_textos.json` + `tests/generar_goldens.py`: congela el estado
  ACTUAL de `textos_observaciones.json` (31 reglas) como referencia fija.
- `tests/test_paridad.py`: compara `render()` contra los goldens (NO contra
  el JSON en caliente — sería tautológico). **Probado el detector**: sabotaje
  manual de un texto → 2 tests fallan inmediatamente; restaurado → verde de
  nuevo. 65 tests (31 render + 31 fuente + 3 guard-rails de inventario).
- `tests/runner.py`: ejecuta `motor.procesador.procesar()` en **subprocess
  aislado** (lección 2026-07-07: `sys.path` compartido dentro del mismo
  proceso de pytest produjo un falso "0 diffs").
- `tests/test_regresion.py`: dos modos — regresión REAL contra Excel de
  producción si se definen `CSMP_EXCEL_ESPERA/CUMPLIMIENTO/INFORMES`, o smoke
  sintético con `tests/fixtures/generar_fixtures.py` si no (este sandbox de
  desarrollo no tiene los 4 Excel reales de Matías — pendiente correr en su
  máquina, ver `SNAPSHOT.json` → `REGRESION_REAL_PENDIENTE`).
- `tests/test_confirmacion_textos.py`: ciclo completo del flujo S4.
- Resultado: **73 passed, 3 skipped** (los 3 skip son exactamente los que
  requieren Excel reales — comportamiento esperado, no un fallo).

### S3 — Vista previa en GUI (motor/procesador.py + gui/app.py)
- Refactor interno de `motor/procesador.py`: se extrajo el cálculo de
  observaciones (`_calcular_simple`, `_calcular_cumplimiento`, incluye cruce
  Hoja2) de el guardado a disco (`_guardar_excel`). Cero cambio de
  comportamiento — mismo resultado, ahora reutilizable.
- Nueva función pública `calcular_preview(df_o_path, modo)`: idéntica lógica
  a `procesar()`, no escribe archivo.
- GUI: botón "👁 Vista previa" por modo → ventana `Toplevel` con `Treeview`
  (columnas NNA | Observación, primeras 30 filas + contador total). Botón
  "Exportar…" dispara recién ahí el pipeline real de guardado.
- **Verificado con Xvfb + Tk real** (no solo lógica): la ventana de preview
  muestra contenido **idéntico byte-a-byte** al Excel que produce el botón
  ▶ normal, fila por fila, para los 3 modos.

### S4 — Flujo de confirmación de textos en GUI
- Nuevo módulo `motor/textos_confirmacion.py`: `exportar_pendientes()` /
  `importar_confirmaciones()` / `contar_pendientes()`. Backup automático con
  timestamp del JSON antes de sobrescribir. El texto que queda en el JSON es
  siempre el que el usuario escribió en `TEXTO_CORRECTO` — copy-paste literal,
  nunca parafraseado por código.
- GUI: sección "Confirmación de textos" con badge `📋 Textos pendientes: N`
  + botones "Exportar pendientes…" / "Importar confirmaciones…".
- Verificado con Xvfb + Tk real: badge muestra "17" al arrancar (coincide con
  el inventario conocido).
- Tras importar, `render()` refleja el cambio en la misma sesión sin editar
  ningún `.py` (invalidación de cache de `motor/textos.py`).

### S5 — Consistencia visual mínima
- Padding unificado a 8px en los 9 `LabelFrame` de la GUI (3 quedaban en
  `(6, 4)`: Bitácora, Resultado ×2). Sin cambio de paleta, tema ni framework.
- Verificado con captura de pantalla real (Xvfb) — cero cambios funcionales,
  regresión S2 sigue verde tras el cambio.

### S6 — Ejecutable distribuible
- Build de prueba con PyInstaller (`--onefile --add-data
  "motor/textos_observaciones.json:motor"`) ejecutado en Linux para validar
  la lógica de resolución de rutas en el bundle. **Verificado directamente**:
  el binario se auto-extrae a `_MEI*/motor/textos_observaciones.json`, JSON
  válido, 31 reglas, 0 errores en stdout/stderr.
- ⚠️ Pendiente: el binario real para producción (.exe) debe compilarse en
  Windows por Matías — mismo comando pero separador `;` en `--add-data`
  (Windows) en vez de `:` (Linux/Mac usado aquí solo para validar).

### Limpieza previa a la entrega
- `config_rus.json` (con paths hardcodeados de pruebas) eliminado del ZIP —
  se regenera portable al primer arranque (`Path.home()/CSMP_RUS`).
- Artefactos de build (`build/`, `dist/`, `*.spec`, `__pycache__/`) excluidos.

### Pendiente explícito (no ejecutado, requiere a Matías — ver SNAPSHOT.json)
- Confirmar los 17 textos `confirmado:false` vía el nuevo flujo S4.
- Correr la regresión real (`tests/test_regresion.py`) contra los 4 Excel de
  producción en su máquina.
- Compilar el `.exe` real en Windows y probarlo contra un Excel real de cada
  modo antes de usar en producción.
