# MEJORAS CSMP ASSISTANT — CORREOS Y OTROS MÓDULOS

**Complementa a:** `Catalogo_Reglas_CSMP_v2_2026-07-14.md`
**Fecha:** 14 de julio de 2026
**Estado:** Decisiones cerradas en sesión — insumo directo para implementación.

> Este documento cubre lo que el catálogo de reglas no cubre: `comunicaciones/` (correos), `validador/`, `resoluciones/`, `gui/app.py` y `motor/procesador.py`. Se revisó código real de `CSMP_Assistant_v8_14_COMPLETO.zip`, contrastado contra el Manual de Funciones y contra el catálogo de reglas ya cerrado.

---

# 0. Corrección al catálogo de reglas (§5, C-10)

**Revierte la decisión B3 de la ronda de auditoría.** Se había decidido que Hoja2 (C-10) nunca se suprime aunque ya haya disparado C-04 (medida vencida) o C-05 (próxima a vencer). Al revisar el código real se confirmó que **el comportamiento correcto es el opuesto**: C-10 **sí se suprime** cuando C-04 o C-05 ya dispararon — igual que hacía el antiguo C-06. Actualizar `Catalogo_Reglas_CSMP_v2_2026-07-14.md` §5 (C-10) y §11 en consecuencia.

---

# 1. Correos — `comunicaciones/`

## 1.1 · Eliminar el correo al Tribunal por informes vencidos

**Motivo:** el Manual dice textual, sobre informes vencidos: *"no se envía mail al juzgado, ya que se generan proyectos de resolución."* El código actual sí lo envía (`_trib_borrador`, plantilla `correo_tribunal_vencidos.html`, título "Informes vencidos en RUS — [Tribunal]"). Contradice al Manual.

**Decisión:** eliminar el envío. Consecuencia de implementación: la llamada a `_trib_borrador` para el grupo `vencidos` en `procesar()` se elimina. `_trib_borrador`, `correo_tribunal_vencidos.html` y la config `contactos_tribunales` (si no se usan en ningún otro lugar) quedan como código muerto a remover, no solo a deshabilitar.

## 1.2 · Corregir `PATRON_ESPERA`

**Motivo:** el fragmento de texto que usa `procesar_espera()` para filtrar filas de "lista de espera" (`"se remite correo electronico al programa consultando respecto de fecha estimada de ingreso"`) no coincide con el texto realmente aprobado, que dice *"...respecto de **la** fecha estimada de ingreso **efectivo**."* — nunca hace match, cero correos de lista de espera se generan.

**Decisión:** actualizar el fragmento a:
```
"se remite correo electronico al programa consultando respecto de la fecha estimada de ingreso efectivo"
```
Cubre uniformemente las tres ramas de E-05 (DCE, Laja/Mulchén, Tomé ambos tramos), ya que todas cierran con esa misma frase.

## 1.3 · Distinguir DCE en las plantillas de correo

**Motivo:** las plantillas de correo (`correo_programa_vencidos.html`, `correo_programa_por_vencer.html`) siempre dicen "informes de avance", sin distinguir DCE — pero la observación de RUS (I-01/I-02) sí distingue "informe diagnóstico" (DCE) de "informe de avance" (otros).

**Decisión:** detectar DCE en el correo (mismo criterio `es_dce()` del motor) y usar "informe diagnóstico" en ese caso, igual que ya hace el texto de la observación.

## 1.4 · Unificar formato de nombre de programa

**Motivo:** `generador_correos.py` tiene su propia función `_formatear_nombre_programa`, con un criterio de sigla distinto al del motor (posición 0-1 + ya en mayúscula, vs. cualquier posición + ≤4 letras en el motor).

**Decisión:** una sola función compartida (la del motor). Ver §1.9 — se integra a la unificación general de columnas y funciones comunes.

## 1.5 · Saludo personalizado con nombre del director — **rechazada**

Se mantiene el saludo genérico institucional ("SEÑORES: [programa] PRESENTE."). El campo `saludo` (nombre del director, ya calculado en `contactos_programas.py`) queda sin usar — no se conecta a las plantillas.

## 1.6 · Alinear `DIAS_POR_VENCER` del correo a 30 días

**Motivo:** el correo usaba `DIAS_POR_VENCER = 45` para clasificar "por vencer", mientras que la regla I-02 (la que genera el texto real en RUS) usa 30 días. Un caso a 40 días no tiene observación en RUS pero sí recibía correo.

**Decisión:** `DIAS_POR_VENCER = 30`.

## 1.7 · Corregir el límite del día 0

**Motivo:** el correo clasificaba `d <= 0` como VENCIDO; I-01 exige estrictamente `d < 0` (el día mismo, según I-02, es "por vencer", no vencido).

**Decisión:** cambiar la condición a `d < 0` para VENCIDO, dejando el día 0 en POR_VENCER — coherente con I-01/I-02.

## 1.8 · Alias RUT motor vs. correos — **absorbida por §1.9**

Se detectó que motor busca RUT solo como `["RUT", "RUT MENOR"]` mientras correos busca `["RUT", "RUT NNA", "RUT LITIGANTE"]`. Se resuelve como parte de la unificación general, no como fix puntual.

## 1.9 · Unificar TODAS las columnas comunes y funciones de formato compartidas (motor ↔ comunicaciones)

**Motivo:** el mismo patrón de desalineación apareció tres veces (formato de nombre de programa, alias de RUT, alias de DERIVACIÓN — motor reconoce "NOMBRE CENTRO" como alias y correos no).

**Decisión:** mover a un módulo compartido (ej. `motor/columnas_comunes.py` o similar) todas las listas de alias (programa/derivación, tribunal, nombre, RUT, RIT) y las funciones de formato (`titulo_programa`, `normalizar`, `detectar_tribunal`) usadas tanto por `motor/` como por `comunicaciones/`. Cambio de estructura, no de lógica de negocio.

**Nota de alcance:** `resoluciones/` queda explícitamente **fuera** de esta unificación (ver §2.3) — mantiene sus propias copias.

## 1.10 · Normalizar el nombre del tribunal en asunto/cuerpo del correo

**Motivo:** el asunto del correo usa el valor crudo de la celda TRIBUNAL (`trib_raw`), no el nombre normalizado — puede verse inconsistente entre filas si el Excel trae variantes de escritura.

**Decisión:** usar el nombre de tribunal normalizado y prolijo (Laja/Mulchén/Tomé) en asunto y cuerpo, igual que en otras partes del sistema.

---

# 2. Otros módulos — validador, resoluciones, GUI, procesador

## 2.1 · `detectar_tipo()` no reconoce ningún texto actual — parche rápido

**Motivo:** `resoluciones/generador_resoluciones.py` decide qué tipo de resolución generar (PC_IE, PC_INFO, NOMENCL) buscando frases fijas dentro del texto de OBSERVACION. Verificado carácter por carácter: **ninguna** de las frases buscadas coincide con los textos actualmente aprobados (ni siquiera antes de los cambios de hoy — venía así desde el documento base). Resultado: `generar_resoluciones()` nunca detecta nada, siempre "0 resoluciones generadas".

**Decisión:** parche rápido — actualizar las frases de búsqueda al texto real, mismo mecanismo (frágil, pero funcionando). **NOMENCL se deja intacto, sin actualizar** — no corresponde a ninguna regla actual del catálogo (podría mapear a los tipos 1/2/4/6/7/8 del listado de proyectos de resolución del Manual, que hoy no se derivan de ninguna columna del Excel) y el usuario prefirió no tocarlo por ahora.

**Patrones sugeridos para implementación:**
```python
_PATRON_PC_IE = [
    "proyecto de resolucion pidiendo cuenta al programa respecto del ingreso efectivo"
]
_PATRON_PC_INFO = [
    "que se encuentra vencido en rus desde el"
]
_PATRON_NOMENCL = [  # SIN CAMBIOS — se deja igual que hoy, dormant
    "aplica nomenclaturas", "regularizar informaticamente",
    "nomenclaturas a fin de regularizar"
]
```
**Se descartó** la alternativa de fondo (identificador de regla por fila en vez de parseo de texto) y también usar la columna `RES` (manual) como filtro — el usuario prefirió el parche de frases.

## 2.2 · Advertencia visible en GUI cuando falta Hoja2

**Motivo:** hoy, si Hoja2 no carga, el sistema solo registra `"ℹ️ Sin Hoja2 — cruce desactivado"` en el log de texto — al mismo nivel que cualquier aviso menor. Como C-06 fue eliminada (§0 de este documento y §5 del catálogo de reglas), Hoja2 ausente significa **cero** observaciones de "próximo informe" en todo Cumplimiento ese día, sin aviso fuerte.

**Decisión:** subir esto a una advertencia visible y prominente en la GUI (no solo el log de texto plano) cuando se procese CUMPLIMIENTO sin Hoja2 disponible.

## 2.3 · Extender la unificación de formato a `resoluciones/` — **rechazada**

`resoluciones/generador_resoluciones.py` tiene su propia tercera copia de `_titulo_programa`, `normalizar` y `detectar_tribunal` (además de motor y comunicaciones). El usuario decidió **no** incluir este módulo en la unificación de §1.9 — mantiene su copia independiente.

## 2.4 · Agregar FECHA DE NACIMIENTO al chequeo de campos vacíos masivos (A7) — **rechazada**

A7 revisa completitud de EDAD para Espera, pero no de FECHA DE NACIMIENTO (la fuente real de las reglas de mayoría de edad). Se decidió no agregarlo.

## 2.5 · Ampliar alias de RIT en el motor — **rechazada**

El validador reconoce más variantes de nombre de columna para RIT ("N° RIT", "N RIT", etc.) que el motor (solo "RIT"). Se decidió no ampliar — el usuario confirma que en sus Excel reales siempre viene como "RIT".

## 2.6 · Eliminar código muerto en el validador

**Motivo:** `COLS_MINIMAS` y `COLS_CRITICAS_VACIAS`, declarados al inicio de `validador/reglas_validacion.py`, no los usa ninguna función del archivo (cada regla arma su propia lista de columnas por dentro).

**Decisión:** eliminarlos. Sin efecto en comportamiento.

## 2.7 · Bajar el techo de edad razonable en A6 (0-25) — **rechazada**

Se evaluó bajar el techo a ~19-20 (dado que mayoría de edad son 18 años). Se decidió dejar 25 sin cambios.

## 2.8 · `TEXTO_CORRECTO` no confirma automáticamente el pendiente — **rechazada**

Se evaluó que escribir una redacción corregida en `TEXTO_CORRECTO` (flujo S4, confirmación de textos) marcara `confirmado=true` automáticamente, sin necesitar además `APROBADO=SI`. El usuario prefirió mantener los dos pasos separados, tal como está hoy.

## 2.9 · Aviso de marcadores `{...}` no reconocidos al confirmar textos — **rechazada**

Se evaluó validar que los marcadores escritos en `TEXTO_CORRECTO` coincidan con las variables reales que usa cada regla (para detectar typos como `{FECHA_MAYORA}` en vez de `{FECHA_MAYORÍA}`). Rechazada — no se agrega.

## 2.10 · Separar el campo de Excel de entrada en la pestaña de correos

**Motivo:** un solo campo "Archivo Excel" compartido entre los botones "Informes vencidos/por vencer" (necesita el Excel de INFORMES) y "Lista de espera" (necesita el Excel de ESPERA ya procesado por el motor) — son estructuras de columnas distintas.

**Decisión:** dos campos de entrada independientes, uno por botón/tipo de correo.

## 2.11 · Muestra de la vista previa (primeras 30 filas) — **rechazada**

Se evaluó que la vista previa (S3) mostrara una muestra representativa (variedad de tribunales/reglas) en vez de literalmente `df.head(30)`. Se decidió dejarlo como está.

## 2.12 · Índice de Hoja2 descarta fechas de hoy/pasadas — **confirmado sin cambio**

`_construir_indice_hoja2` excluye del cruce cualquier fila de Hoja2 cuya fecha de vencimiento sea hoy o anterior (`if vd <= hoy: continue`). Se evaluó si esto debía ampliarse dado que C-06 ya no existe como respaldo. El usuario confirma que está bien así — los informes ya vencidos se cubren por otra vía (fuera del alcance de Hoja2/C-10).

## 2.13 · Corrección de la supresión Hoja2 vs. C-04/C-05

Ver §0 al inicio de este documento — es la corrección más importante de esta sección: el código real ya suprime Hoja2 con medida vencida/por vencer, y esa es la conducta que se confirma como correcta (revirtiendo B3 de la ronda de auditoría de reglas).

---

# 3. Resumen de aceptación/rechazo

| # | Ítem | Módulo | Resultado |
|---|---|---|---|
| 1.1 | Eliminar correo tribunal por vencidos | correos | ✅ Aceptada |
| 1.2 | Corregir PATRON_ESPERA | correos | ✅ Aceptada |
| 1.3 | Distinguir DCE en plantillas | correos | ✅ Aceptada |
| 1.4 | Unificar formato nombre programa | correos | ✅ Aceptada (parte de 1.9) |
| 1.5 | Saludo personalizado director | correos | ❌ Rechazada |
| 1.6 | DIAS_POR_VENCER 45→30 | correos | ✅ Aceptada |
| 1.7 | Corregir día 0 vencido/por vencer | correos | ✅ Aceptada |
| 1.8 | Alias RUT motor/correos | correos | ✅ Aceptada (parte de 1.9) |
| 1.9 | Unificar columnas y funciones comunes | correos+motor | ✅ Aceptada |
| 1.10 | Normalizar tribunal en asunto | correos | ✅ Aceptada |
| 2.1 | Parche detectar_tipo() | resoluciones | ✅ Aceptada (parche, no fix de raíz) |
| 2.2 | Advertencia visible Hoja2 ausente | GUI/procesador | ✅ Aceptada |
| 2.3 | Extender unificación a resoluciones | resoluciones | ❌ Rechazada |
| 2.4 | FECHA NACIMIENTO en A7 | validador | ❌ Rechazada |
| 2.5 | Ampliar alias RIT en motor | motor | ❌ Rechazada |
| 2.6 | Eliminar código muerto validador | validador | ✅ Aceptada |
| 2.7 | Bajar techo edad A6 | validador | ❌ Rechazada |
| 2.8 | TEXTO_CORRECTO auto-confirma | textos_confirmacion | ❌ Rechazada |
| 2.9 | Aviso de marcadores no reconocidos | textos_confirmacion | ❌ Rechazada |
| 2.10 | Separar campo Excel en correos | GUI | ✅ Aceptada |
| 2.11 | Vista previa representativa | GUI | ❌ Rechazada |
| 2.12 | Índice Hoja2 descarta hoy/pasadas | procesador | Sin cambio (confirmado) |
| 2.13 | Supresión Hoja2 vs C-04/C-05 | procesador | Corrección — revierte B3 |

**Total sesión (rondas 3 y 4):** 23 ítems evaluados, 14 aceptados, 8 rechazados, 1 corrección a decisión previa.
