# CATÁLOGO DE REGLAS CSMP — v2

**Reglas consolidadas para generar observaciones · Espera · Cumplimiento · Informes**

**Centro de Seguimiento de Medidas de Protección · Concepción**
**Fecha:** 14 de julio de 2026
**Estado:** Especificación funcional definitiva. Reemplaza a `Reglas_Consolidadas_Observaciones_CSMP_2026-07-14.docx` como fuente única de verdad. Es el insumo directo para implementación en CSMP Assistant.

> **Relación con el documento base:** este catálogo incorpora las 43 decisiones del documento base del 14 de julio de 2026, las 13 correcciones de la auditoría posterior (contraste contra el Manual de Funciones CSMP y el código v8.14), y 11 propuestas adicionales evaluadas en la misma sesión. Donde no se indica cambio, la regla es idéntica al documento base. §12 trae el mapa completo de qué cambió y por qué.

---

# 1. Objeto y alcance

Especificación funcional para generar observaciones administrativas a partir de las planillas de Espera, Cumplimiento e Informes, y para estructurar las columnas de apoyo al registro de gestión (fecha de observación, totales, carga, proyectos de resolución).

## 1.1 Principios operativos

- Las observaciones se construyen únicamente con datos disponibles en las planillas.
- La derivación no sujeta a seguimiento es una regla de corte.
- Curador, oído, próxima audiencia y fichas son reglas complementarias y no deben suprimir la observación principal.
- Las reglas que requieren una fecha sólo se generan cuando existe un valor válido.
- Una regla incompleta se omite sin eliminar las demás observaciones de la fila.
- Las frases acumuladas se separan con punto y espacio; la salida termina con un solo punto.
- La clasificación con/sin carga y el registro de proyectos de resolución **no se calculan automáticamente** — se completan a mano en columnas dedicadas del Excel de salida (§8). El sistema no infiere carga/sin carga.

## 1.2 Fuentes revisadas

- Documento base `Reglas_Consolidadas_Observaciones_CSMP_2026-07-14.docx` (43 decisiones).
- Auditoría de ese documento contra el Manual de Funciones CSMP 2025 y el código real de `CSMP_Assistant_v8_14_COMPLETO.zip` (13 decisiones adicionales).
- 11 propuestas de mejora/simplificación evaluadas en la misma sesión, contrastadas contra columnas reales mapeadas en el código.

---

# 2. Arquitectura de aplicación

1. Leer y normalizar los campos necesarios de la fila.
2. Determinar el tribunal exclusivamente desde la columna TRIBUNAL. Si no se reconoce (no es LAJA/MULCHEN/TOME), **no se aplica ningún valor por defecto**: las reglas tribunal-dependientes (hoy, solo E-05) se omiten para esa fila, y la fila queda marcada en el reporte de validación (ya existe como regla A8 en `validador/reglas_validacion.py`).
3. Aplicar la regla de no seguimiento (E-01); si opera, terminar la evaluación de la fila.
4. Evaluar las reglas principales del módulo, en el orden técnico de §9.
5. Agregar las reglas complementarias que correspondan (curador, oído, fichas, próxima audiencia — todas se tratan como complementarias para efectos de composición, §5 C-09).
6. Formatear programa, fechas, prefijo y puntuación.
7. Omitir fragmentos incompletos y registrar su incidencia en la validación; nunca mostrar marcadores tipo `{DATO}` sin resolver.
8. Completar `FECHA_OBS` automáticamente. Dejar `TT`, `CC`, `RES` vacías para llenado manual.

---

# 3. Reglas de Espera

## E-01 · Derivación no sujeta a seguimiento

**Datos:** DERIVACIÓN o NOMBRE DEL CENTRO.

**Condición:** Coincidencia controlada con OPD, DAM, SALUD PRIVADA, HOSPITAL, UNIDAD DE SALUD MENTAL, CESFAM, RED SALUD, CONSULTA EXTERNA, COLEGIO o CHILE CRECE CONTIGO.

**Aplicación:** Regla de corte. Se aplica antes que edad, espera, curador, oído y audiencia. Rige también en Cumplimiento e Informes.

**Implementación (corrección D1):** el match debe ser cuidadoso — palabra completa, anclada al inicio del nombre del programa — no una coincidencia de subcadena libre. Se evaluó específicamente el riesgo de colisión con DCE (que sí es una derivación seguida activamente) y se confirmó que no ocurre en la práctica: no se requiere exclusión explícita de DCE en esta regla.

**Texto aprobado:**

> *{PRIMER_NOMBRE} {SIGLA}: La derivación {PROGRAMA} no se encuentra sujeta a seguimiento por este Centro.*

**Notas:** texto deliberadamente distinto del literal del Manual ("Este Centro no realiza seguimiento a Red Salud y Educación") — se optó por una redacción neutral y generalizable a todas las derivaciones del catálogo, no solo Red Salud.

---

## E-02 · Persona mayor de edad

**Datos:** FECHA DE NACIMIENTO. Confirmado: este dato nunca falta cuando la persona ya es mayor de edad — no se define texto alternativo ni respaldo vía columna EDAD.

**Condición:** Edad exacta igual o superior a 18 años.

**Aplicación:** Observación principal. No agrega las reglas principales de espera, pero permite acumular curador, oído y próxima audiencia.

**Texto aprobado:**

> *{PRIMER_NOMBRE} {SIGLA}: Se hace presente que {PRIMER_NOMBRE} alcanzó la mayoría de edad el {FECHA_MAYORÍA}, se sugiere egresar la medida.*

---

## E-03 · Próxima mayoría de edad

**Datos:** FECHA DE NACIMIENTO.

**Condición:** Faltan entre 1 y 60 días, ambos inclusive, para cumplir 18 años.

**Aplicación:** Regla acumulativa.

**Texto aprobado:**

> *Se hace presente que {PRIMER_NOMBRE} alcanzará la mayoría de edad el {FECHA_MAYORÍA}.*

---

## E-04 · Resolución reciente que ordena ingreso efectivo

**Datos:** FECHA DE RESOLUCIÓN y PROGRAMA.

**Condición:** Han transcurrido entre 0 y 29 días desde FECHA DE RESOLUCIÓN hasta la ejecución.

**Aplicación:** Observación base para los tres tribunales. Admite todas las reglas acumulativas. Si también se cumple E-05, ambas se agregan.

**Texto aprobado:**

> *{PRIMER_NOMBRE} {SIGLA}: Medida revisada, a la espera de ingreso efectivo. Se hace presente que el Tribunal ordenó el ingreso efectivo al programa {PROGRAMA} con fecha {FECHA_RESOLUCIÓN}. (Observación administrativa, no requiere acción/respuesta del Tribunal).*

---

## E-05 · Espera de 30 días o más — **rediseñada (auditoría, decisiones A1/A2)**

**Datos:** T ESPERA, TRIBUNAL, PROGRAMA (para detectar DCE).

**Condición y ramas — reemplaza íntegramente la versión del documento base:**

**Rama DCE (cualquier tribunal):** T ESPERA ≥ 30 días. **Nunca** genera proyecto de resolución — el Manual lo prohíbe expresamente para DCE. Solo correo.

> *{PRIMER_NOMBRE} {SIGLA}: Medida revisada, a la espera de ingreso efectivo. Se remite correo electrónico al programa consultando respecto de la fecha estimada de ingreso efectivo.*

**Rama Laja y Mulchén (no DCE):** T ESPERA ≥ 30 días. Proyecto de resolución + correo.

> *{PRIMER_NOMBRE} {SIGLA}: Medida revisada, a la espera de ingreso efectivo. Se remite proyecto de resolución pidiendo cuenta al programa respecto del ingreso efectivo. Igualmente, se remite correo electrónico al programa consultando respecto de la fecha estimada de ingreso efectivo.*

**Rama Tomé (no DCE), 30 a 59 días:** solo correo — mismo texto que la rama DCE.

**Rama Tomé (no DCE), 60 días o más:** agrega proyecto de resolución — mismo texto que la rama Laja/Mulchén.

**Notas de la corrección:**
- El documento base estandarizaba 30 días para los tres tribunales sin distinguir DCE ni el tramo superior de Tomé. Confirmado con el usuario: Laja y Mulchén cambian de 30 días → proyecto+correo (sin tramos); Tomé mantiene un esquema de dos tramos (30d correo, 60d agrega proyecto); DCE nunca lleva proyecto, umbral único de 30 días sin distinción de tribunal.
- Cierre unificado a "ingreso efectivo" en los tres tribunales (antes Tomé decía solo "ingreso" — era una discrepancia de redacción no deliberada).

---

## E-06 · Espera menor de 30 días sin fecha de resolución

**Datos:** T ESPERA, PROGRAMA y FECHA DE RESOLUCIÓN.

**Condición:** T ESPERA menor de 30 días y ausencia de una fecha válida de resolución.

**Aplicación:** Observación base uniforme para DCE y los demás programas. Confirmado: DCE no lleva mención de correo en este tramo — la obligación de correo "sin importar los días" del Manual no está operativa en la práctica; DCE sigue el mismo umbral de 30 días que los demás para efectos de correo (ver E-05).

**Texto aprobado:**

> *{PRIMER_NOMBRE} {SIGLA}: Medida revisada, a la espera de ingreso efectivo al programa {PROGRAMA}.*

**Nota de cobertura:** se evaluó un posible vacío (espera corta + resolución antigua sin capturar por ninguna regla). Confirmado con el usuario que esa combinación no ocurre en la práctica — sin cambios.

---

# 4. Reglas complementarias comunes

## T-01 · Curador no registrado

**Datos:** CURADOR.

**Condición:** La columna existe y su contenido no contiene un RUT ni la expresión «Institución:».

**Aplicación:** Acumulativa en Espera y Cumplimiento. Nunca impide la observación principal. No se aplica si falta completamente la columna ni si opera E-01.

**Con/sin carga:** el Manual distingue primera vez (con carga) de repeticiones tras respuesta del tribunal (sin carga). Esta distinción **no se resuelve en el texto ni en el código** — se completa a mano en la columna `CC` del Excel de salida (§8), donde el usuario decide caso a caso si corresponde con carga.

**Texto aprobado:**

> *No registra curador asociado en RUS, se sugiere asociar curador ad litem informáticamente.*

---

## T-02 · Oído reciente

**Datos:** FECHA OÍDO.

**Condición:** Fecha válida entre el día de ejecución y 45 días hacia atrás, ambos inclusive; nunca futura.

**Aplicación:** Acumulativa con cualquier observación principal, incluso «Medida revisada...» y mayoría de edad.

**Texto aprobado:**

> *Oído con fecha {FECHA_OÍDO}.*

---

## T-03 · Próxima audiencia

**Datos:** PRÓXIMA AUDIENCIA.

**Condición:** Fecha válida igual o posterior al día de ejecución, sin límite máximo futuro.

**Aplicación:** Acumulativa en Espera, Cumplimiento e Informes; no se aplica si opera la regla de no seguimiento.

**Nota de disponibilidad de dato:** no está confirmado si el Excel de INFORMES trae esta columna. No es prioritario resolverlo: si la columna no existe en ese export, la regla simplemente no se activa ahí, sin error.

**Texto aprobado:**

> *Se cita a audiencia para el día {FECHA_AUDIENCIA}.*

---

# 5. Reglas de Cumplimiento

## C-01 · Persona mayor de edad

Idéntica a E-02.

> *{PRIMER_NOMBRE} {SIGLA}: Se hace presente que {PRIMER_NOMBRE} alcanzó la mayoría de edad el {FECHA_MAYORÍA}, se sugiere egresar la medida.*

## C-02 · Próxima mayoría de edad

Idéntica a E-03.

> *Se hace presente que {PRIMER_NOMBRE} alcanzará la mayoría de edad el {FECHA_MAYORÍA}.*

## C-03 · Ingreso efectivo reciente — **modificada (propuesta 7)**

**Datos:** DÍAS DE CUMPLIMIENTO (fuente de la condición), FECHA DE INGRESO EFECTIVO (solo para mostrar la fecha en el texto), PROGRAMA.

**Condición:** DÍAS DE CUMPLIMIENTO entre 0 y 30, ambos inclusive.

**Cambio respecto al documento base:** la condición ya no se calcula restando fecha de hoy menos fecha de ingreso efectivo — usa directamente la columna DÍAS DE CUMPLIMIENTO (confirmado: esa columna ya representa exactamente "días desde el ingreso efectivo"). Evita cualquier desfase entre lo que el sistema calcula y lo que RUS muestra.

**Texto aprobado (sin cambio):**

> *Medida revisada. Se hace presente que el ingreso efectivo al programa {PROGRAMA} se registra con fecha {FECHA_INGRESO}.*

## C-04 · Medida vencida

**Datos:** DÍAS PARA EGRESAR, DÍAS DE CUMPLIMIENTO y FECHA DE EGRESO PROYECTADO.

**Condición:** DÍAS PARA EGRESAR es menor que 0 o DÍAS DE CUMPLIMIENTO es menor que 0; además existe una fecha válida de egreso proyectado.

**Aplicación:** Acumulativa, pero excluye C-06 (informe calculado — hoy eliminada, ver más abajo). Se evaluó si DÍAS DE CUMPLIMIENTO negativo podía generar un texto incoherente ("vencida desde una fecha futura"); confirmado que no ocurre en la práctica — se mantiene la condición OR sin cambios, como red de seguridad inofensiva.

**Texto aprobado:**

> *La medida se visualiza vencida en RUS desde el {FECHA_EGRESO_PROYECTADO}.*

## C-05 · Medida próxima a vencer o con vencimiento hoy

**Datos:** DÍAS PARA EGRESAR y FECHA DE EGRESO PROYECTADO.

**Condición:** Fecha válida y DÍAS PARA EGRESAR entre 0 y 45, ambos inclusive.

**Texto aprobado — Día 0:**

> *Se hace presente que la medida se visualiza con vencimiento para el día de hoy, {FECHA_EGRESO_PROYECTADO}.*

**Texto aprobado — Días 1 a 45:**

> *Se hace presente que la medida se visualiza próxima a vencer en RUS el {FECHA_EGRESO_PROYECTADO}.*

## ~~C-06 · Próximo informe de avance calculado~~ — **ELIMINADA**

Se elimina por completo el cálculo interno de hitos mensuales (3/6/9/12/15/18 para Laja-Mulchén; 4/8/12/16/20 para Tomé). Decisión explícita del usuario: se usa **siempre y únicamente** el cruce con Hoja2 (ver C-10). Se confirmó que Hoja2 estará siempre disponible al procesar Cumplimiento. Toda la lógica de meses-hito, `relativedelta` y excepción de coincidencia con egreso proyectado se elimina del código.

## C-07 · Ficha Individual (Residencial) — **fusión de C-07 + C-08 del documento base**

**Corrección crítica de la auditoría:** el documento base trataba "ficha residencial" (columna `FEC.ACT.F.RESIDENCIAL`) y "ficha individual" (columna `FEC.ACT.F.INDIVIDUAL`) como dos fichas distintas, con dos reglas separadas. Son **el mismo documento**. El listado oficial de nomenclaturas del Manual solo reconoce tres fichas: Individual, Ambulatoria y FAE — no existe "ficha residencial" como documento aparte. Además, la columna `FEC.ACT.F.RESIDENCIAL` no existe en los Excel reales — solo existe `FEC.ACT.F.INDIVIDUAL`. Se elimina toda referencia a la columna residencial.

**Datos:** PROGRAMA y FEC.ACT.F.INDIVIDUAL.

**Condición:** Programa con prefijo RTA, RTT, RES, RFA o RVA.

**Aplicación:** Tres ramas acumulativas — ausencia; antigüedad superior a 180 días; actualización entre 0 y 30 días. Entre 31 y 180 días no genera frase. Trata la observación de esta regla como **complementaria** para efectos de C-09 (cierre de Cumplimiento).

**Texto — sin fecha:**

> *No registra ficha individual en RUS, se sugiere confeccionar.*

**Texto — más de 180 días:**

> *Atendido que la ficha individual registra como última actualización el {FECHA_FICHA_INDIVIDUAL}, superando los 180 días, se sugiere actualizar.*

**Texto — entre 0 y 30 días:**

> *Se hace presente que la ficha individual fue actualizada con fecha {FECHA_FICHA_INDIVIDUAL}.*

## C-08 · Ficha FAE ausente *(antes C-09 en el documento base)*

**Datos:** PROGRAMA, FECHA DE INGRESO EFECTIVO y FECHA DE FICHA FAE.

**Condición:** Programa contiene FAE o FAS; existe la columna de ficha; han transcurrido más de 120 días desde ingreso efectivo; no existe fecha de ficha.

**Aplicación:** Se evaluó y se **descarta** agregar ramas de "ficha FAE desactualizada" o "recién actualizada" (paralelas a las de C-07) — el Manual solo exige verificar ausencia para fichas no residenciales, no antigüedad. Se conserva exactamente la lógica de una sola rama. Se ajusta el texto para quitarle el prefijo "Medida revisada," que traía duplicado (chocaba con el prefijo que antepone C-09 cuando esta es la única regla complementaria activa).

**Texto aprobado (modificado):**

> *Se hace presente que {PRIMER_NOMBRE} no tiene ficha FAE, se sugiere confeccionar.*

**Ficha Ambulatoria: confirmado definitivamente que no se agrega ninguna regla — decisión explícita y final, no es una omisión.**

## C-09 · Cumplimiento sin observaciones principales *(antes C-10)*

**Datos:** Resultado de las demás reglas.

**Condición:** No se activa ninguna regla principal.

**Aplicación:** Si no existe ninguna regla, usa el cierre completo. Si sólo hay curador, oído, fichas (C-07, C-08) o próxima audiencia, usa la base breve y agrega esos fragmentos — **todas estas se tratan como complementarias**, incluidas las fichas.

**Texto — ninguna regla:**

> *{PRIMER_NOMBRE} {SIGLA}: Medida revisada, sin observaciones.*

**Texto — sólo complementarias:**

> *{PRIMER_NOMBRE} {SIGLA}: Medida revisada. {REGLAS_COMPLEMENTARIAS}*

## C-10 · Próximo informe de avance (Hoja2) *(antes C-11 — ahora fuente única)*

**Datos:** Cruce por RIT, RUT, NOMBRE, TRIBUNAL y DERIVACIÓN; FECHA DE VENCIMIENTO de Hoja2.

**Condición:** Coincidencia según la normalización actual de CSMP Assistant.

**Aplicación:** Ya no "sustituye" a un cálculo interno — es la **única** fuente de este dato (C-06 fue eliminada). **Se suprime cuando ya disparó C-04 (medida vencida) o C-05 (próxima a vencer)** — misma conducta que tenía el antiguo C-06 con el cálculo interno. *(Corrección post-revisión de código: la decisión original de esta sesión fue "no se suprime, revisión humana caso a caso"; al revisar el código real se confirmó que la supresión sí corresponde, y se revirtió — ver `Mejoras_Correos_y_Otros_Modulos_2026-07-14.md` §0/§2.13.)* Si hay duplicados en Hoja2 se conserva la última fila encontrada.

**Texto aprobado:**

> *Medida revisada, se hace presente que el programa {PROGRAMA} deberá remitir informe de avance a más tardar el {FECHA_VENCIMIENTO}.*

---

# 6. Reglas de Informes

> **Regla previa:** I-01 e I-02 sólo se evalúan si la derivación no quedó excluida por E-01.

## I-01 · Informe vencido

**Datos:** FECHA DE VENCIMIENTO y PROGRAMA.

**Condición:** FECHA DE VENCIMIENTO anterior al día de ejecución.

**Texto — DCE:**

> *Se remite correo electrónico al programa {PROGRAMA} a fin de requerir el informe diagnóstico que se encuentra vencido en RUS desde el {FECHA_VENCIMIENTO}.*

**Texto — otros:**

> *Se remite correo electrónico al programa {PROGRAMA} a fin de requerir el informe de avance que se encuentra vencido en RUS desde el {FECHA_VENCIMIENTO}.*

## I-02 · Informe por vencer

**Datos:** FECHA DE VENCIMIENTO y PROGRAMA.

**Condición:** Faltan entre 0 y 30 días para el vencimiento, ambos inclusive.

**Texto — DCE:**

> *Se remite correo electrónico al programa {PROGRAMA} a fin de señalar que el informe diagnóstico ordenado en autos debe ser remitido a más tardar el {FECHA_VENCIMIENTO}.*

**Texto — otros:**

> *Se remite correo electrónico al programa {PROGRAMA} a fin de señalar que el próximo informe de avance vence el {FECHA_VENCIMIENTO}.*

---

# 7. Reglas transversales de formato y datos

| ID | Materia | Regla aprobada |
|---|---|---|
| G-01 | Prefijo | `{PRIMER_NOMBRE} {SIGLA}:`; la sigla usa las tres primeras letras de la derivación. |
| G-02 | Fechas | Formato largo: 14 de julio de 2026. |
| G-03 | Programa | Formato automático: siglas en mayúsculas, preposiciones en minúsculas, resto capitalizado. Se evaluó reemplazar el criterio actual (palabras de ≤4 letras = sigla) por una lista controlada de siglas reales; **se rechaza el cambio** — el criterio de largo funciona bien en la práctica. |
| G-04 | Puntuación | Punto y espacio entre fragmentos; un solo punto final; sin puntos duplicados. |
| G-05 | Dato indispensable ausente | Se omite sólo la regla incompleta, se conservan las demás y la fila se registra en validación; nunca se muestran marcadores. |
| G-06 | Tribunal | Se obtiene exclusivamente desde la columna TRIBUNAL; se normalizan tildes, mayúsculas y espacios; **no hay valor predeterminado**. Sin tribunal reconocible: se omiten las reglas tribunal-dependientes (hoy, solo E-05) y la fila se marca en validación (regla A8 ya existente). |
| G-07 | Orden | Se conserva el orden técnico de §9, actualizado tras eliminar C-06 y fusionar fichas. |
| G-08 | Hoja2 | Fuente única y permanente del próximo informe en Cumplimiento (ya no "sustituye" un cálculo — el cálculo no existe). Se conservan la normalización actual y la última fila cuando existen duplicados. |

---

# 8. Columnas de salida del Excel — **nuevo**

Se agregan cuatro columnas al Excel de salida, de llenado no automático salvo la primera:

| Columna | Posición | Contenido | Llenado |
|---|---|---|---|
| `FECHA_OBS` | Antes de OBSERVACION | Fecha de procesamiento del Excel | **Automático** — fecha del día en que se corre el motor |
| `TT` | Después de OBSERVACION | Observaciones totales (1 si la fila tiene observación, 0 si no) | Manual |
| `CC` | Después de OBSERVACION | Observaciones con carga (1/0) | Manual |
| `RES` | Después de OBSERVACION | Proyecto de resolución generado (vacío, 0 o 1) | Manual |

**Aritmética de reporte (para el informe de gestión trimestral, §VII.d del Manual):**
- `TT` se suma entre sí → total de observaciones realizadas.
- `CC` se suma entre sí → total de observaciones con carga.
- `TT − CC` = observaciones sin carga realizadas.
- `RES` no se suma — es un marcador por caso, no una métrica agregable en esta etapa.

El sistema **no calcula** con/sin carga por regla — eso quedó descartado explícitamente (ver §11). El usuario completa `TT`, `CC` y `RES` durante su revisión humana de cada fila.

**Pendiente agendado (no ahora):** herramienta separada que lea el Excel ya completado a mano y tabule `TT`/`CC`/`RES` por tribunal, como insumo directo del informe trimestral. Registrado en memoria para retomar cuando corresponda.

---

# 9. Orden técnico

## 9.1 Espera

| Orden | Regla |
|---|---|
| 1 | Próxima mayoría de edad (E-03). |
| 2 | Regla principal: E-01 (corte) → E-02 (mayoría cumplida) → E-04/E-05/E-06 (resolución reciente, demora por tribunal/DCE, o fallback). |
| 3 | Curador (T-01). |
| 4 | Oído (T-02). |
| 5 | Próxima audiencia (T-03). |

**Excepciones:** la derivación sin seguimiento corta todo; la mayoría de edad actúa como observación principal y permite reglas complementarias.

## 9.2 Cumplimiento

| Orden | Regla |
|---|---|
| 1 | Próxima mayoría de edad (C-02). |
| 2 | Ingreso efectivo reciente (C-03) — condición vía DÍAS DE CUMPLIMIENTO. |
| 3 | Medida vencida (C-04). |
| 4 | Medida próxima a vencer o con vencimiento hoy (C-05). |
| 5 | Próximo informe (C-10, Hoja2 — única fuente). |
| 6 | Curador (T-01). |
| 7 | Oído (T-02). |
| 8 | Ficha Individual/Residencial (C-07, fusionada). |
| 9 | Ficha FAE (C-08). |
| 10 | Próxima audiencia (T-03). |

**Nota:** el orden pasó de 11 pasos a 10 — se fusionaron dos fichas en una y se eliminó el paso de cálculo interno de informe (C-06).

---

# 10. Ejemplos de composición

**10.1 · Espera, Tomé, 35 días (solo correo, tramo bajo)**

> *Camila PIE: Medida revisada, a la espera de ingreso efectivo. Se remite correo electrónico al programa consultando respecto de la fecha estimada de ingreso efectivo.*

**10.2 · Espera, Tomé, 65 días (agrega proyecto, tramo alto)**

> *Camila PIE: Medida revisada, a la espera de ingreso efectivo. Se remite proyecto de resolución pidiendo cuenta al programa respecto del ingreso efectivo. Igualmente, se remite correo electrónico al programa consultando respecto de la fecha estimada de ingreso efectivo. No registra curador asociado en RUS, se sugiere asociar curador ad litem informáticamente. Oído con fecha 8 de julio de 2026.*

**10.3 · Espera, DCE, 40 días, cualquier tribunal (nunca proyecto)**

> *Martín DCE: Medida revisada, a la espera de ingreso efectivo. Se remite correo electrónico al programa consultando respecto de la fecha estimada de ingreso efectivo.*

**10.4 · Cumplimiento, residencial (ficha individual fusionada + Hoja2)**

> *Daniela RES: Medida revisada. Se hace presente que el ingreso efectivo al programa RES Nuevo Amanecer se registra con fecha 1 de julio de 2026. Se hace presente que la medida se visualiza próxima a vencer en RUS el 20 de agosto de 2026. Atendido que la ficha individual registra como última actualización el 5 de enero de 2026, superando los 180 días, se sugiere actualizar. Oído con fecha 10 de julio de 2026.*

**10.5 · Cumplimiento, solo ficha FAE ausente (ilustra fusión con cierre C-09)**

> *Ignacio FAE: Medida revisada. Se hace presente que Ignacio no tiene ficha FAE, se sugiere confeccionar.*

**10.6 · Informe, DCE vencido**

> *Martín DCE: Se remite correo electrónico al programa DCE Tomé a fin de requerir el informe diagnóstico que se encuentra vencido en RUS desde el 30 de junio de 2026.*

---

# 11. Reglas y propuestas no incorporadas

| Materia | Decisión final |
|---|---|
| Espera superior a 180 días | No crea observación adicional; continúa E-05. |
| Fecha de resolución futura | No se incorpora; supuesto inexistente. |
| Fecha de ingreso efectivo futura | No se incorpora; supuesto inexistente. |
| Egreso proyectado anterior al ingreso | No se incorpora; supuesto inexistente. |
| Ficha FAE antigua/reciente | No se incorpora; se conserva únicamente la rama de ausencia. Reafirmado tras revisión: el Manual solo exige chequeo de ausencia para fichas no residenciales. |
| Correo al Tribunal por medida vencida/por vencer | No se incorpora; sólo se informa el estado en RUS. |
| Hitos ambulatorios 21/24 (Cumplimiento) | Sin efecto — el cálculo de hitos que los contenía (C-06) fue eliminado por completo. |
| Mes calendario de Residencia Tomé | Sin efecto, mismo motivo. |
| Informe sin fecha de vencimiento | No se define observación; supuesto inexistente. |
| Persona adulta sin fecha de nacimiento | No se define texto alternativo; confirmado que nunca ocurre en la práctica. |
| Conflicto vencida/próxima a vencer | No se define precedencia adicional; supuesto inexistente. |
| Mejora de ficha FAE basada en DÍAS DE CUMPLIMIENTO | No se adopta; se conserva el cálculo desde ingreso efectivo. |
| Duplicidad de derivación por RUT (validador) | Evaluada y **rechazada** — no es necesaria. |
| Columna CARGA/SIN CARGA calculada automáticamente por regla | Evaluada y **rechazada** — reemplazada por columnas `TT`/`CC`/`RES` de llenado manual (§8). |
| Texto distinto para curador "primera vez" vs "repetido" | Evaluado y **rechazado** — se resuelve con la columna `CC` manual, no con lógica ni texto. |
| Ficha residencial y FAE con rama "recién actualizada" | Ficha residencial resultó ser la misma ficha individual (ya tenía esa rama, fusión hecha en C-07). Para FAE, **rechazada** — el Manual no exige ese chequeo ahí. |
| Ficha Ambulatoria (regla nueva) | **Rechazada, definitiva** — nunca se agrega, confirmado explícitamente dos veces. |
| Externalizar umbrales numéricos a archivo de configuración | Evaluada y **rechazada** — quedan en el código por ahora. |
| Herramienta de conteo trimestral de TT/CC | **No ahora** — agendada para más adelante, registrada en memoria de la sesión. |
| Exclusión explícita de DCE en el corte de E-01 | Evaluada y **rechazada** — no es un riesgo real según el usuario. |
| Lista controlada de siglas para formato de programa (G-03) | Evaluada y **rechazada** — el criterio de largo de palabra funciona bien en la práctica. |
| Ficha Ambulatoria como validador/regla general (E1 original) | Superada por la corrección de fichas — confirmado definitivamente fuera de alcance. |
| Orden de búsqueda (E2 original) | Sigue fuera de alcance — las planillas actuales (Espera/Cumplimiento/Informes) no cubren OB. |

---

# 12. Mapa de renumeración y cambios respecto al documento base

| ID nuevo | ID en documento base | Cambio |
|---|---|---|
| E-01 a E-04 | E-01 a E-04 | Sin cambio de contenido; E-01 con nota de implementación (match anclado). |
| E-05 | E-05 | Rediseñada: rama DCE nueva (sin proyecto), Tomé pasa de un tramo a dos (30/60), cierre unificado a "ingreso efectivo". |
| E-06 | E-06 | Sin cambio. |
| T-01 a T-03 | T-01 a T-03 | Sin cambio de texto; T-01 con nota sobre `CC` manual. |
| C-01 a C-02 | C-01 a C-02 | Sin cambio. |
| C-03 | C-03 | Condición cambia de "calculado desde fecha" a "columna DÍAS DE CUMPLIMIENTO directa". |
| C-04, C-05 | C-04, C-05 | Sin cambio. |
| — | C-06 | **Eliminada por completo.** |
| C-07 | C-07 + C-08 | **Fusionadas** — eran la misma ficha (Individual/Residencial). Se usa solo la columna `FEC.ACT.F.INDIVIDUAL`. |
| C-08 | C-09 | Renumerada. Texto ajustado (se quita "Medida revisada," duplicado). |
| C-09 | C-10 | Renumerada, sin cambio de contenido. |
| C-10 | C-11 | Renumerada. Ya no "sustituye" un cálculo — es fuente única (C-06 no existe). |
| I-01, I-02 | I-01, I-02 | Sin cambio. |
| G-01, G-02, G-04 a G-08 | ídem | Sin cambio de contenido; G-06 y G-07 con nota de aplicación actualizada. |
| G-03 | G-03 | Evaluado cambio de criterio, rechazado — sin cambio. |
| §8 (columnas de salida) | No existía | **Nuevo.** |

---

# 13. Criterios mínimos de implementación y prueba

- Cada regla debe tener pruebas para el límite inferior, el límite superior y el valor inmediatamente exterior — incluyendo los nuevos límites de E-05 (30/59/60 para Tomé; 30 para DCE y Laja/Mulchén).
- Las salidas deben comprobar que curador, oído y fichas no eliminan la observación principal.
- E-04 y E-05 deben probarse tanto por separado como simultáneamente.
- C-10 (Hoja2) debe probarse en convivencia con C-04/C-05 activos simultáneamente — **se suprime** cuando cualquiera de las dos ya disparó (ver corrección en §5, C-10).
- C-07 (fusión de fichas) debe probarse con las tres ramas (sin fecha, >180d, 0-30d) usando exclusivamente la columna `FEC.ACT.F.INDIVIDUAL`.
- C-03 debe probarse leyendo la condición desde DÍAS DE CUMPLIMIENTO, no desde un cálculo de fecha.
- Las fechas deben renderizarse siempre en español y formato largo.
- No deben quedar marcadores entre llaves, fechas automáticas no justificadas ni puntos duplicados.
- Los programas residenciales admitidos son exclusivamente RTA, RTT, RES, RFA y RVA.
- Confirmar que `FECHA_OBS` se completa automáticamente y que `TT`/`CC`/`RES` quedan vacías (no en 0 por defecto, salvo que se decida lo contrario al implementar).
- Regenerar goldens completos una vez cargados los nuevos textos — los 73 tests de paridad existentes en v8.14 quedan obsoletos por diseño (unificación de puntuación, fusión de fichas, eliminación de C-06).

---

# 14. Incertidumbres operativas abiertas

- **T-03 en Informes:** no está confirmado si el Excel de esa pestaña trae columna de próxima audiencia. No bloquea implementación — si falta, la regla simplemente no se activa ahí.
- **RUT y RIT:** están mapeados y disponibles en el código, pero hoy no se usan en ninguna regla de observación (solo en el cruce interno con Hoja2). Quedan disponibles para uso futuro si surge una necesidad concreta.

---

**Estado:** catálogo funcional cerrado para esta etapa. Cualquier modificación posterior debe registrarse como nueva versión y actualizar simultáneamente condición, prioridad, texto y pruebas — mismo principio que regía el documento base.
