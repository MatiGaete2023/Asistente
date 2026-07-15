# CSMP Assistant — Asistente de RUS

**Versión actual:** v9.0.2 (auditoría de revision3 aplicada — ver docs/auditoria/AUDITORIA_v9.0.1_REVISION3.md)
**Tribunales:** Laja, Mulchén, Tomé

Sistema desktop para CSMP que genera observaciones judiciales automatizadas, borradores de correos Outlook y resoluciones .docx a partir de exportaciones del módulo RUS del Poder Judicial.

## Instalación

Requiere Python 3.12 o superior.

```bash
python -m pip install -r requirements.lock
```

`pywin32` solo en Windows con Outlook instalado.

Para desarrollo y controles adicionales:
```bash
python -m pip install -e ".[dev]"
python -m pytest -v
```

## Ejecución

```bash
python main.py
```

Se abre la GUI con 3 pestañas.
Si se instala el proyecto con `pip install -e .`, también queda disponible el comando
`csmp-asistente-rus`.

## GUI — 3 pestañas

### 1. Motor RUS
- Procesa archivos Excel de RUS en 3 modos: **ESPERA**, **CUMPLIMIENTO**, **INFORMES**.
- Genera Excel de salida con columna `OBSERVACION` calculada por las reglas del manual CSMP.
- Para CUMPLIMIENTO con Hoja2: cruza ambas hojas por 5 campos (RIT+RUT+NOMBRE+DERIV+TRIBUNAL) y usa la fecha de vencimiento de Hoja2.
- **👁 Vista previa (v8.14):** botón por modo que calcula las observaciones EN MEMORIA (sin escribir ningún archivo) y las muestra en una tabla (NNA | Observación, primeras 30 filas + contador total). El texto es idéntico byte-a-byte al que se exportaría. Botón "Exportar…" dentro de la ventana para recién ahí guardar.
- **📋 Confirmación de textos (v8.14):** sección con el contador de textos pendientes de aprobación (`confirmado:false` en `motor/textos_observaciones.json`). "Exportar pendientes…" genera un Excel editable; tras revisarlo y marcar `APROBADO=SI` / completar `TEXTO_CORRECTO`, "Importar confirmaciones…" aplica los cambios al JSON (con backup automático con timestamp) sin tocar ningún archivo `.py`.

### 2. Correos Outlook
- **Informes vencidos / por vencer**: lee Excel de INFORMES (con `FECHA VENCIMIENTO`) y genera un borrador por programa+tribunal y estado.
- **Lista de espera**: lee Excel de ESPERA procesado por el motor (con columna `OBSERVACION`), 1 borrador por programa+tribunal.
- Encabezado: `SEÑORES: / {programa} / PRESENTE.` (PRESENTE en negrita + subrayado).
- Firma del usuario se inyecta via `mail.Display(False)` (Outlook debe estar instalado).

### 3. Resoluciones .docx
- Lee Excel con columna `OBSERVACION` (salida del motor).
- Detecta automáticamente PC_IE / PC_INFO / NOMENCL según tribunal.
- Genera Word consolidado (Arial 12, justificado, 1.5, datos en negrita). Filas sin datos o plantilla se omiten y se reportan; nunca se cuentan como generadas.

## Textos de observación — fuente única de verdad

Todos los textos que se insertan en RUS viven en `motor/textos_observaciones.json`, nunca hardcodeados en `.py`. Cada regla tiene un flag `confirmado: true/false`. Usar `motor.textos.listar_no_confirmados()` para auditar, o el flujo de la GUI (sección "Confirmación de textos", arriba) para cerrar pendientes sin pasar por una sesión de chat.

Tras confirmar/corregir un texto, correr:
```bash
python tests/generar_goldens.py
```
para que `tests/test_paridad.py` adopte el nuevo texto como referencia (si no se hace esto, el test de paridad fallará a propósito — es el detector de cambios no autorizados al JSON).

## Tests (`tests/`)

```bash
pytest tests/ -v
```

- `test_paridad.py`: compara `render()` contra goldens congelados (`goldens_textos.json`). Detecta sabotaje/edición no autorizada del JSON de textos.
- `test_regresion.py`: corre los 3 modos en un **subprocess aislado** (`runner.py`) contra Excel reales si se definen las variables de entorno `CSMP_EXCEL_ESPERA` / `CSMP_EXCEL_CUMPLIMIENTO` / `CSMP_EXCEL_INFORMES`; si no están definidas, usa fixtures sintéticos (`fixtures/generar_fixtures.py`) solo como smoke test estructural (no reemplaza la regresión real).
- `test_confirmacion_textos.py`: ciclo completo exportar → editar → importar del flujo de confirmación (S4).

**Regla para toda IA futura o toda entrega:** ningún ZIP se entrega sin `pytest tests/` en verde.

## Configuración

La configuración se autogenera en `%LOCALAPPDATA%\CSMP_RUS\config_rus.json`
(o `~/CSMP_RUS` fuera de Windows) y no se incluye en la distribución:
```json
{
  "ruta_entrada_excel": "...",
  "ruta_salida_excel": "...",
  "ruta_correo_informes": "...",
  "ruta_correo_espera": "...",
  "ruta_logs": "...",
  "dias_retencion_logs": 90
}
```

El CC institucional está fijado en código y no puede sobrescribirse desde una configuración local.
La bitácora diaria se guarda bajo `ruta_logs`, rota según `dias_retencion_logs` y registra
solo eventos y conteos operativos: no almacena RIT, RUT ni nombres.

`comunicaciones/catastro_programas.json`: catastro operativo. Su publicación requiere autorización institucional documentada; los campos no utilizados deben minimizarse.

`comunicaciones/aliases_programas.json`: mapeo nombre-RUS → nombre-catastro cuando difieren.

## Agregar nuevo alias de programa

Si RUS exporta un programa con nombre que no resuelve contacto, editar `comunicaciones/aliases_programas.json`:

```json
{
  "alias": "NOMBRE EXACTO EN RUS",
  "nombre_catastro": "NOMBRE EXACTO EN CATASTRO"
}
```

## Restricciones inviolables

- Sin acceso directo a RUS/SITFA/SATURNO — todo input es manual.
- Sin envío automático de correos — solo borradores en Outlook Drafts.
- CC fijo: `ucc_concepcion@pjud.cl` en todos los borradores.
- Output 100% supervisado por humano antes de enviar.
- Textos de observación: copy-paste literal de lo aprobado por la persona usuaria — nunca parafraseados.

## Compilación a .exe

```bash
python -m pip install pyinstaller==6.14.2
pyinstaller --clean --noconfirm Asistente_RUS.spec
```

El `.spec` incluye el catálogo de textos, catastro, aliases y plantillas HTML. En el primer inicio,
el catálogo editable se copia a `%LOCALAPPDATA%\CSMP_RUS` para que cambios y backups persistan
fuera del directorio temporal de PyInstaller.

El `.exe` queda en `dist/`. Antes de una entrega, la CI debe estar verde en Windows y Linux y se
debe verificar manualmente un Excel autorizado de cada modo. Un texto o placeholder ausente ahora
cancela el procesamiento; no se produce un archivo aparentemente exitoso.

## Estado actual

Ver `CHANGELOG.md` para historial completo.
Ver `SNAPSHOT.json` para estado consolidado de proyecto, reglas y pendientes.



## v9.0.2 — Auditoría de la revision3 aplicada

- Guardado atómico reparado para Windows (sin `os.fsync` sobre handle de lectura).
- Anti-fórmulas por tipo de celda (texto exacto, sin apóstrofes que ensucien `---`).
- Hoja2: duplicados conservan la última fila (C-10) y una Hoja2 defectuosa nunca aborta CUMPLIMIENTO.
- Correos: fuzzy de contactos restaurado (con guardas de ambigüedad) y lotes con degradación suave — una fila mala no cancela el lote.
- Orden §9.2 (audiencia al final) y flujo aprobado de resoluciones (COMPLETAR / SIN PLANTILLA) restaurados.
- 65 tests nuevos de bordes (Apéndice C del plan). Suite: 168 passed.

## v9.0.1 — Endurecimiento de integridad y seguridad

- Corrige importación `NaN`, propagación de incidencias y mapeo C-07.
- Endurece destinatarios, HTML, agrupación DCE, Tkinter/COM, validación y resoluciones.
- Incorpora CI Windows/Linux, CodeQL, Dependabot, lock de dependencias y build PyInstaller reproducible.
- Motor actualizado a reglas v9 para ESPERA, CUMPLIMIENTO e INFORMES, con textos confirmados.
- La salida de los tres modos incluye `FECHA_OBS` antes de `OBSERVACION` y `TT`, `CC`, `RES` después; `TT-CC` queda disponible para gestión trimestral.
- Correos elimina el borrador a tribunal, usa dos flujos separados (informes y espera), normaliza tribunal y distingue DCE como informes diagnósticos.
- El validador pre-procesamiento queda activo antes de generar Excel de salida.

Consulta `SECURITY.md` antes de reportar errores que involucren datos de casos.
