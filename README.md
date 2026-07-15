# CSMP Assistant — Asistente de RUS

**Versión actual:** v8.14.0
**Usuario:** Matías Gaete Agüero (CSMP Concepción)
**Tribunales:** Laja, Mulchén, Tomé

Sistema desktop para CSMP que genera observaciones judiciales automatizadas, borradores de correos Outlook y resoluciones .docx a partir de exportaciones del módulo RUS del Poder Judicial.

## Instalación

```bash
pip install pandas openpyxl xlrd python-dateutil python-docx pywin32
```

`pywin32` solo en Windows con Outlook instalado.

Para correr la suite de tests (opcional, recomendado antes de cualquier entrega):
```bash
pip install pytest
pytest tests/ -v
```

## Ejecución

```bash
python main.py
```

Se abre la GUI con 3 pestañas.

## GUI — 3 pestañas

### 1. Motor RUS
- Procesa archivos Excel de RUS en 3 modos: **ESPERA**, **CUMPLIMIENTO**, **INFORMES**.
- Genera Excel de salida con columna `OBSERVACION` calculada por las reglas del manual CSMP.
- Para CUMPLIMIENTO con Hoja2: cruza ambas hojas por 5 campos (RIT+RUT+NOMBRE+DERIV+TRIBUNAL) y usa la fecha de vencimiento de Hoja2.
- **👁 Vista previa (v8.14):** botón por modo que calcula las observaciones EN MEMORIA (sin escribir ningún archivo) y las muestra en una tabla (NNA | Observación, primeras 30 filas + contador total). El texto es idéntico byte-a-byte al que se exportaría. Botón "Exportar…" dentro de la ventana para recién ahí guardar.
- **📋 Confirmación de textos (v8.14):** sección con el contador de textos pendientes de aprobación (`confirmado:false` en `motor/textos_observaciones.json`). "Exportar pendientes…" genera un Excel editable; tras revisarlo y marcar `APROBADO=SI` / completar `TEXTO_CORRECTO`, "Importar confirmaciones…" aplica los cambios al JSON (con backup automático con timestamp) sin tocar ningún archivo `.py`.

### 2. Correos Outlook
- **Informes vencidos / por vencer**: lee Excel de INFORMES (con `FECHA VENCIMIENTO`), genera borradores agrupados por programa y por tribunal.
- **Lista de espera**: lee Excel de ESPERA procesado por el motor (con columna `OBSERVACION`), 1 borrador por programa+tribunal.
- Encabezado: `SEÑORES: / {programa} / PRESENTE.` (PRESENTE en negrita + subrayado).
- Firma del usuario se inyecta via `mail.Display(False)` (Outlook debe estar instalado).

### 3. Resoluciones .docx
- Lee Excel con columna `OBSERVACION` (salida del motor).
- Detecta automáticamente PC_IE / PC_INFO / NOMENCL según tribunal.
- Genera Word consolidado (Arial 12, justificado, 1.5, datos en negrita).

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

`config_rus.json` (se autogenera con valores por defecto la primera vez que se ejecuta — no se incluye en el ZIP de distribución):
```json
{
  "ruta_entrada_excel": "...",
  "ruta_salida_excel": "...",
  "ruta_logs": "...",
  "cc_fijo": "ucc_concepcion@pjud.cl",
  "dias_retencion_logs": 90
}
```

`contactos.json`: destinatarios por tribunal.

`comunicaciones/catastro_programas.json`: 231 programas con director y mail.

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
- Textos de observación: copy-paste verbatim de lo que escribe Matías — nunca parafraseados.

## Compilación a .exe

```bash
pip install pyinstaller
pyinstaller --noconsole --onefile --add-data "motor/textos_observaciones.json;motor" --name "Asistente_RUS" main.py
```

⚠️ El separador de `--add-data` es `;` en Windows y `:` en Linux/Mac — usar `;` en la máquina de producción (Windows). Sin `--add-data`, el .exe no encuentra `textos_observaciones.json` y todas las observaciones fallan con `[TEXTO NO DEFINIDO]`.

El .exe queda en `dist/`. Antes de entregar el .exe a producción, verificar manualmente que procese un Excel real de cada modo y que las observaciones no digan `[TEXTO NO DEFINIDO: ...]`.

## Estado actual

Ver `CHANGELOG.md` para historial completo.
Ver `SNAPSHOT.json` para estado consolidado de proyecto, reglas y pendientes.

