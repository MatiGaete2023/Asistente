# Plan ejecutable de mejoras v9.1 — endurecimiento sin re-arquitectura

## Principios rectores

1. **Cambios pequeños y verificables.** El proyecto ya tiene historial de regresiones introducidas por reescrituras grandes; por tanto, cada fase debe ser acotada, reversible y cubierta por pruebas.
2. **Especificación funcional cerrada.** No se cambian textos, reglas de negocio, normalización histórica ni comportamiento funcional salvo que el usuario lo confirme explícitamente.
3. **Un operador local.** Las mejoras priorizan robustez operativa, diagnósticos claros y protección local razonable, no patrones corporativos sobredimensionados.
4. **CI y pruebas antes que capas.** La disciplina de seguridad del proyecto se basa en catálogo cerrado, pruebas de regresión y checks automatizados; no en una re-arquitectura global.
5. **Sin PII innecesaria.** La bitácora debe seguir sin datos personales/casos. Los reportes de validación pueden mostrar identificadores necesarios para corregir filas, pero deben tener retención.

## Exclusiones explícitas

- No dividir `gui/app.py` en controladores por pestaña durante v9.1.
- No introducir arquitectura nueva de cuatro capas ni mover masivamente reglas entre paquetes.
- No centralizar sanitización en un módulo nuevo si las defensas actuales por destino siguen testeadas.
- No enmascarar RIT/RUT en reportes de validación si impide identificar filas a corregir.
- No cambiar `normalizar_match` ni agregar validadores especulativos de cruces Hoja2 sin evidencia productiva.
- No cambiar firmas del motor para inyectar `Clock`/`fecha_referencia`; el determinismo de tests se resolverá con herramienta de congelamiento temporal.
- No agregar advertencias intrusivas por rutas UNC/nube; solo documentación operativa.
- No implementar UI de gestión de catálogos en v9.1.
- No agregar SBOM/firma de binarios como requisito de v9.1; solo checksum SHA-256 del ejecutable si hay CI de build.
- No bloquear merges por cobertura o estilo heredado no crítico en esta etapa.

## Fase M1 — Límite defensivo de dimensiones Excel

### Objetivo
Evitar consumo excesivo de memoria al abrir libros inesperadamente grandes, sin afectar planillas reales del flujo RUS.

### Alcance
- Agregar prechequeo de dimensiones para `.xlsx`/`.xlsm` usando `openpyxl.load_workbook(..., read_only=True, data_only=True)` sin cargar celdas completas.
- Umbrales fijos, no configurables: **50.000 filas** y **100 columnas** por hoja relevante.
- Mantener los límites existentes de extensión, tamaño físico y tamaño ZIP descomprimido.
- Para `.xls`, conservar validación actual y documentar que no se aplica prechequeo read-only de dimensiones.

### Criterios de aceptación
- Un `.xlsx` con dimensiones superiores al umbral se rechaza antes de `pd.read_excel` con mensaje claro para usuario.
- Un fixture normal de `ESPERA`, `CUMPLIMIENTO` e `INFORMES` sigue procesando sin cambios de salida.
- Hay pruebas unitarias para aceptación bajo umbral y rechazo sobre umbral.
- `pytest -q` pasa completo.

### Exclusiones de fase
- No hacer streaming de procesamiento pandas.
- No externalizar umbrales a configuración.
- No reescribir lectores Excel del motor.

## Fase M2 — Diagnóstico Outlook y dry-run HTML

### Objetivo
Reducir errores COM crípticos y permitir revisar borradores aunque Outlook no esté disponible.

### Alcance
- Agregar función de preflight Outlook con chequeos mínimos: Windows, `pywin32`, creación de `Outlook.Application`, cuenta/perfil accesible y creación de item de correo en memoria.
- Agregar botón GUI **Verificar entorno** en la pestaña Correos con resultado accionable.
- Reutilizar el `despachador` inyectable de `GeneradorCorreos` para exportar borradores a `.html` cuando el usuario elija modo dry-run o Outlook falle antes de crear borradores.
- Exportar un archivo HTML por borrador en carpeta seleccionada o subcarpeta de salida, con asunto saneado como nombre y cuerpo escapado/plantilla ya generada.

### Criterios de aceptación
- En entorno no Windows, el preflight informa que Outlook COM no está disponible y no produce traceback críptico.
- En Windows sin `pywin32`, el mensaje indica dependencia faltante y comando de instalación.
- El dry-run genera archivos `.html` revisables sin invocar Outlook.
- Pruebas unitarias cubren exportación HTML con despachador falso y validación de nombres de archivo.
- `pytest -q` pasa completo.

### Exclusiones de fase
- No generar `.eml`.
- No enviar correos automáticamente.
- No modificar plantillas de texto institucional.

## Fase M3 — Retención de reportes y matriz de seguridad

### Objetivo
Evitar acumulación indefinida de reportes de validación y documentar claramente las defensas existentes por destino.

### Alcance
- Aplicar rotación de 90 días a `validacion_*.html` en la carpeta de salida de reportes.
- Reutilizar patrón simple de rotación similar a logs, sin crear observabilidad nueva.
- Actualizar `SECURITY.md` con matriz: destino, riesgo, protección implementada, prueba asociada y dato sensible esperado.
- Documentar que la bitácora operativa no debe registrar RIT/RUT/nombres.

### Criterios de aceptación
- Reportes `validacion_*.html` con fecha en nombre anterior a 90 días se eliminan durante generación de nuevo reporte o tarea de rotación invocada por el validador.
- Archivos que no coinciden con el patrón no se eliminan.
- `SECURITY.md` incluye matriz de Excel, HTML de validación, correos HTML, Word, configuración y logs.
- Pruebas unitarias cubren rotación y preservación de archivos ajenos.
- `pytest -q` pasa completo.

### Exclusiones de fase
- No enmascarar identificadores en reportes de validación.
- No cifrar carpetas locales desde la aplicación.
- No introducir JSONL de observabilidad.

## Fase M4 — CI mínimo de calidad y seguridad

### Objetivo
Bloquear bugs reales y dependencias vulnerables sin exigir limpieza masiva de estilo heredado.

### Alcance
- Agregar workflow CI con:
  - `pytest -q`.
  - `ruff check --select F,B,E9 .` como gate de errores/bugs reales.
  - `bandit -r . -ll`.
  - `pip-audit` u OSV sobre dependencias instalables/lockfile.
  - secret scanning simple con gitleaks si está disponible en acción mantenida.
- Reporte de cobertura informativo con `pytest-cov`, sin umbral duro inicial.
- Si existe build de ejecutable, publicar checksum SHA-256 como artefacto.

### Criterios de aceptación
- El workflow ejecuta en pull requests y push a la rama principal.
- El gate Ruff no falla por estilo heredado fuera de `F,B,E9`.
- Bandit corre en severidad media/alta y no bloquea por issues de baja severidad.
- La cobertura se publica como salida informativa.
- Documentación de comandos locales equivalentes en README o docs.

### Exclusiones de fase
- No agregar pre-commit obligatorio.
- No exigir umbral mínimo de cobertura.
- No activar todo Ruff como bloqueo.
- No exigir SBOM/firma de binarios.

## Fase M5 — Determinismo temporal de pruebas

### Objetivo
Hacer reproducibles las pruebas sensibles a fechas sin tocar firmas de producción.

### Alcance
- Agregar `time-machine` a dependencias dev.
- Congelar fecha en pruebas donde cambian observaciones, vencimientos, mayoría de edad, reportes o nombres de archivo cuando sea necesario.
- Mantener `FECHA_OBS` como constancia operacional del día de procesamiento.

### Criterios de aceptación
- Pruebas sensibles a fecha pasan igual al ejecutar en distintas fechas del calendario.
- No se cambian firmas públicas de `procesar`, `calcular_preview`, reglas ni generadores.
- `pytest -q` pasa completo con fecha congelada en tests relevantes.

### Exclusiones de fase
- No implementar clase `Clock`.
- No parametrizar fecha de referencia en GUI.
- No cambiar reglas de negocio por fecha.

## Fase M6 — CLI operativo mínimo

### Objetivo
Permitir ejecutar regresiones y flujos básicos en la máquina del usuario sin depender de Tkinter ni de variables internas de pytest.

### Alcance
- Agregar entrypoint CLI, por ejemplo `csmp-rus`, con subcomandos:
  - `procesar --modo ESPERA|CUMPLIMIENTO|INFORMES --entrada ARCHIVO --salida DIR`.
  - `preview --modo ... --entrada ARCHIVO` con resumen tabular o exportación opcional.
  - `correos --tipo informes|espera --entrada ARCHIVO --dry-run-html DIR`.
  - `resoluciones --entrada ARCHIVO --salida DIR`.
- Devolver códigos de salida claros: `0` éxito, `1` error de validación/entrada, `2` error interno.
- Reusar casos de uso existentes; no duplicar reglas.

### Criterios de aceptación
- Cada subcomando tiene `--help` y ejemplos documentados.
- Un flujo básico por modo se prueba con fixtures existentes.
- Los errores se muestran sin traceback por defecto y con modo verbose si se implementa.
- `pytest -q` pasa completo.

### Exclusiones de fase
- No reemplazar la GUI.
- No introducir servidor web/API.
- No automatizar envío de correos.

## Fase M7 — Mantenimiento acotado y trazabilidad técnica

### Objetivo
Mejorar diagnósticos y reducir deuda pequeña sin reescrituras.

### Alcance
- Agregar `logger.exception` técnico en catch amplios de GUI o adaptadores donde ya exista logger disponible, manteniendo mensajes amigables al usuario.
- Quitar versiones hardcodeadas en docstrings cuando sean claramente divergentes; la fuente de verdad es `motor/version.py` y metadata del paquete.
- Evaluar `TypedDict` o dataclasses solo en módulos nuevos o resultados con bajo acoplamiento, sin tocar contratos dict ampliamente testeados salvo beneficio claro.
- Aplicar autofix de imports (`ruff --select I --fix`) solo si el diff es pequeño y no altera lógica.

### Criterios de aceptación
- Errores inesperados quedan con traceback en log técnico sin exponer PII en mensajes operativos.
- No quedan docstrings principales anunciando versión contradictoria con `motor/version.py`.
- Cualquier tipado nuevo no requiere cambios masivos de llamadas existentes.
- `pytest -q` pasa completo.

### Exclusiones de fase
- No refactorizar `RUSApp` completo.
- No convertir todos los dicts del proyecto a dataclasses en una sola tanda.
- No activar Ruff completo como obligación.

## Orden recomendado de ejecución

1. **M4 CI mínimo**: primero asegurar red de seguridad automática.
2. **M1 Excel dimensions**: endurecimiento de mayor valor defensivo con bajo alcance.
3. **M3 retención/SECURITY**: cierra documentación y acumulación de artefactos locales.
4. **M2 Outlook/dry-run**: mejora operativa visible y aislada.
5. **M5 determinismo temporal**: estabiliza regresiones a largo plazo.
6. **M6 CLI**: habilita soporte y ejecución reproducible fuera de GUI.
7. **M7 mantenimiento acotado**: ejecutar en pequeños PRs cuando M1-M6 estén verdes.

## Definición de terminado para v9.1

- Todas las fases aceptadas M1-M6 completadas o explícitamente reprogramadas con justificación.
- `pytest -q` verde.
- CI ejecutándose en PR/push con gates mínimos definidos.
- `SECURITY.md` actualizado con matriz de salidas y retención.
- No se introdujeron cambios funcionales fuera del alcance de las fases.
- No se reescribieron componentes completos sin necesidad demostrada.
