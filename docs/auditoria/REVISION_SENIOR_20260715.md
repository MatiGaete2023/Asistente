# Revisión senior de arquitectura, seguridad y calidad — 2026-07-15

## Alcance

Repositorio Python/Tkinter para procesamiento local de planillas RUS, generación de borradores Outlook y documentos Word. La revisión cubrió estructura, empaquetado, flujo GUI, motor de reglas, validación, correos, resoluciones, pruebas y controles de seguridad existentes.

## Resumen ejecutivo

El proyecto muestra una orientación sólida a operación local supervisada: valida archivos Excel antes de abrirlos, neutraliza fórmulas al exportar, escapa HTML en correos/reportes, usa guardado atómico en varias salidas y contiene regresiones de seguridad. Los riesgos principales no son de exposición web sino de robustez ante insumos grandes/malformados, tratamiento de datos sensibles en artefactos locales, bloqueo de la interfaz durante cargas pesadas y dependencias/automatización COM en Windows.

## 1. Vulnerabilidades y bugs críticos

### Hallazgos de alta prioridad

1. **Lectura completa de Excel en memoria sin límite de filas/columnas.**  
   `validar_archivo_excel` limita tamaño de archivo y tamaño ZIP descomprimido, pero luego `pd.read_excel` carga hojas completas. Un Excel de 100 MB o 200 MB descomprimido puede causar consumo de memoria alto, congelar Tkinter o agotar el proceso, especialmente en equipos de oficina. Recomendación: agregar límites configurables de filas/columnas por modo, prelectura de metadatos con `openpyxl` en modo read-only para `.xlsx`, mensajes de rechazo tempranos y pruebas de estrés.

2. **Reportes HTML y bitácoras pueden contener identificadores sensibles.**  
   El validador genera HTML local y la bitácora registra rutas y resultados. Aunque el HTML escapa contenido, puede almacenar RIT/RUT/nombres en detalle si las reglas incluyen dichos datos. Recomendación: política explícita de minimización, enmascarar RUT/RIT en reportes por defecto, cifrado opcional de carpeta local o instrucciones operativas, y retención diferenciada para reportes de validación.

3. **Automatización Outlook por COM depende de estado local no controlado.**  
   La generación de borradores usa `Display(False)` para obtener firma y luego modifica `HTMLBody`. Si Outlook muestra diálogos, una cuenta no está configurada, una política corporativa bloquea COM o existe un add-in defectuoso, el flujo puede fallar. El código captura excepciones y no envía correos, lo cual es positivo, pero falta una comprobación diagnóstica previa y guía de recuperación. Recomendación: preflight de Outlook, modo dry-run/exportación `.html` y log técnico separado.

4. **Riesgo residual de inyección por fórmulas fuera de `.xlsx` generado.**  
   El guardado Excel neutraliza fórmulas en libros generados con `openpyxl`, pero los módulos que generan Word/HTML/correos siguen incorporando valores derivados de Excel en otros formatos. En HTML se escapa correctamente; en Word no ejecuta fórmulas, pero sí puede contener texto manipulado. Recomendación: mantener una política central de saneamiento por destino (`excel`, `html`, `docx`, `log`) y pruebas para cada salida.

### Hallazgos de prioridad media

1. **Uso repetido de `datetime.now()` impide reproducibilidad.**  
   Reglas, validaciones, nombres de archivos y cálculos de vencimiento dependen del reloj del sistema. Esto dificulta auditoría histórica y pruebas deterministas. Recomendación: inyectar un `Clock`/`fecha_referencia` en motor, correos y resoluciones.

2. **Manejo de excepciones demasiado amplio en puntos críticos.**  
   Hay `except Exception` que convierten fallas en mensajes amigables, útil para GUI, pero puede ocultar bugs de programación. Recomendación: separar excepciones esperadas de entrada/IO de excepciones internas; incluir `logger.exception` técnico cuando exista bitácora.

3. **Normalización de cruces Hoja2 frágil por diseño histórico.**  
   `normalizar_match` conserva tildes y guiones por compatibilidad, lo que puede reducir matches entre hojas con pequeñas variaciones. Recomendación: conservar modo histórico, pero agregar modo diagnóstico que reporte candidatos no cruzados por diferencias leves.

4. **Configuración local sin validación de rutas de confianza.**  
   La GUI acepta rutas arbitrarias de entrada/salida/logs. En operación local esto es esperable, pero rutas UNC/remotas o carpetas sincronizadas podrían exponer datos. Recomendación: advertir si la ruta apunta a red/nube y documentar ubicaciones aprobadas.

## 2. Errores de sintaxis y buenas prácticas

### Estado general

No se observaron errores de sintaxis evidentes en los módulos revisados. La estructura de paquetes es clara (`motor`, `gui`, `comunicaciones`, `resoluciones`, `validador`, `logs`) y `pyproject.toml` define dependencias y herramientas de desarrollo.

### Buenas prácticas ya presentes

- Uso de `defusedxml` como dependencia defensiva del ecosistema Excel.
- Guardados atómicos con archivo temporal y `os.replace` en configuración, reportes y salidas.
- Escapado HTML con `html.escape` en correos y reportes.
- Neutralización de fórmulas al exportar Excel.
- Pruebas de regresión para inyección HTML, destinatarios de correo, alias de contactos y validaciones.

### Oportunidades de mejora de estilo

1. **Orden de imports y formato consistente.**  
   Algunos archivos mezclan imports estándar, terceros y locales sin separación uniforme. Ruff puede automatizar parte de esto si se agregan reglas `I` de isort.

2. **Funciones largas con múltiples responsabilidades.**  
   `gui/app.py` concentra construcción de UI, coordinación de hilos, carga/guardado de configuración y acciones de tres módulos. Conviene separar controladores por pestaña y servicios de aplicación.

3. **Nombres de versión divergentes.**  
   Algunos docstrings mencionan `v9.0.1` mientras `pyproject.toml` declara `9.0.2`. Recomendación: derivar versión desde `motor.version`/metadata y evitar versiones hardcodeadas en docstrings.

4. **Mensajes y strings de negocio dispersos.**  
   Existen textos en JSON, plantillas HTML y módulos de reglas. Recomendación: catálogo central con IDs, severidad y trazabilidad para facilitar auditoría.

5. **Tipado parcial.**  
   Se usan anotaciones en algunos helpers, pero muchas funciones públicas aceptan `dict`/`DataFrame` sin contratos formales. Recomendación: `TypedDict`/dataclasses para resultados de validación, correos y resoluciones.

## 3. Sugerencias de refactorización

### Arquitectura propuesta

1. **Capa de dominio pura.**  
   Mover reglas RUS y generación de observaciones a funciones puras con dependencias inyectadas (`fecha_referencia`, catálogos, configuración). Esto reduce acoplamiento con pandas/Tkinter.

2. **Capa de IO/adaptadores.**  
   Centralizar lectura Excel, escritura Excel, escritura HTML, Word y Outlook en adaptadores con límites, saneamiento y telemetría homogénea.

3. **Capa de aplicación/orquestación.**  
   Convertir `procesar`, `calcular_preview`, `generar_resoluciones` y `GeneradorCorreos` en casos de uso que devuelvan resultados estructurados, no solo mensajes a cola.

4. **Capa GUI delgada.**  
   Cada pestaña debería invocar casos de uso y renderizar resultados. Reducir lógica de negocio y manejo de archivos dentro de widgets.

### Refactors concretos

- Crear `motor/clock.py` o parámetro `fecha_referencia` para eliminar llamadas directas a `datetime.now()` en reglas.
- Crear `infra/excel_io.py` con `ExcelReadPolicy(max_bytes, max_rows, max_cols, allowed_extensions)` y métricas de carga.
- Crear `infra/sanitizers.py` con funciones `for_excel`, `for_html`, `for_docx`, `for_log`.
- Reemplazar resultados `dict` por dataclasses (`ValidationResult`, `EmailGenerationResult`, `ResolutionGenerationResult`).
- Extraer `RUSApp` en `MotorTab`, `CorreosTab`, `ResolucionesTab` y `ConfigService`.
- Añadir un `AuditLogger` con eventos estructurados JSONL para auditoría técnica, separado de mensajes de usuario.
- Parametrizar catálogos y plantillas con validación JSON Schema.

## 4. Complementos y nuevas funcionalidades

### Seguridad y compliance

- **pip-audit/OSV en CI:** detectar CVEs de dependencias.
- **Bandit en CI:** mantener escaneo de patrones inseguros en Python.
- **Secret scanning:** aunque no se vieron secretos en el muestreo, añadir gitleaks/trufflehog evita filtraciones futuras.
- **SBOM:** generar CycloneDX para entregables PyInstaller.
- **Firma/código de integridad:** firmar ejecutables Windows y publicar checksum SHA-256.

### Calidad y mantenibilidad

- **pre-commit:** ejecutar Ruff, formato, chequeos de fin de archivo y validación JSON.
- **mypy incremental:** empezar por módulos con dataclasses/resultados.
- **pytest-cov:** establecer umbral mínimo por paquetes críticos (`motor`, `validador`, `comunicaciones`).
- **Hypothesis:** pruebas property-based para fechas, normalización, RUT y límites de Excel.
- **freezegun/time-machine:** pruebas deterministas basadas en fecha.

### Escalabilidad operativa

- **Modo CLI:** comandos `procesar`, `preview`, `correos`, `resoluciones`, útiles para automatización supervisada y soporte.
- **Dry-run/exportación de correos:** generar `.eml` o `.html` si Outlook no está disponible.
- **Panel de diagnóstico:** validar dependencias, versión de Python, permisos de carpetas, Outlook y catálogos antes de operar.
- **Observabilidad local:** eventos JSONL con IDs de corrida, conteos, tiempos y errores técnicos sin PII por defecto.
- **Gestión de catálogos:** UI para detectar contactos duplicados, alias conflictivos y correos inválidos.

## Priorización recomendada

1. **P0:** límites de filas/columnas al leer Excel, política de PII en reportes/logs, preflight de Outlook.
2. **P1:** inyección de reloj/fecha de referencia, dataclasses de resultados, sanitización centralizada por destino.
3. **P2:** modularización de GUI, CLI, CI con pip-audit/bandit/ruff/mypy, validación JSON Schema de catálogos.
4. **P3:** SBOM, firma de binarios, property-based tests y panel de diagnóstico avanzado.

## Conclusión

La base es razonablemente segura para una aplicación local supervisada y ya incorpora controles importantes contra fórmulas, HTML injection y errores de datos. Para elevarla a estándar institucional robusto, el foco debe ponerse en límites de recursos, protección de datos sensibles en artefactos locales, contratos tipados de resultados y separación más estricta entre dominio, IO y GUI.

## Addendum de calibración — 2026-07-15

Este informe queda calibrado por el plan ejecutable v9.1. En particular, se corrigen o acotan las siguientes afirmaciones:

- El riesgo de Excel grande se interpreta como consumo de memoria del proceso, no como congelamiento directo de Tkinter, porque las cargas se ejecutan en hilos y la GUI se actualiza vía cola.
- La bitácora operativa debe permanecer sin PII; el punto aceptado es rotación de reportes `validacion_*.html` y documentación de matriz de salidas, no enmascaramiento general de identificadores que el usuario necesita para corregir filas.
- Los textos de negocio ya tienen catálogo central; no se debe proponer una centralización nueva que duplique `motor/textos_observaciones.json`.
- Las defensas por destino ya existen y están testeadas; se documentan en `SECURITY.md` en vez de introducir un refactor centralizado de sanitización.
- Se descartan re-arquitecturas grandes por el riesgo histórico de regresión. La ruta aceptada prioriza CI, cambios pequeños y pruebas.

La planificación aplicable se encuentra en `docs/plan/PLAN_MEJORAS_V9_1.md`.
