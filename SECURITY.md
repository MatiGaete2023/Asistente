# Política de seguridad

## Reporte responsable

No publiques datos reales de causas, RIT, RUT, nombres, planillas, correos ni capturas en un issue.
Usa **Security → Report a vulnerability** para reportar una vulnerabilidad de forma privada al
mantenedor. Incluye versión, pasos mínimos con datos sintéticos e impacto esperado.

## Datos permitidos en el repositorio

- Los fixtures y pruebas deben usar únicamente datos inventados.
- Los logs no deben contener RUT ni nombres completos.
- El catastro de contactos solo puede publicarse con autorización institucional documentada.
- Credenciales, tokens, archivos de configuración local y salidas del motor nunca se versionan.

## Liberaciones

Una versión es liberable únicamente con CI verde en Windows y Linux, cero hallazgos críticos o
altos abiertos, dependencias auditadas y revisión humana del ejecutable firmado.


## Matriz de salidas locales y controles

| Destino | Dato esperado | Riesgo principal | Control implementado | Prueba/validación |
|---|---|---|---|---|
| Excel de salida del motor | RIT, RUT, nombres, observaciones | Fórmulas inyectadas desde planilla origen | Escritura con `openpyxl` y neutralización de celdas que parezcan fórmula | Regresiones de seguridad de salida Excel |
| Reporte HTML de validación | Detalle necesario para corregir filas | Acumulación local e inyección HTML | `html.escape(..., quote=True)` y rotación de `validacion_*.html` a 90 días | Tests de reporte/rotación |
| Correos HTML / borradores Outlook | Identificadores de casos que requieren gestión | HTML injection y destinatarios inyectados | Escape de celdas en tablas y validación estricta de email único | Tests de correos y seguridad |
| Resoluciones Word | Datos de causa para edición humana | Contenido manipulado en documento final | Documento local supervisado; no ejecuta fórmulas ni macros | Revisión humana obligatoria |
| Configuración local | Rutas y retención | Escritura parcial/corrupción | Guardado atómico con temporal y `os.replace` | Pruebas de configuración |
| Bitácora operativa | Eventos, conteos y tipos de error | Filtración de PII | Política explícita: no registrar RIT/RUT/nombres; rotación diaria | Revisión de eventos en GUI/log manager |

Los reportes de validación pueden contener identificadores necesarios para corregir planillas. Deben permanecer en carpetas locales autorizadas y se eliminan automáticamente cuando superan la retención operativa de 90 días.
