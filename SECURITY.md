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
