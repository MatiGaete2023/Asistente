# CSMP Assistant — Documentación de la versión v9 (planificación)

**Estado del repositorio (15 de julio de 2026):** la raíz contiene el código **v8.14 tal cual** vino en `CSMP_Assistant_v8_14_COMPLETO.zip` (baseline sin modificar, suite verificada: 73 passed, 3 skipped). Todo lo nuevo de esta etapa vive en `docs/`.

## Contenido

| Documento | Qué es |
|---|---|
| `especificaciones/Catalogo_Reglas_CSMP_v2_20260714.md` | **Fuente única de verdad** de las reglas de observación (E/T/C/I/G), columnas de salida (§8) y criterios de prueba (§13). |
| `especificaciones/Mejoras_Correos_y_Otros_Modulos_20260714.md` | Decisiones cerradas sobre correos, validador, resoluciones, GUI y procesador (14 aceptadas, 8 rechazadas, 1 corrección). |
| `auditoria/AUDITORIA_CODIGO_v8_14.md` | Revisión senior del código v8.14: 17 hallazgos propios (H-01…H-17) con archivo:línea, verificación de los hallazgos de los MD, y notas de riesgo. |
| `plan/PLAN_IMPLEMENTACION_v9.md` | **El plan ejecutable**: 12 fases ordenadas (F1–F12) con tareas, archivos, criterios de aceptación, decisiones por default (D1–D13), textos definitivos del JSON (Apéndice A), pseudocódigo de los tres motores (Apéndice B) y matriz de pruebas de borde (Apéndice C). |

## Cómo ejecutar el plan (para Codex / Sonnet / cualquier agente)

1. Leer `plan/PLAN_IMPLEMENTACION_v9.md` §0 (instrucciones y Definición de HECHO) y §2 (decisiones).
2. Ejecutar las fases **en orden F1 → F12**, un commit por fase, suite verde al cierre de cada una (excepción documentada entre F2 y F5).
3. Ante cualquier ambigüedad: catálogo v2 > mejoras > plan > PROMPT_CONTINUACION.md (este último es **histórico**, contiene decisiones ya superadas).
4. Los textos de observación se copian **verbatim** del Apéndice A del plan. Jamás parafrasear.

## Jerarquía de precedencia de documentos

```
Catalogo_Reglas_CSMP_v2  >  Mejoras_Correos_y_Otros_Modulos  >  PLAN_IMPLEMENTACION_v9  >  AUDITORIA  >  README/CHANGELOG/PROMPT_CONTINUACION (histórico v8.14)
```
