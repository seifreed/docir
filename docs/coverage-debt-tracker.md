# Tracker de deuda técnica de cobertura

## Estado base (2026-03-02)

- `coverage_check` objetivo operativo: `quality_gate.sh` con umbral 88.23.
- Umbral base: `scripts/quality_coverage_threshold.txt`.
- Cobertura total actual (resumen de `cargo llvm-cov --workspace --all-features --summary-only`): `88.23%`.
- Método: no bajar cobertura con excepciones de umbral; mejorar mediante tests dirigidos y ajustes de comportamiento con casos reales.

## Saldo de deuda residual (cerrado por semana)

- Saldo objetivo inicial: `84` hallazgos productivos de `unwrap/expect/panic`.
- Regla operativa: cada semana debe bajar el saldo en `8` puntos como mínimo.
- Saldo actual registrado:
  - Semana 1 (base): `84`
  - Semana 2 objetivo: `76`
  - Semana 3 objetivo: `68`
  - Semana 4 objetivo: `60`
  - Semana 5 objetivo: `52`

Checklist corto por semana:
- Mantener registro de qué patrón quedó con justificación técnica y su fecha de remediación.
- Si un patrón no se puede cerrar, documentar explícitamente y moverlo a `P1` con owner/ETA.

## Lista de deuda residual (prioridad de impacto)

| Prioridad | Archivo | Cobertura actual | Acción propuesta |
|---|---|---:|---|
| P0 | `docir-parser/src/odf/limits.rs` | 50.44% | Añadir escenarios de límites y errores de validación para ramas no ejecutadas. |
| P0 | `docir-parser/src/odf/spreadsheet/spreadsheet_parse.rs` | 54.77% | Cobertura de rutas de parsing de pestañas y celdas con fixtures no triviales. |
| P0 | `docir-parser/src/rtf/core/field_utils.rs` | 63.77% | Completar pruebas de estados de campo y utilidades de campo. |
| P0 | `docir-parser/src/ooxml/xlsx/worksheet/worksheet_parse.rs` | 60.42% | Agregar pruebas de workbook/worksheet con variantes de estilos y fórmulas. |
| P1 | `docir-parser/src/rtf/core/core_parse.rs` | 70.55% | Aumentar fixtures con mutaciones de estado y cierre de errores sintácticos. |
| P1 | `docir-parser/src/odf/helpers/helpers_parse.rs` | 68.29% | Cubrir ramas de normalización y fallback de helpers con entradas mixtas. |
| P1 | `docir-parser/src/odf/ods/ods_parse.rs` | 78.70% | Reforzar pruebas de parser de hojas para condiciones alternativas. |
| P1 | `docir-parser/src/odf/ods/ods_normalize.rs` | 74.47% | Tests de normalización y validación de estado. |
| P1 | `docir-parser/src/odf/mod.rs` | 70.00% | Cubrir flujo de integración y errores de módulos externos. |
| P1 | `docir-parser/src/ooxml/docx/document/numbering.rs` | 79.27% | Cubrir variantes de numeración anidadas (sin fixtures sintéticos vacíos). |
| P2 | `docir-core/src/query.rs` | 72.07% | Añadir pruebas de query-builder en casos límite de nodos nulos. |
| P2 | `docir-core/src/ir/mod.rs` | 75.56% | Expandir pruebas por constructor de nodos y rutas de mutación. |
| P2 | `docir-core/src/visitor/store.rs` | 76.09% | Mejorar cobertura de almacenamiento de visitantes y limpieza de estado. |
| P2 | `docir-core/src/types.rs` | 79.31% | Completar casos de conversión de tipos y errores de coerción. |
| P2 | `docir-parser/src/odf/sampling.rs` | 76.15% | Cobertura de casos de muestreo parcial y límites de ventana. |

## Cadencia de cierre (ciclos semanales)

### Semana 1
- Objetivo: cubrir los 5 P0.
- Criterio: bajar cobertura mínima de estos 5 archivos por encima de 70% y sin aumentar `CC-12`/`CC-13`.

### Semana 2
- Objetivo: cubrir al menos 50% de los P1 según capacidad del equipo.
- Criterio: +1.5pp de cobertura total por semana sin cambios de arquitectura.

### Semana 3
- Objetivo: cierre parcial de P2 y revisar regresión de umbral.
- Criterio: umbral base +2.0pp acumulado en módulos críticos objetivo (parser/core).

## Revisión semanal (obligatoria)

1. Ejecutar `./scripts/quality_gate.sh`.
2. Ejecutar `cargo llvm-cov --workspace --all-features --summary-only` y actualizar este tracker.
3. Registrar cada acción en formato:
   - `archivo | acción | tests ejecutados | cobertura antes | cobertura después | responsable | estado`.
4. Validar que no aparecen importaciones CC-12/CC-13 nuevas y que no se reabren dependencias prohibidas.

## Nota operativa

- Si se decide subir el umbral base, actualizar:
  - `scripts/quality_coverage_threshold.txt`
  - este tracker (baseline y objetivo de semana)
  - `docs/coverage-debt-tracker.md` (alineado con evidencia de CI y scripts).
