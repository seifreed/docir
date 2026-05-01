# Dashboard de deuda técnica de Fase 1

Fuente de verdad: snapshots generados por `scripts/quality_phase1_snapshot.sh`.

| Semana | Commit | fmt_check | api_hygiene | layer_policy | dependency_cycles | parser_pipeline_contracts | presentation_boundary | Files > 800 LOC | Funciones > 100 LOC | unwrap/expect/panic | Duplicados (grupos) | Baseline report |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| 2026-10 | cea3128 | PASS | PASS | PASS | PASS | PASS | PASS | 0 | 0 | 0 | 0 | /Users/seifreed/tools/malware/docir/target/quality-baseline/quality-baseline-20260304T101336Z.md |

## Política operativa vigente

1. `quality_gate` es la única puerta de aceptación.
2. Wildcard imports (`use super::*`) en producción: modo actual `no nuevos en diff`; transición planificada a inventario estricto (`QUALITY_NO_WILDCARD_INVENTORY_FAIL=1`).
3. Robustez productiva (`panic/unwrap/expect/unreachable`): inventario semanal obligatorio y saldo decreciente.
4. Tamaño de funciones:
   - hard control actual: `>100 LOC` (higiene API);
   - objetivo técnico de madurez: reducir progresivamente a `>80 LOC` o justificar excepción.
5. Tamaño de archivos:
   - hard fail `>800 LOC` en producción;
   - soft warning recomendado a partir de `650 LOC`.

## Scripts de seguimiento semanal

- `bash scripts/quality_phase1_snapshot.sh`
- `bash scripts/quality_no_unwrap_expect_in_production.sh inventory`
- `bash scripts/quality_no_wildcard_super_in_production.sh inventory`
- `bash scripts/quality_duplicate_patterns.sh 3 12`
- `bash scripts/quality_dashboard_freshness.sh`
