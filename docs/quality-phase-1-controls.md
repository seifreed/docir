# Fase 1 – Medición y controles base

Estado objetivo (1 semana): controles iniciales objetivos, repetibles y trazables para sostener refactors posteriores sin regresión.

## Umbrales obligatorios

- Límite de archivo: `> 800 LOC` → violación (excepto justificación documentada de deuda técnica justificada).
- Límite de función: `> 100 LOC` en código de producción → violación (excepto función algorítmica justificada con motivo en documento de excepción).
- Prohibición en producción: `unwrap!`, `expect!`, `panic!`, `unreachable!`.
- Errores tipados obligatorios en rutas de producción.

## Política de visibilidad interna (obligatoria)

- `pub(crate)` por defecto en internals del parser y servicios.
- `pub(super)` solo para cooperación dentro del subárbol inmediato.
- `pub` únicamente para contratos de frontera entre módulos/capas.
- `use super::*`:
  - permitido temporalmente solo como deuda histórica existente;
  - prohibido introducir nuevos usos en diff (`quality_no_wildcard_super_in_production.sh`).

## Checklist de arquitectura

### Capas y dependencias permitidas

- `docir-core` (dominio): no depende de crates de aplicación, parser, serialización, reglas, diff, CLI ni bindings.
  - Permitidas actuales: `serde` (opcional), `thiserror`.
- `docir-parser` (infraestructura de parsing): puede depender de `docir-core` y librerías de parsing/IO/crypto.
  - Permitidas actuales: `zip`, `quick-xml`, `encoding_rs`, `calamine`, `flate2`, `sha2`, `sha1`, `pbkdf2`, `base64`, `aes`, `cbc`, `log`, `thiserror`, `docir-core`.
- `docir-app` (casos de uso/orquestación): puede depender de `docir-core` + servicios especializados.
  - Permitidas actuales: `docir-core`, `docir-parser`, `docir-security`, `docir-serialization`, `docir-diff`, `docir-rules`, `thiserror`.
- `docir-security` / `docir-rules` / `docir-diff` / `docir-serialization`: capa intermedia por dominio; dependencia hacia núcleo y librerías de apoyo únicamente.
- `docir-cli`, `docir-python`: capas de entrada/salida, sin lógica de negocio propia crítica.
  - Dependencias esperables: dominio/app y utilidades de presentación/bindeo.

### Puntos de fuga de infraestructura a vigilar (base de checklist)

- Infraestructura entrando al dominio (`docir-core`)
  - `quick-xml`, `zip`, `flate2`, `calamine`, `clap`, `pyo3`, `env_logger`, `tokio`/`async-std` u otras APIs de I/O de alto nivel.
- Aplicación con salida de formato en núcleo
  - Formato JSON/CSV/CLI embebido dentro de `docir-core` o funciones de dominio.
- Dependencias inversas
  - `docir-core` usado por `docir-parser` (correcto) y no al revés.
  - `docir-app` como orquestador; sin parseo directo de contenedor/CLI dentro de `docir-core`.

### Riesgos de arquitectura que deben revisarse por defecto

- Ficheros "god-file" con responsabilidades múltiples en `docir-parser`.
- Cierre de contratos por reglas de negocio dentro de `docir-parser` o módulos de salida.
- Funciones grandes sin separación de parsing / normalización / postproceso.
- Repetición de patrones de control de flujo entre formatos.

## Baseline de métricas (punto de control)

Se añadió un script de baseline para medir métricas por crate y por archivo:

```bash
./scripts/quality_phase1_baseline.sh
```

Comportamiento esperado:

- Genera un reporte en `target/quality-baseline/quality-baseline-<timestamp>.md`.
- Resume por crate:
  - número de ficheros
  - LOC total
  - ficheros > 800 LOC
  - apariciones de `unwrap/expect/panic/unreachable` en producción
  - funciones > 100 LOC (estimación heurística)
- Lista candidatos de fuga:
  - dependencias fuera de lo permitido por capa
  - usos de librerías de infraestructura detectados en `docir-core`.
- Scope explícito de robustez (coherente con inventario):
  - `docir-parser`, `docir-app`, `docir-diff`, `docir-security`.

Opcional:

```bash
./scripts/quality_phase1_baseline.sh --fail-on-violations
```

Este modo devuelve error si hay cualquier violación detectada por umbrales/regex de fase 1.

## Ejecución semanal operativa (Fase 1)

Comando operativo recomendado:

```bash
bash scripts/quality_phase1_snapshot.sh
```

Salida esperada:

- `target/quality-phase1/<timestamp>/quality_phase1_snapshot.md`
- `target/quality-phase1/<timestamp>/quality_duplicate_patterns.md`
- `docs/quality-debt-dashboard.md` (append de la fila semanal)

Verificaciones requeridas por semana:

1. `cargo fmt --all --check`
2. `./scripts/quality_api_hygiene.sh`
3. `./scripts/quality_layer_policy.sh`
4. `./scripts/quality_dependency_cycles.sh`
5. `./scripts/quality_parser_pipeline_contracts.sh`
6. `./scripts/quality_presentation_boundary.sh`
7. `./scripts/quality_phase1_baseline.sh`
8. `./scripts/quality_duplicate_patterns.sh`

La fila semanal debe registrar al menos:

- unwrap/expect/panic en producción
- archivos > 800 LOC
- funciones > 100 LOC
- grupos de duplicados funcionales (mínimo 3 apariciones)

## Referencia de cumplimiento

- Este baseline es **medición y control**, no reemplaza `./scripts/quality_gate.sh`.
- Debe ejecutarse al inicio y al cierre de cada refactor de la fase 2+.
