# Fase 2 – Estabilidad de errores

Objetivo (1–2 semanas): endurecer el manejo de errores en `docir-parser`, `docir-app` y `docir-diff` para evitar fallos por pánico en producción y centralizar errores por crate.

## Controles obligatorios (fase actual)

- Prohibir nuevos usos de `unwrap!`/`expect!`/`panic!`/`unreachable!` en código de producción de:
  - `crates/docir-parser/src`
  - `crates/docir-app/src`
  - `crates/docir-diff/src`
- Mantener un tipo de error único por crate:
  - `docir-parser` → `ParseError`
  - `docir-app` → `AppError`
  - `docir-diff` → `DiffError`
- Conversiones explícitas entre capas (`From` o `map_err`) cuando un caso de uso cambia de capa.
- Errores de producción de `docir-app` y `docir-diff` deben propagarse como `Result` (sin `panic`).

## Punto de control de revisión

Se incluye un gate adicional en `./scripts/quality_gate.sh`:

- `stage_no_unwrap_expect_in_production`
  - Ejecuta `./scripts/quality_no_unwrap_expect_in_production.sh`
  - Revisa diferencias contra `origin/main...HEAD` (configurable con `QUALITY_NO_UNWRAP_BASE`)
  - Falla la ejecución si detecta nuevas apariciones en producción.

### Comando del control

```bash
./scripts/quality_no_unwrap_expect_in_production.sh
```

Modo de trabajo rápido (solo cambios locales sin base):

```bash
./scripts/quality_no_unwrap_expect_in_production.sh working
```

## Observabilidad de cumplimiento

- Cada ejecución del gate imprime un listado de líneas agregadas por archivo cuando se detecta una violación.
- En CI/local acceptance, cualquier violación convierte `./scripts/quality_gate.sh` en fallo (`QUALITY_GATE_RESULT=FAIL`).

## Excepciones

- Los usos existentes de `unwrap/expect/panic/unreachable` dentro de `tests` se mantienen permitidos.
- Otros cambios en reglas de parsing/compatibilidad requieren justificación de deuda en el plan de fase.
