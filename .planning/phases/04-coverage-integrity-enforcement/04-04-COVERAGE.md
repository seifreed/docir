# 04-04 Coverage Evidence

## Canonical Commands

```bash
cargo llvm-cov --workspace --all-features --summary-only --fail-under-lines 95 || true
cargo llvm-cov --workspace --all-features --summary-only
```

## Workspace Total (Line Coverage)

- Baseline from 04-03 summary: `65.43%`
- Current 04-04 canonical total: `67.10%`
- Delta: `+1.67` percentage points

## 95% Gate Status

- `cargo llvm-cov --workspace --all-features --summary-only --fail-under-lines 95`
- Exit code: `1` (FAIL, threshold still unmet)
- Observed total in fail-under run: `67.10%`

## Targeted File Snapshots (Current Run)

- `docir-parser/src/parser/security.rs`: `85.43%` lines (`59` lines missed)
- `docir-parser/src/parser/metadata.rs`: `86.87%` lines (`39` lines missed)
- `docir-security/src/enrich.rs`: `89.77%` lines (`27` lines missed)
- `docir-security/src/enrich/dde.rs`: `100.00%` lines (`0` lines missed)
- `docir-security/src/enrich/helpers.rs`: `86.57%` lines (`18` lines missed)
- `docir-security/src/enrich/xlm.rs`: `97.41%` lines (`5` lines missed)

## Residual Highest-Impact Untouched Candidates (Next Gap Closure)

Excluding 04-03-improved modules (`parser/vba.rs`, `parser/analysis.rs`, `security/analyzer.rs`, `serialization/json.rs`) and this plan's target files:

1. `docir-parser/src/odf/spreadsheet.rs` (`350` lines missed, `43.64%`)
2. `docir-parser/src/odf/ods.rs` (`277` lines missed, `55.61%`)
3. `docir-parser/src/odf/helpers.rs` (`244` lines missed, `57.34%`)
4. `docir-parser/src/ooxml/docx/document/inline.rs` (`234` lines missed, `60.27%`)
5. `docir-parser/src/odf/formula.rs` (`207` lines missed, `45.67%`)

