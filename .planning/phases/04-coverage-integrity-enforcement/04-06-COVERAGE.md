# 04-06 Coverage Evidence

## Canonical Commands

```bash
cargo llvm-cov --workspace --all-features --summary-only --fail-under-lines 95 || true
cargo llvm-cov --workspace --all-features --summary-only
```

## Workspace Total (Line Coverage)

- Baseline from 04-05 canonical total: `68.10%`
- Current 04-06 canonical total (`summary-only` run): `68.62%`
- Delta vs 04-05: `+0.52` percentage points

## 95% Gate Status (Canonical Truth)

- Command: `cargo llvm-cov --workspace --all-features --summary-only --fail-under-lines 95`
- Exit code: `1` (FAIL, threshold still unmet)
- Observed total in fail-under run: `68.61%`
- Required threshold: `95.00%`
- Remaining gap to threshold: `26.39` percentage points

## Targeted Module Snapshots (04-06 Run)

- `docir-parser/src/ooxml/docx/document/inline.rs`: `57.75%` lines (`447` lines missed)
- `docir-parser/src/ooxml/xlsx/worksheet.rs`: `68.36%` lines (`323` lines missed)
- `docir-parser/src/rtf/core.rs`: `72.59%` lines (`281` lines missed)

## Targeted Module Comparison vs 04-05

- `inline.rs`: `60.27%` -> `57.75%` (`-2.52` pp)
- `worksheet.rs`: `65.50%` -> `68.36%` (`+2.86` pp)
- `rtf/core.rs`: `67.66%` -> `72.59%` (`+4.93` pp)

## Residual Highest-Impact Candidates for 04-07

Excluding this plan's targeted modules:

1. `docir-parser/src/odf/spreadsheet.rs` (`538` lines missed)
2. `docir-parser/src/odf/ods.rs` (`407` lines missed)
3. `docir-parser/src/odf/presentation_helpers.rs` (`324` lines missed)
4. `docir-parser/src/odf/helpers.rs` (`302` lines missed)
5. `docir-parser/src/odf/styles_support.rs` (`261` lines missed)
