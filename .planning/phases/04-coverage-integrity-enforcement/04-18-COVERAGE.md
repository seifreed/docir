# 04-18 Coverage Evidence

## Canonical Commands

```bash
cargo llvm-cov --workspace --all-features --summary-only --fail-under-lines 95
```

## Workspace Total (Line Coverage)

- Baseline from 04-17 canonical fail-under total: `78.26%`
- 04-18 canonical fail-under run total: `78.75%`
- Net delta vs 04-17 baseline: `+0.49` percentage points

## 95% Gate Status (Canonical Truth)

- Command: `cargo llvm-cov --workspace --all-features --summary-only --fail-under-lines 95`
- Exit code: `1` (FAIL)
- Observed total: `78.75%`
- Required threshold: `95.00%`
- Remaining gap: `16.25` percentage points

## Cross-Crate Hotspot Snapshot (post-04-18)

Ranked by missed lines:

1. `docir-parser/src/ooxml/docx/document/inline.rs` - `193` missed
2. `docir-parser/src/rtf/core.rs` - `180` missed
3. `docir-parser/src/ooxml/pptx.rs` - `172` missed
4. `docir-parser/src/ooxml/docx/document/table.rs` - `158` missed
5. `docir-parser/src/ooxml/pptx/metadata.rs` - `140` missed
6. `docir-parser/src/odf/helpers.rs` - `136` missed
7. `docir-parser/src/ole.rs` - `128` missed
8. `docir-parser/src/ooxml/docx/document/paragraph.rs` - `126` missed
9. `docir-parser/src/odf/ods.rs` - `123` missed
10. `docir-parser/src/rtf/core/controls.rs` - `121` missed

## Wave Notes

- `docir-core/src/visitor/visitors.rs` reached `100%` lines in workspace summary.
- `docir-diff/src/summary.rs` moved to `94.82%` lines.
- Parser/rules utility hotspots improved materially:
  - `rtf/objects.rs`: `44.83% -> 92.86%`
  - `zip_handler.rs`: `70.00% -> 94.87%`
  - `rules/profile.rs`: `45.28% -> 85.98%`
- `docir-python/src/lib.rs` remains `0.00%` because unit tests that exercise bindings fail to link in this environment (missing `_Py*` symbols in test binary).
