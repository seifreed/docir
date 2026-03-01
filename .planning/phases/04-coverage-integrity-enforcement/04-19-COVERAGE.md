# 04-19 Coverage Evidence

## Canonical Commands

```bash
cargo llvm-cov --workspace --all-features --summary-only --fail-under-lines 95
```

## Workspace Total (Line Coverage)

- Baseline from 04-18 canonical fail-under total: `78.75%`
- 04-19 canonical fail-under run total: `79.22%`
- Net delta vs 04-18 baseline: `+0.47` percentage points

## 95% Gate Status (Canonical Truth)

- Command: `cargo llvm-cov --workspace --all-features --summary-only --fail-under-lines 95`
- Exit code: `1` (FAIL)
- Observed total: `79.22%`
- Required threshold: `95.00%`
- Remaining gap: `15.78` percentage points

## Cross-Crate Hotspot Snapshot (post-04-19)

Ranked by missed lines:

1. `docir-parser/src/ooxml/docx/document/inline.rs` - `193` missed
2. `docir-parser/src/ooxml/pptx.rs` - `172` missed
3. `docir-parser/src/ooxml/docx/document/table.rs` - `158` missed
4. `docir-parser/src/ooxml/pptx/metadata.rs` - `140` missed
5. `docir-parser/src/odf/helpers.rs` - `136` missed
6. `docir-parser/src/ole.rs` - `128` missed
7. `docir-parser/src/ooxml/docx/document/paragraph.rs` - `126` missed
8. `docir-parser/src/odf/ods.rs` - `123` missed
9. `docir-parser/src/rtf/core/controls.rs` - `121` missed
10. `docir-parser/src/hwp/section.rs` - `118` missed

## Wave Notes

- `rtf/core.rs` improved from `73.65%` to `92.11%` line coverage in the canonical summary.
- The blocker on `docir-python/src/lib.rs` remains unchanged (`0.00%`), due to missing Python link symbols in this environment for test binaries.
