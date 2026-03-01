# 04-17 Coverage Evidence

## Canonical Commands

```bash
cargo llvm-cov --workspace --all-features --summary-only --fail-under-lines 95
```

## Workspace Total (Line Coverage)

- Baseline from 04-16 canonical fail-under total: `78.12%`
- 04-17 canonical fail-under run total: `78.26%`
- Net delta vs 04-16 baseline: `+0.14` percentage points

## 95% Gate Status (Canonical Truth)

- Command: `cargo llvm-cov --workspace --all-features --summary-only --fail-under-lines 95`
- Exit code: `1` (FAIL)
- Observed total: `78.26%`
- Required threshold: `95.00%`
- Remaining gap: `16.74` percentage points

## Cross-Crate Hotspot Snapshot (post-04-17)

Ranked by missed lines:

1. `docir-parser/src/ooxml/docx/document/inline.rs` - `193` missed
2. `docir-parser/src/rtf/core.rs` - `179` missed
3. `docir-parser/src/ooxml/pptx.rs` - `172` missed
4. `docir-parser/src/ooxml/docx/document/table.rs` - `158` missed
5. `docir-parser/src/ooxml/pptx/metadata.rs` - `140` missed
6. `docir-parser/src/odf/helpers.rs` - `136` missed
7. `docir-parser/src/ole.rs` - `128` missed
8. `docir-parser/src/ooxml/docx/document/paragraph.rs` - `126` missed
9. `docir-parser/src/odf/ods.rs` - `123` missed
10. `docir-parser/src/rtf/core/controls.rs` - `121` missed

## Outcome

Wave 04-17 raised canonical total to `78.26%` and kept fail-under deterministic (`EXIT:1`). Phase 04 remains quantitatively blocked by parser-heavy hotspots.
