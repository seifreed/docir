# 04-16 Coverage Evidence

## Canonical Commands

```bash
cargo llvm-cov --workspace --all-features --summary-only --fail-under-lines 95
```

## Workspace Total (Line Coverage)

- Baseline from 04-15 canonical fail-under total: `77.67%`
- 04-16 canonical fail-under run total: `78.12%`
- Net delta vs 04-15 baseline: `+0.45` percentage points

## 95% Gate Status (Canonical Truth)

- Command: `cargo llvm-cov --workspace --all-features --summary-only --fail-under-lines 95`
- Exit code: `1` (FAIL)
- Observed total: `78.12%`
- Required threshold: `95.00%`
- Remaining gap: `16.88` percentage points

## Cross-Crate Hotspot Snapshot (post-04-16)

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

Wave 04-16 reduced legacy HWP parser residuals materially and lifted canonical total to `78.12%`, but fail-under remains non-zero.
