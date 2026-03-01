# 04-14 Coverage Evidence

## Canonical Commands

```bash
cargo llvm-cov --workspace --all-features --summary-only --fail-under-lines 95
cargo llvm-cov report --json --summary-only --output-path target/llvm-cov-summary.json
```

## Workspace Total (Line Coverage)

- Baseline from 04-13 canonical fail-under total: `73.44%`
- After Wave 1 (docir-diff/docir-core hotspot tests): `75.89%` (`EXIT:1`)
- After Wave 2 (CLI summary/security integration tests): `76.61%` (`EXIT:1`)
- Net delta vs 04-13 baseline: `+3.17` percentage points

## 95% Gate Status (Canonical Truth)

- Command: `cargo llvm-cov --workspace --all-features --summary-only --fail-under-lines 95`
- Exit code: `1` (FAIL)
- Observed total: `76.61%`
- Required threshold: `95.00%`
- Remaining gap: `18.39` percentage points

## Cross-Crate Hotspot Snapshot (post-04-14)

Top missed-line files from `target/llvm-cov-summary.json`:

1. `crates/docir-parser/src/ooxml/docx/document/inline.rs` - `193` missed
2. `crates/docir-diff/src/summary.rs` - `189` missed
3. `crates/docir-rules/src/rules.rs` - `187` missed
4. `crates/docir-parser/src/rtf/core.rs` - `179` missed
5. `crates/docir-parser/src/ooxml/pptx.rs` - `172` missed
6. `crates/docir-parser/src/ooxml/docx/document/table.rs` - `158` missed
7. `crates/docir-parser/src/hwp/legacy.rs` - `145` missed
8. `crates/docir-parser/src/ooxml/pptx/metadata.rs` - `140` missed
9. `crates/docir-parser/src/odf/helpers.rs` - `136` missed
10. `crates/docir-parser/src/ole.rs` - `128` missed

## Blocker Characterization

- CC-04 is still quantitatively blocked by large residual debt concentrated in parser core paths (`docx/pptx/rtf/hwp/odf`) and two cross-crate modules (`docir-diff::summary`, `docir-rules::rules`).
- Additional measurable progress is still possible, but 95% requires many more high-miss modules than a single bounded increment.
