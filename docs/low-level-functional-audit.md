# Low-Level Functional Audit

Date:

- 2026-03-10

## Scope

This audit reviews the current analyst-facing low-level utility surface in `docir`:

- `probe-format`
- `list-times`
- `inspect-metadata`
- `inspect-directory`
- `inspect-sectors`
- `report-indicators`
- `extract-links`
- `inspect-sheet-records`
- `inspect-slide-records`
- `extract-flash`

## Functional Status

### Covered Well

- `probe-format`
  - useful triage for CFB, OOXML, ZIP, RTF, PE, PDF, image types and SWF signatures
- `list-times`
  - dedicated CFB FILETIME listing
- `inspect-directory`
  - strong structural visibility over CFB directory state, references, reachability and anomalies
- `inspect-sectors`
  - strong low-level FAT/MiniFAT and stream-chain visibility
- `report-indicators`
  - useful structural scorecard with stable taxonomy and family-specific buckets
- `extract-links`
  - dedicated DDE-style extraction path exists

### Covered Partially

- `inspect-metadata`
  - useful classic property-set extraction
  - still not exhaustive for all classic property IDs
- `inspect-sheet-records`
  - minimal BIFF header walking works
  - useful for offsets, types, sizes and BOF-derived substreams
  - still not a deeper BIFF/XLS semantic inspector
- `inspect-slide-records`
  - minimal PowerPoint binary record walking works
  - useful for offsets, type, length, container bit and nesting depth
  - still not a deeper legacy PowerPoint semantic inspector
- `extract-flash`
  - useful SWF signature-based extraction and raw export
  - still intentionally shallow: no decompression or SWF semantic parsing

## report-indicators Overlap Review

Current conclusion:

- keep the current score layers

Reason:

- `cfb-structural-score` answers overall structural risk
- `cfb-directory-score` answers directory/tree integrity
- `cfb-sector-score` answers allocation/FAT integrity
- `cfb-stream-score` answers stream-level health
- `cfb-dominant-anomaly-class` explains what class dominates the current failure mode

This is still acceptable because each score answers a different question.

What should be avoided next:

- adding another generic structural score
- adding new classes without stable taxonomy prefixes
- turning evidence strings into ad-hoc prose instead of stable buckets

Current redundancy check:

- acceptable overlap:
  - `cfb-structural-score` vs family scores, because one is aggregate and the others are diagnostic
  - `cfb-dominant-anomaly-class` vs evidence buckets, because one explains dominance and the others preserve raw context
- not worth changing now:
  - the current score stack is still interpretable and stable enough for analyst use

### Not Covered

- deep XLS binary semantics
- deep PPT binary semantics
- SWF decompression/semantic analysis
- full long-tail replacement of specialized `oletools` utilities

## Validation Status

### Confirmed

- formatting passes (`cargo fmt`)
- targeted code paths are implemented and wired through parser/app/CLI
- targeted unit and E2E tests exist for XLS, PPT and SWF low-level commands
- parser-only low-level reader tests have completed successfully:
  - `cargo test -p docir-parser xls_records --lib -- --nocapture`
  - `cargo test -p docir-parser ppt_records --lib -- --nocapture`

### Not Fully Confirmed In This Environment

The current environment intermittently hangs during `rustc` link or crate-test startup for some workspace targets.

Observed affected routes:

- `cargo test -p docir-cli inspect_sheet_records -- --nocapture`
- `cargo test -p docir-cli inspect_slide_records -- --nocapture`
- `cargo test -p docir-cli extract_flash -- --nocapture`

Most recent observation:

- the parser-only targets are confirmed green
- the instability most consistently reappears when a CLI target recompiles `docir-parser` as a dependency
- this means the current status is still:
  - implemented
  - partially validated
  - not stable enough to promote to `partial validated` in the roadmap taxonomy

Because of that, the newer low-level XLS/PPT/SWF phases should remain:

- `partial`

and should not yet be promoted to:

- `partial validated`

## Overlap Review: XLS vs PPT Inspectors

There is visible structural similarity between:

- `crates/docir-app/src/inspect_sheet_records.rs`
- `crates/docir-app/src/inspect_slide_records.rs`

and between the corresponding CLI commands.

Decision:

- do not extract shared helpers yet

Reason:

- the duplicate shape is still small
- the domain models differ enough (`substream_kind` vs `depth/container`)
- a shared abstraction now would add coupling faster than it removes maintenance cost

Revisit only if:

- both commands gain another comparable round of summary/count/formatter growth

## Explicit Gap vs oletools

### Covered

- format/container probing
- OLE/CFB time listing
- low-level directory inspection
- low-level FAT/MiniFAT inspection
- structural triage indicators
- basic classic OLE metadata
- basic DDE-style extraction

### Partial

- low-level XLS record inspection
- low-level PPT record inspection
- SWF extraction
- classic property-set completeness

### Not Covered

- deep legacy binary format introspection
- SWF analysis beyond detection/extraction
- complete one-to-one coverage of all specialized `oletools` utilities

## Final Brief: docir vs oletools

Current state:

- `docir` is now strong in structural OLE/CFB triage
- `docir` is useful for analyst-facing low-level inspection of:
  - directory state
  - FAT/MiniFAT/sector allocation
  - classic metadata
  - DDE-style links
  - minimal XLS/PPT binary record walking
  - minimal SWF extraction

Where `docir` is currently better:

- cleaner analyst-focused command naming
- stronger structural CFB visibility
- richer CFB corruption scoring and taxonomy

Where `oletools` still keeps the advantage:

- broader historical depth across specialized utilities
- deeper legacy binary semantics
- more mature coverage in long-tail edge cases

Conclusion:

- `docir` is not yet a full replacement for `oletools`
- `docir` is already competitive for structural triage and modernized analyst workflows around OLE/CFB inspection
- XLS/PPT/SWF remain implemented but not promoted beyond `partial` until the documented validation path runs green in a clean environment

Operational note:

- the current blocker is validation reliability, not missing command surface

## Next Priorities

1. metadata deeper
   - widen classic property-name coverage only where IDs are verified
   - avoid speculative labels
2. more fixtures
   - XLS/PPT/SWF realistic fixtures and visible E2E stability
3. future phases out of immediate scope
   - deeper BIFF/PPT semantics
   - SWF decompression/semantic parsing
   - long-tail utility parity

Metadata note:

- the last metadata pass intentionally stopped at safe, well-documented IDs
- no additional ambiguous property names were added just to increase apparent coverage

This leaves the remaining work explicitly grouped into three blocks:

- metadata more deeply
- more fixtures and validation
- future phases outside immediate scope
