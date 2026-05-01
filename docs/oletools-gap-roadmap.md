# Remaining oletools Gap Roadmap

This roadmap covers the remaining gaps after the current `docir` work on:

- `probe-format`
- `list-times`
- `inspect-metadata`
- `inspect-directory`
- `inspect-sectors`
- `report-indicators`
- `extract-links`

It is focused on the remaining areas that are still only partially covered or not covered when compared with the practical analyst value of `oletools`.

## 1. Functional Audit Snapshot

### Covered Well

- format/container triage
- CFB/OLE timestamps
- basic OLE metadata extraction
- low-level directory inspection
- low-level FAT/MiniFAT/sector inspection
- structural corruption scorecard
- dedicated DDE-style link extraction

### Covered Partially

- structural taxonomy and scoring are now stable, but still narrower than a full forensic suite
- `inspect-metadata` now covers a broader useful classic-property set, but still not a near-exhaustive catalog
- `inspect-sheet-records` now covers minimal low-level XLS BIFF record listing, but not deeper BIFF semantics or XLS formula/XLM behavior
- `inspect-slide-records` now covers minimal low-level PPT binary record listing, but not richer container semantics or deeper legacy PowerPoint parsing
- the current command set is strong on CFB triage and structural inspection, but still not equivalent to every specialized `oletools` utility

### Not Covered

- legacy XLS/PPT binary record introspection at useful low-level depth
- deep SWF extraction and analysis
- full replacement for the whole long-tail of specialized `oletools` utilities

## 2. Phase Roadmap

### Phase A. Consolidate Partial Coverage

#### A1. Expand classic OLE metadata property coverage

Status:

- partially completed

Goal:

- improve `inspect-metadata` so it covers a wider classic property set without changing the basic command model

What to add:

- more SummaryInformation IDs
- more DocumentSummaryInformation IDs
- clearer mapping for common legacy Office fields:
  - author
  - last saved by
  - template
  - revision number
  - application name
  - company
  - manager
  - category
  - keywords
  - comments
  - byte/page/word counts where present

Acceptance:

- JSON output uses stable property names
- unknown IDs remain visible instead of being dropped
- formatter text remains readable for mixed typed values

Files likely involved:

- `crates/docir-app/src/metadata.rs`
- `crates/docir-cli/src/commands/inspect_metadata.rs`
- `crates/docir-app/src/test_support.rs`

#### A2. Audit and reduce signal overlap in `report-indicators`

Goal:

- keep the current structural indicators useful without producing redundant analyst noise

What to review:

- `cfb-structural-anomalies`
- `cfb-structural-score`
- `cfb-directory-score`
- `cfb-sector-score`
- `cfb-stream-score`
- `cfb-dominant-anomaly-class`

Acceptance:

- each indicator answers a different question
- evidence strings use one stable taxonomy
- no two indicators are effectively duplicates with different names

Files likely involved:

- `crates/docir-app/src/report_indicators.rs`
- `crates/docir-cli/src/commands/report_indicators.rs`

### Phase B. Legacy Binary Record Introspection

#### B1. Add low-level XLS record inspection

Status:

- partially completed

Goal:

- cover the most important missing legacy-binary inspection capability

Command proposal:

- `inspect-sheet-records`

Minimum useful scope:

- enumerate record headers from legacy workbook streams
- report:
  - record type
  - offset
  - size
  - substream grouping when possible

Out of scope for this phase:

- semantic formula interpretation
- XLM execution/deobfuscation
- full BIFF semantic parser

Acceptance:

- works on synthetic or minimal `.xls` fixtures
- JSON and text outputs are stable
- malformed streams degrade gracefully

Current state:

- low-level BIFF header walking is implemented
- current scope reports:
  - record type
  - offset
  - size
  - substream grouping derived from BOF records
- parser-only validation has completed successfully
- still missing:
  - broader BIFF record-name coverage
  - richer low-level `.xls` fixtures
  - clean CLI validation pass in an environment without the intermittent `rustc` target-test/link hang when `docir-parser` is rebuilt as a dependency

#### B2. Add low-level PPT record inspection

Status:

- partially completed

Goal:

- provide the same class of low-level introspection for binary PowerPoint

Command proposal:

- `inspect-slide-records`

Minimum useful scope:

- enumerate record containers and leaf records
- report:
  - type
  - offset
  - size
  - nesting/container context when available

Out of scope:

- full semantic rendering of slide content
- full parser parity with specialized PowerPoint tooling

Acceptance:

- stable listing for minimal `.ppt` fixtures
- malformed inputs do not crash inspection

Current state:

- low-level record-header walking is implemented for `PowerPoint Document`
- current scope reports:
  - record type
  - offset
  - length
  - container bit
  - nesting depth
- parser-only validation has completed successfully
- still missing:
  - broader PowerPoint record-name coverage
  - richer `.ppt` fixtures
  - clean CLI validation pass in an environment without the intermittent `rustc` target-test/link hang when `docir-parser` is rebuilt as a dependency

Likely files for Phase B:

- new low-level reader module under `crates/docir-parser/src/`
- new app-layer wrappers in `crates/docir-app/src/`
- new CLI commands in `crates/docir-cli/src/commands/`

### Phase C. SWF Extraction

#### C1. Add embedded SWF detection/extraction

Status:

- partially completed

Goal:

- cover the explicit SWF gap

Command proposal:

- `extract-flash`

Minimum useful scope:

- detect `FWS` / `CWS`
- report:
  - signature
  - version
  - declared size
  - hash
  - origin path/container
- optionally export raw payload

Candidate sources:

- OLE payloads
- RTF object data
- OOXML embedded payloads where applicable

Acceptance:

- JSON and text outputs are stable
- extracted payload is reproducible
- compressed `CWS` is at least recognized even if not fully decoded in first phase

Current state:

- `extract-flash` is implemented with dedicated app/CLI surfaces
- current scope detects:
  - `FWS`
  - `CWS`
  - `ZWS`
- current scope reports:
  - signature
  - compression family
  - version
  - declared size
  - extracted size
  - truncation
  - hash
  - source path
- current scope can export raw SWF payloads with `--out`
- still missing:
  - decompression/normalization of compressed SWF payloads
  - richer fixture coverage across OOXML/RTF/OLE variants
  - clean CLI validation pass in an environment without the intermittent `rustc` target-test/link hang

### Phase D. Remaining Specialized Utility Gap

Goal:

- decide explicitly what `docir` will and will not replace

What to do:

- maintain a capability matrix against conceptual `oletools` use-cases
- classify each use-case as:
  - covered
  - partially covered
  - intentionally out of scope

Reason:

- full 1:1 replacement is not a good engineering target if the command surface starts drifting into low-value long-tail utilities

## 3. Test Plan

### Required E2E Coverage

- JSON output for:
  - `probe-format`
  - `inspect-metadata`
  - `inspect-directory`
  - `inspect-sectors`
  - `report-indicators`
- text output for:
  - `inspect-metadata`
  - `inspect-directory`
  - `inspect-sectors`
  - `report-indicators`

### Required Additional Low-Level Fixtures

- `inspect-directory`
  - mixed 3-cycles
  - dead references from multiple source types
  - orphaned `ObjectPool` entries
- `inspect-sectors`
  - high-severity shared chains
  - invalid-start on main streams
  - truncated chains by allocation
  - MiniFAT incoherence with small streams
- `inspect-metadata`
  - more property IDs and mixed value types
- future record-inspection commands
  - minimal valid `.xls`
  - minimal valid `.ppt`

## 4. Command-by-Command Gap Summary

| Command | Current State | Main Gap |
|---|---|---|
| `probe-format` | Good | no major short-term gap |
| `list-times` | Good | narrow by design |
| `inspect-metadata` | Partial | classic property-name coverage is improved, but still not exhaustive |
| `inspect-directory` | Good | more fixtures, not more surface |
| `inspect-sectors` | Good | more fixtures, not more surface |
| `report-indicators` | Partial | signal overlap discipline and scope boundaries |
| `extract-links` | Partial | still narrower than every DDE edge case in the wider ecosystem |
| `inspect-sheet-records` | Partial | minimal BIFF listing exists; needs richer fixtures/names and clean validation |
| `inspect-slide-records` | Partial | minimal PPT listing exists; needs richer fixtures/names and clean validation |
| `extract-flash` | Partial | dedicated SWF extraction exists; needs richer fixtures and clean validation |

## 5. Remaining Coverage Matrix vs oletools Gap

### Covered

- format/container identification
- CFB timestamp listing
- basic OLE property-set metadata
- CFB directory inspection
- CFB sector/FAT/MiniFAT inspection
- structural anomaly reporting
- dedicated DDE-style extraction

### Covered Partially

- structural taxonomy and scoring
- classic metadata property-name coverage, now broader but still incomplete
- overall practical analyst parity for CFB-focused triage
- low-level XLS/PPT/SWF coverage is implemented but remains unvalidated in this environment because the documented target-test/link route still hangs intermittently

### Not Covered

- deep low-level XLS binary records
- deep low-level PPT binary records
- deep SWF extraction/analysis
- full replacement for all specialized `oletools` utilities

## 6. Recommended Execution Order

1. finish remaining `inspect-metadata` property coverage gaps
2. trim overlap in `report-indicators`
3. add/finish missing regression fixtures for `inspect-directory` and `inspect-sectors`
4. deepen `inspect-sheet-records`
5. deepen `inspect-slide-records`
6. deepen `extract-flash`
7. update capability matrix after each phase

## 7. Recommendation

Do not aim for “replace all of `oletools`” as a single milestone.

Use this scope split instead:

- CFB/OLE structural triage: mostly covered already
- classic metadata coverage: small follow-up phase
- legacy binary record introspection: separate medium-size phase
- SWF extraction: separate small phase
- long-tail `oletools` equivalence: explicit accept/reject decisions, not organic drift

## 8. Immediate Value Priorities

1. metadata deeper
   - widen verified classic property names and low-risk typed coverage
2. more fixtures
   - realistic XLS/PPT/SWF fixtures and stable analyst-facing text/JSON outputs
3. future phases out of immediate scope
   - deeper BIFF semantics
   - deeper PPT semantics
   - SWF decompression/semantic parsing
   - full long-tail specialized utility parity

These three blocks should remain the explicit closing split for the current roadmap:

- metadata more deeply
- more fixtures and validation
- future phases outside immediate scope
