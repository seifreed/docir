# Structural Indicator Taxonomy

This document fixes the analyst-facing taxonomy used by the low-level CFB/OLE inspection commands and `report-indicators`.

## Scope

Applies to:

- `inspect-directory`
- `inspect-sectors`
- `report-indicators`

It does not redefine parser-internal node types. It only fixes the exported analyst vocabulary.

## Score Fields

- `directory_score`
  Aggregated severity for CFB directory/tree corruption.
- `sector_score`
  Aggregated severity for FAT/MiniFAT and sector-allocation corruption.
- `stream_score`
  Aggregated severity for stream-level chain/path corruption.
- `cfb-structural-score`
  Global structural score combining directory, sector and stream evidence.

Allowed values:

- `none`
- `low`
- `medium`
- `high`

## Dominant Anomaly Class

`cfb-dominant-anomaly-class` is the dominant structural class across the container.

Allowed values:

- `shared-sector`
- `cycle`
- `unreachable-live`
- `invalid-start`
- `mini-fat`
- `none`

Tie-break order is stable and explicit:

1. `shared-sector`
2. `cycle`
3. `unreachable-live`
4. `invalid-start`
5. `mini-fat`

## Prefix Families

### `directory:*`

Directory-graph evidence and summaries.

Examples:

- `directory:cycle:sibling-2-cycle=1`
- `directory:reachability:live-unreachable=2`
- `directory:incoming:incoming:state:anomalous=1`

### `sector:*`

Sector/FAT/MiniFAT evidence and summaries.

Examples:

- `sector:shared-sector:0=WordDocument,VBA/PROJECT`
- `sector:truncated-chain:fat:WordDocument=1`
- `sector:structural-incoherence:mini-fat-without-consumers=1 [medium]`

### `health:*`

Chain-health buckets exported from `inspect-sectors`.

Examples:

- `health:shared:root:WordDocument`
- `health:start-reused:allocation:fat`
- `health:invalid-start:root:VBA`

### `dead-reference:*`

References into dead directory slots.

Examples:

- `dead-reference:state:orphaned`
- `dead-reference:source-type:stream`

### `incoming:*`

Incoming-reference summaries in the directory graph.

Examples:

- `incoming:state:normal`
- `incoming:state:anomalous`
- `incoming:source-type:storage`

### `objectpool:*`

ObjectPool-specific corruption evidence.

Examples:

- `objectpool:orphaned:ObjectPool/1/Ole10Native [high]`
- `objectpool:shared:ObjectPool/1/Ole10Native [shared:high]`
- `objectpool:invalid-start:ObjectPool/1/Ole10Native [invalid-start:high]`

### `vba:*`

VBA-structure-specific evidence.

Allowed bucket families:

- `vba:storage`
- `vba:project-stream`
- `vba:module-stream`

Examples:

- `vba:storage:VBA [medium]`
- `vba:project-stream:VBA/PROJECT [shared:high]`
- `vba:module-stream:VBA/Module1 [truncated:medium]`

### `main-stream:*`

Main legacy Office streams.

Allowed bucket families:

- `main-stream:word`
- `main-stream:xls`
- `main-stream:ppt`

Examples:

- `main-stream:word:WordDocument [invalid-start:high]`
- `main-stream:xls:Workbook [shared:high]`
- `main-stream:ppt:PowerPoint Document [truncated:medium]`

## Coverage vs Conceptual oletools Gap

Covered:

- format/container triage
- CFB timestamps
- OLE metadata property sets
- directory-level CFB inspection
- sector/FAT/MiniFAT inspection
- structural indicator scorecard
- DDE-style active link extraction

Covered partially:

- structural anomaly scoring, which is richer than before but still narrower than a full forensic suite
- metadata property-name coverage for some legacy SummaryInformation/DocumentSummaryInformation IDs
- low-level XLS record introspection, now present through `inspect-sheet-records` but still intentionally minimal
- low-level PPT record introspection, now present through `inspect-slide-records` but still intentionally minimal
- SWF extraction, now present through `extract-flash` but still intentionally minimal

Validation note:

- the taxonomy above is considered stable
- XLS/PPT/SWF remain `partial`, not `partial validated`, until the documented
  low-level validation path runs green in a clean environment
- later fixture additions must preserve these prefixes instead of introducing
  near-duplicate families

Not covered:

- deep low-level record introspection for legacy XLS/PPT binaries
- deep SWF extraction/analysis
- dedicated DDE coverage across all oletools-supported edge formats
- full suite replacement for every specialized oletools utility
