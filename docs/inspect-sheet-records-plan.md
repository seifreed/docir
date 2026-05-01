# inspect-sheet-records Plan

This document scopes the next large missing capability after the current CFB/OLE structural inspection work.

## Goal

Add a low-level analyst-facing command for legacy binary Excel record inspection without turning `docir` into a full BIFF semantic parser.

Proposed command:

- `inspect-sheet-records`

## Why This Phase Exists

Current `docir` coverage is strong for:

- CFB/OLE structural triage
- metadata property sets
- directory/FAT/MiniFAT inspection
- structural indicator reporting

The largest remaining low-level analyst gap is legacy XLS binary record introspection.

## In Scope

- read legacy workbook streams from `.xls`
- enumerate record headers
- expose:
  - record type
  - offset
  - size
  - substream/workbook grouping when recoverable
- emit stable JSON and readable text output
- tolerate malformed streams gracefully

## Out of Scope

- formula semantic evaluation
- XLM execution or deobfuscation
- full BIFF semantic reconstruction
- VBA AST or behavioral analysis

## Suggested Architecture

### Parser layer

Add a low-level record reader in `docir-parser` that:

- walks workbook bytes sequentially
- decodes BIFF record headers
- emits a neutral record struct

### App layer

Add an app-level inspection report in `docir-app` that:

- groups records into a stable analyst-facing structure
- summarizes counts by record type / substream

### CLI layer

Add a command in `docir-cli`:

- `inspect-sheet-records <file> [--json]`

## Minimum Acceptance Criteria

- works on a minimal valid `.xls` fixture
- JSON output includes record offset/type/size
- text output is readable for analysts
- malformed input does not panic

## Test Plan

- unit tests for header parsing
- fixture test for minimal workbook globals
- fixture test for worksheet substream listing
- CLI JSON/text E2E tests

## Dependencies

- none of the current CFB structural commands need to change
- should reuse current legacy Office dispatch where possible
