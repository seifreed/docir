# VBA Recognition Contract

## Purpose

This contract defines what "VBA recognition" means in Phase 0. It is intentionally narrower than VBA analysis.

## Recognition Definition

A document satisfies VBA recognition when the extractor can determine one or more of the following from document structure alone:

- A VBA project container is present.
- Project-level metadata is present.
- One or more VBA modules or forms are present.
- Source text for one or more VBA-bearing modules can be extracted.

Recognition is based on persisted document content, not on execution, interpretation, or reconstruction of missing code.

## Required Outputs

Phase 0 VBA recognition must output, when available:

- Document identifier and format.
- Extraction status for the VBA project.
- Project container location.
- Project name, if recoverable.
- Module inventory with:
  - module name
  - module kind
  - storage path or stream identifier
  - extraction status
  - source text when recoverable
- Error or limitation details for partial recovery.

The canonical JSON contract is `schema/vba-extraction.schema.json`.

## Module Kinds

Phase 0 recognizes the following categories:

- `standard`
- `class`
- `document`
- `form`
- `unknown`

The extractor may preserve a more specific raw type internally, but cross-component interchange must map to these normalized values.

## Status Semantics

- `recognized`: VBA project or module presence confirmed.
- `extracted`: source bytes or text recovered successfully.
- `partial`: project recognized but one or more expected components were not fully recovered.
- `unsupported`: the container pattern is known but not yet supported by the extractor.
- `error`: extraction attempted and failed.
- `absent`: no VBA project recognized.

## Non-Goals

The following are explicitly outside the VBA recognition contract:

- AST generation.
- Identifier resolution.
- Procedure graph extraction.
- Auto-exec inference from code semantics.
- Deobfuscation of string builders, `Chr`, `StrReverse`, Base64, or similar constructs.
- Macro emulation or runtime simulation.

## Robustness Requirements

- Malformed or encrypted VBA storage must be reported with status and diagnostic detail.
- Missing source must not be represented as an empty successful extraction.
- Unknown module kinds must remain visible as `unknown`; they must not be coerced to a false known type.
- Extracted source text is evidence, not interpretation.

## Interop Rule

Any later phase that performs AST generation, deobfuscation, emulation, or higher-order analysis must consume Phase 0 output as input, not redefine Phase 0 semantics retroactively.
