# Phase 0 Extraction Scope

## Objective

Phase 0 defines the minimum interoperable scope for document parsing, artifact extraction, and VBA recognition. The goal is to produce deterministic, machine-readable outputs that downstream components can consume without requiring code execution or semantic reconstruction.

## In Scope

Phase 0 includes:

- Parsing document containers and package structures.
- Enumerating extractable artifacts from supported document formats.
- Extracting raw or normalized artifact payloads when available.
- Recognizing the presence of VBA projects.
- Extracting VBA project metadata and module source text when the source is stored in the document.
- Emitting JSON outputs validated by:
  - `schema/artifact-manifest.schema.json`
  - `schema/vba-extraction.schema.json`

## Supported Output Categories

Phase 0 outputs may describe:

- Container/package entries.
- Embedded files and OLE payloads.
- Relationship or part-level extraction targets.
- VBA project metadata.
- VBA modules, class modules, forms, and document modules as extracted source-bearing artifacts.

## Explicitly Out Of Scope

Phase 0 does not include:

- AST generation for VBA or document content.
- VBA token stream to AST conversion.
- Deobfuscation of VBA, XLM, formulas, strings, or shellcode.
- Code emulation, sandboxing, or execution.
- Behavioral scoring or malware classification.
- Control-flow graph reconstruction.
- Semantic interpretation beyond recognition and extraction.

## Contract Boundaries

- Parsing is structural: it locates containers, parts, streams, and metadata boundaries.
- Extraction is representational: it materializes bytes, text, and descriptors for artifacts.
- VBA recognition is declarative: it identifies that a VBA project exists and records what was extracted from it.
- If a VBA project is encrypted, truncated, malformed, or partially recoverable, the output must preserve that status instead of guessing missing semantics.

## Expected Consumer Assumptions

Consumers may rely on Phase 0 outputs for:

- Inventorying what a document contains.
- Persisting extracted artifacts for later analysis.
- Determining whether VBA content is present and which modules were recovered.
- Routing artifacts into later phases that may add AST, deobfuscation, or emulation.

Consumers must not assume:

- Parsed macro semantics.
- Canonicalized VBA syntax.
- Deobfuscated strings.
- Executable intent or runtime behavior.

## Success Criteria

Phase 0 is complete when the implementation can:

- Produce a stable artifact manifest for a document.
- Extract recoverable VBA-related artifacts without executing them.
- Distinguish successful, partial, unsupported, and failed extraction outcomes.
- State exclusions clearly so later phases can extend the pipeline without redefining Phase 0 outputs.
