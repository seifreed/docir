# External Integrations Audit

## Scope
- Focus: external system touchpoints and protocol/library integrations
- Includes runtime integrations and build/distribution integrations
- Excludes internal crate-to-crate calls unless they expose external boundary behavior

## 1) Integration Inventory (At a Glance)

| Integration | Type | Direction | Purpose | Key Evidence |
|---|---|---|---|---|
| Local filesystem | Runtime IO | In/Out | Read input docs, read optional rules profile, write results | `crates/docir-parser/src/input.rs:25-27`, `crates/docir-cli/src/commands/util.rs:64-83`, `crates/docir-cli/src/commands/rules.rs:18-20` |
| ZIP container parsing | Runtime format integration | Inbound | Parse OOXML/ODF/HWPX zip containers with safety limits | `crates/docir-parser/src/zip_handler.rs:57-125` |
| XML standards parsing | Runtime format integration | Inbound | Parse OOXML/ODF/HWPX XML parts via streaming parser | `crates/docir-parser/src/xml_utils.rs:2-3`, `crates/docir-parser/src/ooxml/docx/document.rs:17-18`, `crates/docir-parser/src/odf/mod.rs:23-24` |
| OLE/CFB parsing (HWP legacy) | Runtime format integration | Inbound | Parse legacy HWP container streams | `crates/docir-parser/src/parser/document.rs:47-49`, `crates/docir-parser/src/hwp/io.rs:3` |
| XLSB via calamine | Runtime format integration | Inbound | Parse binary Excel sheets | `crates/docir-parser/src/parser/parser_xlsx.rs:26-37` |
| Crypto primitives | Runtime security integration | Inbound processing | Decrypt/check protected document parts, hash streams | `crates/docir-parser/src/odf/container.rs:231-260`, `crates/docir-parser/src/hwp/io.rs:115-146` |
| Python binding (PyO3) | Runtime language bridge | In/Out | Expose parsing/rules/query/summary into Python | `crates/docir-python/Cargo.toml:11,22`, `crates/docir-python/src/lib.rs:157-164` |
| CLI surface | Runtime user interface | In/Out | Command-based usage and JSON/text outputs | `crates/docir-cli/src/cli.rs:110-320`, `crates/docir-cli/src/main.rs:15-20` |
| Cargo/crates.io | Build supply chain | Inbound (build-time) | Resolve and lock third-party dependencies | `Cargo.toml:23-49`, `Cargo.lock` crates.io registry entries |
| GitHub repository metadata | Project/distribution | Outbound reference | Repository and package references in metadata/docs | `Cargo.toml:20`, `README.md:12-15` |

## 2) Runtime IO Integrations

### 2.1 Filesystem Input
- Input documents are opened from disk through parser entry helpers.
  - Evidence: `crates/docir-parser/src/input.rs:25-27`
- Application facade exposes `parse_file` path-based API used by CLI and Python wrappers.
  - Evidence: `crates/docir-app/src/lib.rs:236-239`, `crates/docir-python/src/lib.rs:139-142`

### 2.2 Filesystem Output
- CLI writes JSON either to stdout or file; text output follows same model.
  - Evidence: `crates/docir-cli/src/commands/util.rs:64-89`
- Coverage command exports report payloads to files.
  - Evidence: `crates/docir-cli/src/commands/coverage.rs:209-216`

## 3) Document Container/Format Integrations

### 3.1 ZIP Packages (OOXML/ODF/HWPX)
- Uses `zip::ZipArchive` with explicit protections:
  - file-count guard
  - max file size guard
  - max total expanded size guard
  - compression ratio guard (zip bomb mitigation)
  - path traversal/depth guard
  - Evidence: `crates/docir-parser/src/zip_handler.rs:58-125`, `crates/docir-parser/src/zip_handler.rs:135-154`

### 3.2 Format Detection and Dispatch
- Detects RTF, OLE/CFB, and ZIP; then dispatches to OOXML/ODF/HWPX parser path.
  - Evidence: `crates/docir-parser/src/parser/document.rs:43-83`

### 3.3 Standard/Namespace Parsing
- OOXML package markers and part schemas (OpenXML package content types and relationships) are parsed/recognized.
  - Evidence: `crates/docir-parser/src/parser/document.rs:58-63`, test fixtures in `crates/docir-parser/src/parser/tests/helpers.rs:19-34`
- ODF manifest/content/signatures processing is integrated.
  - Evidence: `crates/docir-parser/src/odf/container.rs:24-31`, `crates/docir-parser/src/odf/container.rs:57-63`

## 4) Security and Crypto Integrations

### 4.1 ODF Encrypted Parts
- Integrates AES-CBC + PBKDF2 + SHA1 to decrypt encrypted ODF parts when password/encryption metadata exist.
  - Evidence: `crates/docir-parser/src/odf/container.rs:3-8`, `crates/docir-parser/src/odf/container.rs:231-260`

### 4.2 HWP Encrypted Streams
- Integrates SHA1-derived key + AES-128-CBC decryption path for encrypted legacy HWP streams.
  - Evidence: `crates/docir-parser/src/hwp/io.rs:5-10`, `crates/docir-parser/src/hwp/io.rs:115-146`

### 4.3 Hashing/Indicators
- Uses SHA-256 for stream hashing and analysis support.
  - Evidence: `crates/docir-parser/src/hwp/io.rs:6`, `crates/docir-parser/src/hwp/io.rs:70-81`

## 5) Language/Platform Integrations

### 5.1 Python (PyO3)
- Exposes Rust functionality into Python module `docir` as a native extension.
- Published API functions: `parse_json`, `rules`, `query`, `summary`.
  - Evidence: `crates/docir-python/src/lib.rs:18-129`, `crates/docir-python/src/lib.rs:157-164`
- ABI strategy: stable `abi3` targeting Python 3.7+.
  - Evidence: `crates/docir-python/Cargo.toml:22`

### 5.2 CLI Integration Boundary
- Command-based API with subcommands for parse/security/coverage/diff/rules/query/etc.
  - Evidence: `crates/docir-cli/src/cli.rs:110-320`
- Structured logging initialized via `env_logger`.
  - Evidence: `crates/docir-cli/src/main.rs:16`

## 6) Build and Distribution Integrations
- Build and dependency resolution through Cargo workspace and crates.io ecosystem.
  - Evidence: `Cargo.toml:1-13`, `Cargo.toml:23-49`, `Cargo.lock`
- Repository metadata links GitHub as canonical upstream.
  - Evidence: `Cargo.toml:20`
- README advertises crates.io and Python package install flows.
  - Evidence: `README.md:12`, `README.md:231`

## 7) Explicit Non-Integrations (Current State)
- No direct outbound network clients detected in runtime crates (`reqwest`, `hyper`, `tokio` absent from manifests/imports).
- No database connectivity layer detected (`sqlx`, `diesel`, `rusqlite`, `mongodb`, `redis` absent).
- No cloud SDK integrations detected (AWS/GCP/Azure SDK crates absent).

## 8) Integration Risk Notes (Practical)
- Highest-risk external boundary is untrusted document parsing (archive/XML/binary formats); mitigations are present in zip safety limits and parser size guards.
  - Evidence: `crates/docir-parser/src/zip_handler.rs:11-35`, `crates/docir-parser/src/input.rs:34-44`
- Python bridge broadens consumption surface; exceptions are normalized to `PyValueError` for predictable host behavior.
  - Evidence: `crates/docir-python/src/lib.rs:153-155`
