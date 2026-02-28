<p align="center">
  <img src="https://img.shields.io/badge/docir-Document%20Security%20Analysis-red?style=for-the-badge" alt="docir">
</p>

<h1 align="center">docir</h1>

<p align="center">
  <strong>Security-focused Document Intermediate Representation toolkit for Office malware analysis</strong>
</p>

<p align="center">
  <a href="https://crates.io/crates/docir"><img src="https://img.shields.io/crates/v/docir?style=flat-square&logo=rust&logoColor=white" alt="Crates.io Version"></a>
  <a href="https://github.com/seifreed/docir/blob/main/LICENSE"><img src="https://img.shields.io/badge/license-GPL--3.0-green?style=flat-square" alt="License"></a>
  <a href="https://github.com/seifreed/docir/actions"><img src="https://img.shields.io/github/actions/workflow/status/seifreed/docir/ci.yml?style=flat-square&logo=github&label=CI" alt="CI Status"></a>
  <img src="https://img.shields.io/badge/rust-2021-orange?style=flat-square&logo=rust" alt="Rust Edition">
</p>

<p align="center">
  <a href="https://github.com/seifreed/docir/stargazers"><img src="https://img.shields.io/github/stars/seifreed/docir?style=flat-square" alt="GitHub Stars"></a>
  <a href="https://github.com/seifreed/docir/issues"><img src="https://img.shields.io/github/issues/seifreed/docir?style=flat-square" alt="GitHub Issues"></a>
  <a href="https://buymeacoffee.com/seifreed"><img src="https://img.shields.io/badge/Buy%20Me%20a%20Coffee-support-yellow?style=flat-square&logo=buy-me-a-coffee&logoColor=white" alt="Buy Me a Coffee"></a>
</p>

---

## Overview

**docir** is a security-focused Rust toolkit that transforms Office documents into a semantic Intermediate Representation (IR) - conceptually "LLVM IR for documents". It enables static malware analysis, threat detection, and document forensics without executing any code.

### Key Features

| Feature | Description |
|---------|-------------|
| **Multi-format Support** | Parse DOCX, XLSX, PPTX, ODT, ODS, ODP, HWP, HWPX, RTF |
| **Security Analysis** | Detect VBA macros, DDE, XLM macros, OLE objects, ActiveX |
| **Threat Indicators** | Auto-exec triggers, suspicious API calls, external references |
| **Semantic IR** | 79+ strongly-typed node types for document representation |
| **Safe Parsing** | ZIP bomb protection, recursion limits, size constraints |
| **Python Bindings** | PyO3-based Python library for integration |
| **Diff Engine** | Structural comparison between documents |
| **Rule Engine** | Customizable security rules and profiles |

### Supported Document Formats

```
Microsoft Office    DOCX, XLSX, PPTX (and macro-enabled variants)
OpenDocument        ODT, ODS, ODP
Hangul              HWP (legacy binary), HWPX (XML)
Legacy              RTF
```

### Threat Detection Capabilities

```
VBA Macros          Auto-exec procedures, suspicious API calls, obfuscation
XLM Macros          Excel 4.0 macros, dangerous functions (EXEC, CALL, RUN)
DDE Fields          Dynamic Data Exchange, auto-update triggers
OLE Objects         Embedded/linked objects with SHA-256 hashing
ActiveX Controls    CLSID identification, property extraction
External References Remote templates, hyperlinks, data connections
```

---

## Installation

### From Source (Recommended)

```bash
git clone https://github.com/seifreed/docir.git
cd docir
cargo build --release
```

The binary will be available at `target/release/docir`.

### From Crates.io

```bash
cargo install docir
```

---

## Quick Start

```bash
# Parse a document to IR (JSON format)
docir parse document.docx --pretty

# Security analysis
docir security suspicious.xlsm --verbose

# Quick summary
docir summary report.pptx

# Compare two documents
docir diff original.docx modified.docx
```

---

## Usage

### Command Line Interface

```bash
# Parse document to JSON IR
docir parse document.docx --format json --pretty --output ir.json

# Security analysis with detailed findings
docir security malware.xlsm --verbose --json

# Document summary
docir summary report.pptx

# Parser coverage analysis
docir coverage document.xlsx --details --inventory

# Query specific elements
docir query document.docx --predicate macros
docir query document.xlsx --predicate external-refs

# Semantic text search
docir grep "password" document.docx

# Run security rules
docir rules document.docm --profile strict

# Extract specific nodes
docir extract document.docx --node-type MacroModule
```

### Available Commands

| Command | Description |
|---------|-------------|
| `parse` | Parse document and output IR in JSON format |
| `security` | Perform security analysis and threat detection |
| `summary` | Quick document overview |
| `coverage` | Report parser coverage and content-type inventory |
| `diff` | Structural comparison between two documents |
| `rules` | Run security rule engine |
| `query` | Query IR with predicates (macros, external-refs, etc.) |
| `grep` | Semantic text search across document |
| `extract` | Extract nodes by ID or type |
| `dump-node` | Dump specific IR node by ID |

### Quality Workflow (Canonical Gate)

Final quality acceptance is authorized only through the canonical command:

```bash
./scripts/quality_gate.sh
```

Direct invocations such as `cargo fmt`, `cargo clippy`, or `cargo test` are non-authoritative for acceptance, even when they pass locally.

Policy details: [Quality Gate Non-Bypass Policy](docs/quality-gate-policy.md)

### Security Options

| Option | Description |
|--------|-------------|
| `--verbose` | Show detailed threat findings |
| `--json` | Output in JSON format |

### Parser Configuration

| Option | Description |
|--------|-------------|
| `--zip-max-total-size` | Maximum total uncompressed size |
| `--zip-max-file-size` | Maximum per-file size |
| `--zip-max-file-count` | Maximum number of files in archive |
| `--zip-max-compression-ratio` | ZIP bomb detection threshold |
| `--max-input-size` | Maximum input document size |

### Format-Specific Options

| Option | Description |
|--------|-------------|
| `--odf-password` | Password for encrypted ODF documents |
| `--odf-fast` | Fast mode for large ODF files |
| `--hwp-password` | Password for encrypted HWP documents |
| `--hwp-dump-streams` | Dump HWP internal streams |

---

## Security Analysis

### Threat Levels

| Level | Description |
|-------|-------------|
| **Critical** | VBA macros with auto-exec, XLM macros |
| **High** | OLE objects, DDE fields, ActiveX controls |
| **Medium** | External templates, remote resources |
| **Low** | Standard hyperlinks |
| **None** | No security concerns detected |

### Threat Indicators

| Indicator | Description |
|-----------|-------------|
| `AutoExecMacro` | Auto-execution procedures (AutoOpen, Document_Open, etc.) |
| `SuspiciousApiCall` | Shell, WScript, CreateObject, PowerShell, etc. |
| `ExternalTemplate` | Remote template injection |
| `DdeCommand` | Dynamic Data Exchange commands |
| `OleObject` | Embedded OLE objects |
| `ActiveXControl` | ActiveX control instances |
| `XlmMacro` | Excel 4.0 macro formulas |
| `HiddenMacroSheet` | Hidden macro sheets in Excel |
| `SuspiciousFormula` | Potentially malicious formulas |

### Suspicious VBA API Calls Detected

```
Shell Execution     Shell, WScript.Shell, ShellExecute
Process Control     CreateObject, GetObject, CallByName
File System         FileSystemObject, Scripting.FileSystemObject
Network             XMLHTTP, WinHTTP, MSXML2, InternetExplorer
Registry            RegRead, RegWrite, RegDelete
Windows API         Declare Function, Declare Sub
PowerShell          PowerShell, Invoke-Expression
Obfuscation         Chr, ChrW, Base64, StrReverse
Environment         Environ
```

### Dangerous XLM Functions Detected

```
EXEC, CALL, REGISTER, RUN, FOPEN, FWRITE, FREAD, FCLOSE,
URLDOWNLOADTOFILE, ALERT, HALT, FORMULA, SET.VALUE, SET.NAME
```

---

## Python Library

### Installation

```bash
pip install docir
```

### Basic Usage

```python
import docir

# Parse document to IR
ir = docir.parse("document.docx")

# Security analysis
security_info = docir.analyze_security("suspicious.xlsm")
print(f"Threat level: {security_info.threat_level}")

for indicator in security_info.threat_indicators:
    print(f"  - {indicator.indicator_type}: {indicator.description}")

# Check for macros
if security_info.macro_project:
    for module in security_info.macro_project.modules:
        print(f"Module: {module.name}")
        if module.suspicious_api_calls:
            print(f"  Suspicious calls: {module.suspicious_api_calls}")
```

---

## IR Node Types

docir uses 79+ strongly-typed IR nodes organized by category:

### Document Structure
```
Document, Section, Paragraph, Run, Text, Hyperlink,
Table, TableRow, TableCell
```

### Presentation (PPTX)
```
Slide, Shape, TextFrame, SlideMaster, SlideLayout,
NotesMaster, HandoutMaster, NotesSlide
```

### Spreadsheet (XLSX)
```
Worksheet, Cell, Formula, SharedStringTable, DefinedName,
ConditionalFormat, DataValidation, PivotTable, CalcChain
```

### Security-Related
```
MacroProject, MacroModule, OleObject, ExternalReference,
ActiveXControl, DdeField, XlmMacro
```

### Metadata & Styling
```
Metadata, CustomProperty, StyleSet, NumberingSet, Theme
```

---

## Architecture

docir is organized as a Rust workspace with 9 specialized crates:

| Crate | Description |
|-------|-------------|
| `docir-core` | Core IR definitions and types |
| `docir-parser` | OOXML, ODF, HWP, HWPX, RTF parsers |
| `docir-security` | Security analysis and threat indicators |
| `docir-serialization` | IR serialization (JSON) |
| `docir-cli` | Command-line interface |
| `docir-app` | Application layer orchestration |
| `docir-diff` | Structural diff engine |
| `docir-rules` | Security rule engine |
| `docir-python` | Python bindings (PyO3) |

---

## Examples

### Analyze a Suspicious Document

```bash
docir security malware.xlsm --verbose --json | jq
```

### Extract VBA Macro Code

```bash
docir parse macro.docm --pretty | jq '.nodes[] | select(.type == "MacroModule")'
```

### Find External References

```bash
docir query document.docx --predicate external-refs
```

### Compare Document Versions

```bash
docir diff original.xlsx modified.xlsx
```

### Export Parser Coverage

```bash
docir coverage document.pptx --export coverage.csv --export-format csv
```

### Batch Processing

```bash
for doc in *.docx; do
    echo "=== $doc ==="
    docir security "$doc" --json
done
```

---

## Requirements

- Rust 2021 edition (1.60+)
- See [Cargo.toml](Cargo.toml) for full dependency list

### Key Dependencies

```
Parsing         quick-xml, zip, calamine, encoding_rs
Cryptography    sha2, sha1, aes, pbkdf2
CLI             clap, anyhow
Serialization   serde, serde_json
Python          pyo3 (optional)
```

---

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

---

## Support the Project

If you find docir useful, consider supporting its development:

<a href="https://buymeacoffee.com/seifreed" target="_blank">
  <img src="https://cdn.buymeacoffee.com/buttons/v2/default-yellow.png" alt="Buy Me A Coffee" height="50">
</a>

---

## License

This project is licensed under the GPL-3.0 License - see the [LICENSE](LICENSE) file for details.

**Attribution Required:**
- Author: **Marc Rivero López** | [@seifreed](https://github.com/seifreed)
- Repository: [github.com/seifreed/docir](https://github.com/seifreed/docir)

---

<p align="center">
  <sub>Made with dedication for the malware analysis and threat intelligence community</sub>
</p>
