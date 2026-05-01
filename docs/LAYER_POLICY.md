# Layer Policy

## Scope

This policy defines the allowed dependencies and boundary responsibilities for the whole
`docir` workspace to keep the architecture aligned with Clean Architecture.

## Canonical boundary intent

- **Core (`docir-core`)** contains only business/domain primitives and invariants.
- **Parser (`docir-parser`)** is an infrastructure adapter implementing parsing and
  normalization concerns for concrete formats.
- **Application (`docir-app`)** defines orchestration use cases and port contracts.
- **Interfaces (`docir-cli`, `docir-python`)** orchestrate input/output for the user or
  embedding technologies.
- **Supporting services** (`docir-security`, `docir-rules`, `docir-serialization`,
  `docir-diff`) operate on core domain structures and expose application-level behaviour.

## Allowed dependencies by crate

| Crate | Allowed direct dependencies |
|---|---|
| `docir-core` | `serde` (optional), `thiserror` |
| `docir-parser` | `docir-core`, `zip`, `quick-xml`, `encoding_rs`, `flate2`, `calamine`, `sha2`, `sha1`, `pbkdf2`, `base64`, `aes`, `cbc`, `log`, `serde`, `thiserror` |
| `docir-app` | `docir-core`, `docir-parser`, `docir-security`, `docir-serialization`, `docir-rules`, `docir-diff`, `thiserror` |
| `docir-security` | `docir-core`, `sha2`, `log`, `thiserror` |
| `docir-serialization` | `docir-core`, `serde`, `serde_json`, `thiserror` |
| `docir-rules` | `docir-core`, `serde` |
| `docir-diff` | `docir-core`, `serde`, `serde_json`, `sha2` |
| `docir-cli` | `docir-core`, `docir-app`, `clap`, `anyhow`, `env_logger`, `log`, `serde`, `serde_json` |
| `docir-python` | `docir-core`, `docir-app`, `pyo3`, `serde`, `serde_json`, `anyhow` |

## Forbidden inward leaks

- `docir-core` must not import infrastructure crates (`clap`, `quick-xml`, `zip`,
  `flate2`, `calamine`, `env_logger`, `anyhow`, `pyo3`, `serde_json`, `tokio`, etc.).
- `docir-app` must not import CLI / embedding frameworks (`clap`, `env_logger`, `pyo3`)
  or infra-only parsing internals.
- `docir-cli` must not consume `docir-security` directly (security concerns go through
  app use cases).
- `docir-cli` must not declare direct dependencies on `docir-parser` or `docir-security`.
- `docir-app` and `docir-core` must not import or depend on presentation/output
  format crates (for example `serde_json`) in domain-facing code.
- `docir-parser` may consume `docir-core` only; it must not depend on `docir-app`.
- `docir-python` is an adapter boundary and must not depend directly on parser, rules,
  or serialization infrastructure crates.
- Shared mutable state should cross boundaries through domain objects and explicit ports, not
  raw subsystem globals.

## Required boundary contracts

### Parsing / application boundary

- Parsing entry is an application port (`ParserPort`) implemented by a parser adapter
  (`AppParser`) in `docir-app`.
- Parser-facing logic in `docir-cli` and `docir-python` must use `DocirApp` methods
  and must not invoke parser internals.

### Output boundary

- Serialization and presentation are separated from parsing by application ports:
  - `SerializerPort`
  - `SummaryPresenterPort`
- Serialization adapters may live in application layer adapters and may use external crates.

## Reference cases by frontier

- **CLI → App parsing frontier**
  - `crates/docir-cli/src/commands/util.rs`: `build_app_and_parse`.
  - `crates/docir-app/src/lib.rs`: `parse_file`, `parse_bytes`, `parse_reader`.
- **App → Parser adapter frontier**
  - `crates/docir-app/src/adapters.rs`: `ParserPort` implementation by `AppParser`.
- **App → Core frontier**
  - `crates/docir-app/src/summary.rs`: domain summary is built from `docir_core` IR nodes.
- **Security frontier**
  - `crates/docir-app/src/use_cases.rs`: `AnalyzeSecurityUseCase` and public `analyze_security`.

## Automation and enforcement

`scripts/quality_layer_policy.sh` must pass for the canonical gate. It validates:

1. Dependency direction and whitelist conformance per crate.
2. Explicit forbidden imports in policy-critical crates (`docir-core`, `docir-app`,
   `docir-cli`).
3. Direct dependency prohibition for `docir-cli` against `docir-parser` and
   `docir-security`.
4. Presence of required boundary constructs (`ParserPort`, `SerializerPort`, `SummaryPresenterPort`).

This script is wired into `scripts/quality_gate.sh`.
