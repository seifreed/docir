#!/usr/bin/env bash
set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
source "${REPO_ROOT}/scripts/lib/dependency_allowlist.sh"

scan_toml_dependencies() {
  awk '
    /^\[dependencies\]/ { in_deps = 1; next }
    /^\[[^\]]+\]/ {
      if (in_deps) in_deps = 0
      next
    }
    in_deps && $0 ~ /^[[:space:]]*[A-Za-z0-9_-]+[[:space:]]*=/ {
      dep=$0
      gsub(/^[[:space:]]*/, "", dep)
      sub(/[[:space:]]*=.*/, "", dep)
      print dep
    }
  ' "$1"
}

check_dependency_policy() {
  local crate_file="$1"
  local crate_name
  crate_name="$(basename "$(dirname "${crate_file}")")"
  local has_violation=0
  local dep

  while IFS= read -r dep; do
    [ -z "$dep" ] && continue
    if ! is_allowed_dependency "$crate_name" "$dep"; then
      printf 'Dependency policy violation: %s -> %s\n' "$crate_name" "$dep"
      has_violation=1
    fi
  done < <(scan_toml_dependencies "${crate_file}")

  return "${has_violation}"
}

crate_has_dependency() {
  local crate_file="$1"
  local expected="$2"
  local dep

  while IFS= read -r dep; do
    [ -z "${dep}" ] && continue
    if [ "${dep}" = "${expected}" ]; then
      return 0
    fi
  done < <(scan_toml_dependencies "${crate_file}")

  return 1
}

check_forbidden_import_pattern() {
  local target="$1"
  local pattern="$2"
  local label="$3"
  local matches

  if rg -n -q -P "${pattern}" "${target}"; then
    echo "Forbidden ${label} imports detected:"
    rg -n -P "${pattern}" "${target}" || true
    return 1
  fi

  return 0
}

check_forbidden_dependency_pattern() {
  local crate_file="$1"
  local pattern="$2"
  local label="$3"
  local crate_name="$4"
  local violation=0

  if scan_toml_dependencies "${crate_file}" | rg -q -P "${pattern}"; then
    echo "${label} dependency violation in ${crate_name}:"
    scan_toml_dependencies "${crate_file}" | rg -n -P "${pattern}" | sed -E "s/^/${crate_name}-> /"
    violation=1
  fi

  return "${violation}"
}

check_forbidden_imports() {
  local file="$1"
  local pattern="$2"
  rg -n --color=never -P "$pattern" "${file}" || true
}

check_frontier_contracts() {
  local status=0
  local app_lib="${REPO_ROOT}/crates/docir-app/src/lib.rs"
  local app_ports="${REPO_ROOT}/crates/docir-app/src/ports.rs"
  local app_adapters="${REPO_ROOT}/crates/docir-app/src/adapters.rs"

  if ! rg -q "pub trait ParserPort" "${app_ports}"; then
    echo "Missing boundary contract: ParserPort in crates/docir-app/src/ports.rs"
    status=1
  fi

  if ! rg -q "pub trait SerializerPort" "${app_ports}"; then
    echo "Missing boundary contract: SerializerPort in crates/docir-app/src/ports.rs"
    status=1
  fi

  if ! rg -q "pub trait SummaryPresenterPort" "${app_ports}"; then
    echo "Missing boundary contract: SummaryPresenterPort in crates/docir-app/src/ports.rs"
    status=1
  fi

  if ! rg -q "impl ParserPort for AppParser" "${app_adapters}"; then
    echo "Missing adapter implementation: impl ParserPort for AppParser in crates/docir-app/src/adapters.rs"
    status=1
  fi

  if ! rg -q "impl ParserPort for DocumentParser" "${app_adapters}"; then
    echo "Missing adapter implementation: impl ParserPort for DocumentParser in crates/docir-app/src/adapters.rs"
    status=1
  fi

  if ! rg -q "impl SecurityScannerPort for AppParser" "${app_adapters}"; then
    echo "Missing adapter implementation: impl SecurityScannerPort for AppParser in crates/docir-app/src/adapters.rs"
    status=1
  fi

  if ! rg -q "pub struct DocirApp<P: ParserPort \\+ SecurityScannerPort" "${app_lib}"; then
    echo "Missing DocirApp parser boundary generic contract in crates/docir-app/src/lib.rs"
    status=1
  fi

  if ! rg -q "ParseDocument::new\\(&self\\.parser, &self\\.parser, self\\.security_enricher\\.as_ref\\(\\)\\)" "${app_lib}"; then
    echo "Missing parser use-case boundary wiring in crates/docir-app/src/lib.rs"
    status=1
  fi

  if ! rg -q "Self::with_parser\\(AppParser::with_config\\(config\\)\\)" "${app_lib}"; then
    echo "Missing default parser port wiring in DocirApp::new (crates/docir-app/src/lib.rs)"
    status=1
  fi

  return "${status}"
}

FAIL=0

echo "layer policy: checking crate dependency boundaries"
while IFS= read -r -d '' crate_toml; do
  if ! check_dependency_policy "${crate_toml}"; then
    FAIL=1
  fi
done < <(find "${REPO_ROOT}/crates" -maxdepth 2 -name Cargo.toml -print0)

echo
echo "layer policy: checking forbidden infrastructure leaks in core/app"

if ! check_forbidden_import_pattern \
  "${REPO_ROOT}/crates/docir-core/src" \
  "(use|extern crate)\\s+(clap|anyhow|env_logger|pyo3|quick[_-]xml|zip::|flate2::|calamine::|serde_json::|tokio::|actix|warp::)" \
  "core"; then
  FAIL=1
fi

if ! check_forbidden_import_pattern \
  "${REPO_ROOT}/crates/docir-app/src" \
  "(use|extern crate)\\s+(clap|anyhow|env_logger|pyo3|serde_json::|serde_yaml::|tokio::|actix|warp::)" \
  "app"; then
  FAIL=1
fi

if rg -n -q "docir-security|docir-parser" "${REPO_ROOT}/crates/docir-cli/src"; then
  echo "Forbidden CLI crate imports detected:"
  check_forbidden_imports "${REPO_ROOT}/crates/docir-cli/src" "docir-(security|parser)"
  FAIL=1
fi

if rg -n -q "docir-(parser|rules|serialization)" "${REPO_ROOT}/crates/docir-python/src"; then
  echo "Forbidden Python boundary import detected:"
  check_forbidden_imports "${REPO_ROOT}/crates/docir-python/src" "docir-(parser|rules|serialization)"
  FAIL=1
fi

if ! check_forbidden_dependency_pattern \
  "${REPO_ROOT}/crates/docir-core/Cargo.toml" \
  "^(clap|anyhow|env_logger|pyo3|quick-xml|zip|flate2|calamine|serde_json|tokio|actix|warp)$" \
  "core infrastructure leak" \
  "docir-core"; then
  FAIL=1
fi

if ! check_forbidden_dependency_pattern \
  "${REPO_ROOT}/crates/docir-app/Cargo.toml" \
  "^(clap|anyhow|env_logger|pyo3|serde_json)$" \
  "app infra/leak" \
  "docir-app"; then
  FAIL=1
fi

if ! check_forbidden_dependency_pattern \
  "${REPO_ROOT}/crates/docir-cli/Cargo.toml" \
  "^(docir-parser|docir-security)$" \
  "cli direct dependency" \
  "docir-cli"; then
  FAIL=1
fi

if crate_has_dependency "${REPO_ROOT}/crates/docir-cli/Cargo.toml" "docir-parser"; then
  echo "Forbidden direct dependency: docir-cli -> docir-parser"
  FAIL=1
fi

if crate_has_dependency "${REPO_ROOT}/crates/docir-cli/Cargo.toml" "docir-security"; then
  echo "Forbidden direct dependency: docir-cli -> docir-security"
  FAIL=1
fi

if ! check_frontier_contracts; then
  FAIL=1
fi

if [ "${FAIL}" -ne 0 ]; then
  echo
  echo "layer policy: FAIL"
  exit 1
fi

echo
echo "layer policy: PASS"
exit 0
