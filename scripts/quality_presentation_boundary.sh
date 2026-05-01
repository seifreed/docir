#!/usr/bin/env bash
set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
FAIL=0

check_file_patterns() {
  local target="$1"
  local pattern="$2"
  local name="$3"
  if rg -n -q -P "$pattern" "${target}"; then
    echo "Presentation dependency leak detected in ${name}:"
    rg -n -P "$pattern" "${target}" || true
    return 1
  fi
  return 0
}

check_domain_dependency_policy() {
  local crate_file="$1"
  local crate_name="$2"
  local forbidden=(
    "clap"
    "anyhow"
    "pyo3"
    "env_logger"
    "actix"
    "warp"
    "axum"
    "rocket"
    "serde_json"
  )
  local dep
  local pattern='^[[:space:]]*[A-Za-z0-9_-]+[[:space:]]*='
  local name
  local state

  while IFS= read -r dep; do
    dep="$(echo "${dep}" | sed -E 's/^[[:space:]]+|[[:space:]]+$//g')"
    [ -z "$dep" ] && continue

    for name in "${forbidden[@]}"; do
      if [ "$dep" = "$name" ]; then
        echo "Forbidden domain dependency in ${crate_name} Cargo.toml: ${dep}"
        return 1
      fi
    done
  done < <(
    awk -v pattern="$pattern" '
      /^\[dependencies\]/ { in_deps = 1; next }
      /^\[[^\]]+\]/ {
        if (in_deps) in_deps = 0
        next
      }
      in_deps && $0 ~ pattern {
        dep=$0
        sub(/^[[:space:]]*/, "", dep)
        sub(/[[:space:]]*=.*/, "", dep)
        print dep
      }
    ' "$crate_file"
  )

  return 0
}

if ! check_file_patterns \
  "${REPO_ROOT}/crates/docir-core/src" \
  "(use|extern crate)\\s+(clap::|pyo3::|env_logger::|quick[_-]xml::|zip::|flate2::|calamine::|serde_json::|actix|warp::|tokio::|axum::)" \
  "crates/docir-core/src"; then
  FAIL=1
fi

if ! check_file_patterns \
  "${REPO_ROOT}/crates/docir-app/src" \
  "(use|extern crate)\\s+(clap::|pyo3::|env_logger::|serde_json::|actix|warp::|tokio::|axum::|rocket::|rocket|anyhow::|clap::)" \
  "crates/docir-app/src"; then
  FAIL=1
fi

if ! check_domain_dependency_policy \
  "${REPO_ROOT}/crates/docir-core/Cargo.toml" \
  "docir-core"; then
  FAIL=1
fi

if ! check_domain_dependency_policy \
  "${REPO_ROOT}/crates/docir-app/Cargo.toml" \
  "docir-app"; then
  FAIL=1
fi

if [ "${FAIL}" -ne 0 ]; then
  echo
  echo "presentation boundary policy: FAIL"
  exit 1
fi

echo
echo "presentation boundary policy: PASS"
exit 0
