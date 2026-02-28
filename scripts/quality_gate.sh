#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "${SCRIPT_DIR}/.." && pwd)"
LIB_PATH="${REPO_ROOT}/scripts/lib/quality_gate_lib.sh"

if [ ! -f "${LIB_PATH}" ]; then
  echo "Missing internal library: ${LIB_PATH}"
  echo "QUALITY_GATE_RESULT=FAIL"
  exit 2
fi

# shellcheck source=/dev/null
source "${LIB_PATH}"

print_help() {
  cat <<'USAGE'
Usage: ./scripts/quality_gate.sh [--help]

Canonical quality gate entrypoint for this repository.
USAGE
}

stage_validate_repo_root() {
  local cwd
  cwd="$(pwd)"
  if [ "${cwd}" != "${REPO_ROOT}" ]; then
    gate_log "ERROR" "Run from repository root: ${REPO_ROOT}"
    return 2
  fi

  return 0
}

stage_validate_tooling() {
  gate_require_tool cargo
}

stage_contract_scaffold() {
  return 0
}

dispatch_stage() {
  local stage="$1"
  case "$stage" in
    validate_repo_root)
      stage_validate_repo_root
      ;;
    validate_tooling)
      stage_validate_tooling
      ;;
    contract_scaffold)
      stage_contract_scaffold
      ;;
    *)
      gate_log "ERROR" "Unknown stage: ${stage}"
      return 2
      ;;
  esac
}

run_default_stages() {
  local stage
  local stage_exit=0
  local gate_exit=0

  for stage in validate_repo_root validate_tooling contract_scaffold; do
    if run_stage "${stage}" dispatch_stage "${stage}"; then
      continue
    fi

    stage_exit=$?
    gate_exit="$(classify_failure "${stage_exit}")"
    break
  done

  return "${gate_exit}"
}

main() {
  if [ "${1:-}" = "--help" ] || [ "${1:-}" = "-h" ]; then
    print_help
    emit_result "PASS"
    return 0
  fi

  local gate_exit=0
  if run_default_stages; then
    gate_exit=0
  else
    gate_exit=$?
  fi

  if [ "$gate_exit" -eq 0 ]; then
    emit_result "PASS"
  else
    emit_result "FAIL"
  fi

  return "$gate_exit"
}

main "$@"
