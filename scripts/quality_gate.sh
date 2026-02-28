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

COVERAGE_THRESHOLD=95

print_help() {
  cat <<'USAGE'
Usage: ./scripts/quality_gate.sh [--help]

Canonical quality gate entrypoint for this repository.
USAGE
}

stage_validate_repo_root() {
  if [ "${QUALITY_GATE_FORCE_PRECONDITION_FAIL:-0}" = "1" ]; then
    gate_log "ERROR" "Forced precondition failure requested via QUALITY_GATE_FORCE_PRECONDITION_FAIL=1"
    return 2
  fi

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

stage_fmt_check() {
  if [ "${QUALITY_GATE_FORCE_FAIL:-0}" = "1" ]; then
    gate_log "ERROR" "Forced quality failure requested via QUALITY_GATE_FORCE_FAIL=1"
    return 1
  fi

  gate_run_command cargo fmt --all --check
}

stage_clippy_strict() {
  gate_run_command cargo clippy --all-targets --all-features -- -D warnings
}

stage_test_workspace() {
  gate_run_command cargo test
}

stage_coverage_check() {
  gate_run_command cargo llvm-cov --workspace --all-features --summary-only --fail-under-lines "${COVERAGE_THRESHOLD}"
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
    fmt_check)
      stage_fmt_check
      ;;
    clippy_strict)
      stage_clippy_strict
      ;;
    test_workspace)
      stage_test_workspace
      ;;
    coverage_check)
      stage_coverage_check
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

  for stage in validate_repo_root validate_tooling fmt_check clippy_strict test_workspace coverage_check; do
    set +e
    run_stage "${stage}" dispatch_stage "${stage}"
    stage_exit=$?
    set -e

    if [ "$stage_exit" -eq 0 ]; then
      continue
    fi

    gate_exit="$(classify_failure "${stage_exit}")"
    break
  done

  return "${gate_exit}"
}

emit_final_result() {
  local gate_exit="$1"
  local status
  local class

  case "$gate_exit" in
    0)
      status="PASS"
      class="pass"
      ;;
    1)
      status="FAIL"
      class="quality_failure"
      ;;
    2)
      status="FAIL"
      class="precondition_failure"
      ;;
    *)
      status="FAIL"
      class="quality_failure"
      gate_exit=1
      ;;
  esac

  printf 'QUALITY_GATE_RESULT=%s CLASS=%s EXIT_CODE=%s\n' "$status" "$class" "$gate_exit"
}

main() {
  if [ "${1:-}" = "--help" ] || [ "${1:-}" = "-h" ]; then
    print_help
    emit_final_result 0
    return 0
  fi

  local gate_exit=0
  if run_default_stages; then
    gate_exit=0
  else
    gate_exit=$?
  fi

  emit_final_result "$gate_exit"

  return "$gate_exit"
}

main "$@"
