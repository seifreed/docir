#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "${SCRIPT_DIR}/.." && pwd)"
LIB_PATH="${REPO_ROOT}/scripts/lib/quality_gate_lib.sh"
DEFAULT_COVERAGE_THRESHOLD_FILE="${REPO_ROOT}/scripts/quality_coverage_threshold.txt"

if [ ! -f "${LIB_PATH}" ]; then
  echo "Missing internal library: ${LIB_PATH}"
  echo "QUALITY_GATE_RESULT=FAIL"
  exit 2
fi

source "${LIB_PATH}"

QUALITY_GATE_DEFAULT_COVERAGE_THRESHOLD=88.23

resolve_coverage_threshold() {
  local source_path="${QUALITY_GATE_COVERAGE_THRESHOLD_FILE:-${DEFAULT_COVERAGE_THRESHOLD_FILE}}"

  if [ -n "${QUALITY_GATE_COVERAGE_THRESHOLD:-}" ]; then
    COVERAGE_THRESHOLD="${QUALITY_GATE_COVERAGE_THRESHOLD}"
    return 0
  fi

  if [ -f "${source_path}" ]; then
    COVERAGE_THRESHOLD="$(sed -e 's/[[:space:]]*#.*$//' "${source_path}" | tr -d ' \t\r\n')"
  fi

  if [ -z "${COVERAGE_THRESHOLD:-}" ]; then
    COVERAGE_THRESHOLD="${QUALITY_GATE_DEFAULT_COVERAGE_THRESHOLD}"
  fi

  if ! [[ "${COVERAGE_THRESHOLD}" =~ ^[0-9]+(\.[0-9]+)?$ ]]; then
    gate_log "ERROR" "Invalid coverage threshold value '${COVERAGE_THRESHOLD}'"
    return 2
  fi

  gate_log "INFO" "Using coverage threshold ${COVERAGE_THRESHOLD}% (source: ${source_path})"
}

print_help() {
  cat <<'USAGE'
Usage: ./scripts/quality_gate.sh [--help] [--with-layer-policy] [--only STAGE]

Canonical quality gate entrypoint for this repository.

Options:
  --with-layer-policy  Include the layer_policy stage (off by default)
  --with-api-hygiene   Include the api_hygiene stage (off by default)
  --only STAGE         Run only the specified stage
  --help               Show this help message
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

stage_no_unwrap_expect_in_production() {
  gate_run_command bash "${SCRIPT_DIR}/quality_no_unwrap_expect_in_production.sh" "${QUALITY_NO_UNWRAP_MODE:-working}"
}

stage_no_wildcard_super_in_production() {
  gate_run_command bash "${SCRIPT_DIR}/quality_no_wildcard_super_in_production.sh" "${QUALITY_NO_WILDCARD_MODE:-working}"
}

stage_layer_policy() {
  gate_run_command bash "${SCRIPT_DIR}/quality_layer_policy.sh"
}

stage_presentation_boundary_policy() {
  gate_run_command bash "${SCRIPT_DIR}/quality_presentation_boundary.sh"
}

stage_parser_pipeline_contracts() {
  gate_run_command bash "${SCRIPT_DIR}/quality_parser_pipeline_contracts.sh"
}

stage_crate_dependency_cycles() {
  gate_run_command bash "${SCRIPT_DIR}/quality_dependency_cycles.sh"
}

stage_api_hygiene() {
  gate_run_command bash "${SCRIPT_DIR}/quality_api_hygiene.sh"
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
  resolve_coverage_threshold
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
    no_unwrap_expect_in_production)
      stage_no_unwrap_expect_in_production
      ;;
    no_wildcard_super_in_production)
      stage_no_wildcard_super_in_production
      ;;
    layer_policy)
      stage_layer_policy
      ;;
    presentation_boundary_policy)
      stage_presentation_boundary_policy
      ;;
    parser_pipeline_contracts)
      stage_parser_pipeline_contracts
      ;;
    crate_dependency_cycles)
      stage_crate_dependency_cycles
      ;;
    api_hygiene)
      stage_api_hygiene
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

DEFAULT_STAGES=(
  validate_repo_root
  validate_tooling
  no_unwrap_expect_in_production
  no_wildcard_super_in_production
  presentation_boundary_policy
  parser_pipeline_contracts
  crate_dependency_cycles
  fmt_check
  clippy_strict
  test_workspace
  coverage_check
)

ON_DEMAND_STAGES=(
  layer_policy
  api_hygiene
)

run_default_stages() {
  local -a stages=()
  local stage
  local stage_exit=0
  local gate_exit=0

  if [ "${#RUN_ONLY_STAGES[@]}" -gt 0 ]; then
    stages=("${RUN_ONLY_STAGES[@]}")
  else
    stages=("${DEFAULT_STAGES[@]}")
    if [ "${INCLUDE_LAYER_POLICY:-0}" = "1" ]; then
      stages+=(layer_policy)
    fi
    if [ "${INCLUDE_API_HYGIENE:-0}" = "1" ]; then
      stages+=(api_hygiene)
    fi
  fi

  for stage in "${stages[@]}"; do
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

INCLUDE_LAYER_POLICY=0
INCLUDE_API_HYGIENE=0
RUN_ONLY_STAGES=()

main() {
  while [ "${#}" -gt 0 ]; do
    case "$1" in
      --help|-h)
        print_help
        emit_final_result 0
        return 0
        ;;
      --with-layer-policy)
        INCLUDE_LAYER_POLICY=1
        shift
        ;;
      --with-api-hygiene)
        INCLUDE_API_HYGIENE=1
        shift
        ;;
      --only)
        if [ -z "${2:-}" ]; then
          gate_log "ERROR" "--only requires a stage name"
          emit_final_result 2
          return 2
        fi
        RUN_ONLY_STAGES+=("$2")
        shift 2
        ;;
      *)
        gate_log "ERROR" "Unknown argument: $1"
        emit_final_result 2
        return 2
        ;;
    esac
  done

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
