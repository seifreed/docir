gate_log() {
  local level="$1"
  shift
  printf '[quality-gate][%s] %s\n' "$level" "$*"
}

gate_require_tool() {
  local tool="$1"
  if command -v "$tool" >/dev/null 2>&1; then
    return 0
  fi

  gate_log "ERROR" "Missing required tool: ${tool}"
  return 2
}

run_stage() {
  local stage="$1"
  shift

  gate_log "INFO" "STAGE_START ${stage}"
  set +e
  "$@"
  local exit_code=$?
  set -e

  if [ "$exit_code" -eq 0 ]; then
    gate_log "INFO" "STAGE_PASS ${stage}"
  else
    gate_log "ERROR" "STAGE_FAIL ${stage} exit_code=${exit_code}"
  fi

  return "$exit_code"
}

classify_failure() {
  local exit_code="$1"
  case "$exit_code" in
    2)
      echo 2
      ;;
    *)
      echo 1
      ;;
  esac
}

emit_result() {
  local status="$1"
  printf 'QUALITY_GATE_RESULT=%s\n' "$status"
}

if [ "${BASH_SOURCE[0]}" = "$0" ]; then
  echo "This file is an internal library and must be sourced from scripts/quality_gate.sh"
  exit 2
fi
