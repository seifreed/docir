#!/usr/bin/env bash
set -euo pipefail

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
cd "${repo_root}"

if [ ! -x "./scripts/quality_gate.sh" ]; then
  echo "Canonical gate is missing or not executable: ./scripts/quality_gate.sh"
  exit 1
fi

fake_bin="$(mktemp -d)"
cat >"${fake_bin}/cargo" <<'SH'
#!/usr/bin/env bash
set -euo pipefail
exit 0
SH
chmod +x "${fake_bin}/cargo"
trap 'rm -rf "${fake_bin}"' EXIT

run_case() {
  local name="$1"
  local expected_exit="$2"
  local expected_line_fragment="$3"
  shift 3

  local output_file
  output_file="$(mktemp)"

  set +e
  env PATH="${fake_bin}:${PATH}" "$@" >"${output_file}" 2>&1
  local actual_exit=$?
  set -e

  local result_line
  result_line="$(tail -n 1 "${output_file}")"

  if [ "${actual_exit}" -ne "${expected_exit}" ]; then
    echo "${name}: expected exit ${expected_exit}, got ${actual_exit}"
    cat "${output_file}"
    rm -f "${output_file}"
    exit 1
  fi

  if [[ "${result_line}" != QUALITY_GATE_RESULT=* ]]; then
    echo "${name}: missing final QUALITY_GATE_RESULT line"
    cat "${output_file}"
    rm -f "${output_file}"
    exit 1
  fi

  if [[ "${result_line}" != *"${expected_line_fragment}"* ]]; then
    echo "${name}: final status line mismatch"
    echo "Expected fragment: ${expected_line_fragment}"
    echo "Actual line: ${result_line}"
    cat "${output_file}"
    rm -f "${output_file}"
    exit 1
  fi

  rm -f "${output_file}"
  echo "${name}: OK"
}

run_case \
  "pass" \
  0 \
  "QUALITY_GATE_RESULT=PASS CLASS=pass EXIT_CODE=0" \
  ./scripts/quality_gate.sh

run_case \
  "quality-fail" \
  1 \
  "QUALITY_GATE_RESULT=FAIL CLASS=quality_failure EXIT_CODE=1" \
  env QUALITY_GATE_FORCE_FAIL=1 ./scripts/quality_gate.sh

run_case \
  "precondition-fail" \
  2 \
  "QUALITY_GATE_RESULT=FAIL CLASS=precondition_failure EXIT_CODE=2" \
  env QUALITY_GATE_FORCE_PRECONDITION_FAIL=1 ./scripts/quality_gate.sh

bash scripts/tests/quality_gate_baseline_commands.sh

echo "quality_gate_exit_codes: OK"
