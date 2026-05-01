#!/usr/bin/env bash
set -euo pipefail

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
cd "${repo_root}"

if [ ! -x "./scripts/quality_gate.sh" ]; then
  echo "Canonical gate is missing or not executable: ./scripts/quality_gate.sh"
  exit 1
fi

fake_bin="$(mktemp -d)"
log_file="$(mktemp)"
trap 'rm -rf "${fake_bin}"; rm -f "${log_file}"' EXIT
coverage_threshold="$(sed -e 's/[[:space:]]*#.*$//' "${repo_root}/scripts/quality_coverage_threshold.txt" | tr -d ' \t\r\n')"
coverage_threshold="${coverage_threshold:-88.23}"

cat >"${fake_bin}/cargo" <<'SH'
#!/usr/bin/env bash
set -euo pipefail

: "${QUALITY_GATE_COVERAGE_LOG:?QUALITY_GATE_COVERAGE_LOG is required}"

subcmd="${1:-}"
shift || true

printf '%s %s\n' "${subcmd}" "$*" >> "${QUALITY_GATE_COVERAGE_LOG}"

if [ "${subcmd}" = "llvm-cov" ] && [ "${QUALITY_GATE_COVERAGE_FAIL:-0}" = "1" ]; then
  exit 101
fi

if [ "${subcmd}" = "metadata" ]; then
  cat <<'JSON'
{"workspace_members":[],"packages":[],"resolve":{}}
JSON
  exit 0
fi

exit 0
SH
chmod +x "${fake_bin}/cargo"

run_case() {
  local name="$1"
  local expected_exit="$2"
  local expected_line_fragment="$3"
  shift 3
  local -a expected_calls=("$@")

  : > "${log_file}"
  local output_file
  output_file="$(mktemp)"

  set +e
  env \
    PATH="${fake_bin}:${PATH}" \
    QUALITY_GATE_COVERAGE_LOG="${log_file}" \
    ./scripts/quality_gate.sh >"${output_file}" 2>&1
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

  actual_calls=()
  while IFS= read -r line; do
    actual_calls+=("${line}")
  done < "${log_file}"

  if [ "${#actual_calls[@]}" -ne "${#expected_calls[@]}" ]; then
    echo "${name}: expected ${#expected_calls[@]} cargo calls, got ${#actual_calls[@]}"
    printf 'Expected:\n%s\n' "${expected_calls[*]}"
    printf 'Actual:\n%s\n' "${actual_calls[*]}"
    rm -f "${output_file}"
    exit 1
  fi

  local idx
  for idx in "${!expected_calls[@]}"; do
    if [ "${actual_calls[$idx]}" != "${expected_calls[$idx]}" ]; then
      echo "${name}: call $((idx + 1)) mismatch"
      echo "Expected: ${expected_calls[$idx]}"
      echo "Actual:   ${actual_calls[$idx]}"
      rm -f "${output_file}"
      exit 1
    fi
  done

  rm -f "${output_file}"
  echo "${name}: OK"
}

run_case \
  "coverage-command-contract" \
  0 \
  "QUALITY_GATE_RESULT=PASS CLASS=pass EXIT_CODE=0" \
  "metadata --format-version 1 --no-deps --offline" \
  "check --workspace --all-targets --all-features" \
  "fmt --all --check" \
  "clippy --all-targets --all-features -- -D warnings" \
  "test " \
  "llvm-cov --workspace --all-features --summary-only --fail-under-lines ${coverage_threshold}"

set +e
output_file="$(mktemp)"
: > "${log_file}"
env \
  PATH="${fake_bin}:${PATH}" \
  QUALITY_GATE_COVERAGE_LOG="${log_file}" \
  QUALITY_GATE_COVERAGE_FAIL=1 \
  ./scripts/quality_gate.sh >"${output_file}" 2>&1
actual_exit=$?
set -e

result_line="$(tail -n 1 "${output_file}")"
if [ "${actual_exit}" -ne 1 ]; then
  echo "coverage-threshold-fail: expected exit 1, got ${actual_exit}"
  cat "${output_file}"
  rm -f "${output_file}"
  exit 1
fi

if [[ "${result_line}" != *"QUALITY_GATE_RESULT=FAIL CLASS=quality_failure EXIT_CODE=1"* ]]; then
  echo "coverage-threshold-fail: final status line mismatch"
  echo "Actual line: ${result_line}"
  cat "${output_file}"
  rm -f "${output_file}"
  exit 1
fi

if ! rg -q '^metadata --format-version 1 --no-deps --offline$' "${log_file}" \
  || ! rg -q '^check --workspace --all-targets --all-features$' "${log_file}" \
  || ! rg -q "^llvm-cov --workspace --all-features --summary-only --fail-under-lines ${coverage_threshold}\$" "${log_file}"; then
  echo "coverage-threshold-fail: missing expected llvm-cov command invocation"
  cat "${log_file}"
  rm -f "${output_file}"
  exit 1
fi

rm -f "${output_file}"
echo "coverage-threshold-fail: OK"
echo "quality_gate_coverage_commands: OK"
