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

: "${QUALITY_GATE_BASELINE_LOG:?QUALITY_GATE_BASELINE_LOG is required}"

subcmd="${1:-}"
shift || true

printf '%s %s\n' "${subcmd}" "$*" >> "${QUALITY_GATE_BASELINE_LOG}"

fail_stage="${QUALITY_GATE_BASELINE_FAIL_STAGE:-}"
if [ "${subcmd}" = "metadata" ]; then
  cat <<'JSON'
{"workspace_members":[],"packages":[],"resolve":{}}
JSON
  exit 0
fi

if [ "${subcmd}" = "cyclonedx" ] && [ "$*" = "--help" ]; then
  exit 0
fi

if [ -n "${fail_stage}" ] && [ "${subcmd}" = "${fail_stage}" ]; then
  exit 101
fi

if [ "${subcmd}" = "cyclonedx" ]; then
  mkdir -p crates/docir-core
  cat > crates/docir-core/docir-quality-sbom.json <<'JSON'
{
  "bomFormat": "CycloneDX",
  "specVersion": "1.5",
  "version": 1,
  "metadata": {
    "component": {
      "type": "application",
      "name": "docir-core",
      "version": "0.1.0",
      "purl": "pkg:cargo/docir-core@0.1.0"
    }
  },
  "components": [],
  "dependencies": []
}
JSON
  exit 0
fi

exit 0
SH
chmod +x "${fake_bin}/cargo"

cat >"${fake_bin}/sbom-tools" <<'SH'
#!/usr/bin/env bash
exit 0
SH
chmod +x "${fake_bin}/sbom-tools"

cat >"${fake_bin}/sbomqs" <<'SH'
#!/usr/bin/env bash
printf '10.0\tA\tNTIA Minimum Elements (2021)\t1.5\tjson\t%s\n' "${@: -1}"
SH
chmod +x "${fake_bin}/sbomqs"

run_case() {
  local name="$1"
  local fail_stage="$2"
  local expected_exit="$3"
  local expected_line_fragment="$4"
  shift 4
  local -a expected_calls=("$@")

  : > "${log_file}"
  local output_file
  output_file="$(mktemp)"

  set +e
  env \
    PATH="${fake_bin}:${PATH}" \
    QUALITY_GATE_BASELINE_LOG="${log_file}" \
    QUALITY_GATE_BASELINE_FAIL_STAGE="${fail_stage}" \
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
  "baseline-pass" \
  "" \
  0 \
  "QUALITY_GATE_RESULT=PASS CLASS=pass EXIT_CODE=0" \
  "metadata --format-version 1 --no-deps --offline" \
  "check --workspace --all-targets --all-features" \
  "deny check" \
  "audit " \
  "cyclonedx --help" \
  "cyclonedx --manifest-path Cargo.toml --format json --all-features --spec-version 1.5 --override-filename docir-quality-sbom" \
  "fmt --all --check" \
  "clippy --all-targets --all-features -- -D warnings" \
  "test " \
  "llvm-cov --workspace --all-features --summary-only --fail-under-lines ${coverage_threshold}"

run_case \
  "baseline-fail-fmt" \
  "fmt" \
  1 \
  "QUALITY_GATE_RESULT=FAIL CLASS=quality_failure EXIT_CODE=1" \
  "metadata --format-version 1 --no-deps --offline" \
  "check --workspace --all-targets --all-features" \
  "deny check" \
  "audit " \
  "cyclonedx --help" \
  "cyclonedx --manifest-path Cargo.toml --format json --all-features --spec-version 1.5 --override-filename docir-quality-sbom" \
  "fmt --all --check"

run_case \
  "baseline-fail-clippy" \
  "clippy" \
  1 \
  "QUALITY_GATE_RESULT=FAIL CLASS=quality_failure EXIT_CODE=1" \
  "metadata --format-version 1 --no-deps --offline" \
  "check --workspace --all-targets --all-features" \
  "deny check" \
  "audit " \
  "cyclonedx --help" \
  "cyclonedx --manifest-path Cargo.toml --format json --all-features --spec-version 1.5 --override-filename docir-quality-sbom" \
  "fmt --all --check" \
  "clippy --all-targets --all-features -- -D warnings"

run_case \
  "baseline-fail-test" \
  "test" \
  1 \
  "QUALITY_GATE_RESULT=FAIL CLASS=quality_failure EXIT_CODE=1" \
  "metadata --format-version 1 --no-deps --offline" \
  "check --workspace --all-targets --all-features" \
  "deny check" \
  "audit " \
  "cyclonedx --help" \
  "cyclonedx --manifest-path Cargo.toml --format json --all-features --spec-version 1.5 --override-filename docir-quality-sbom" \
  "fmt --all --check" \
  "clippy --all-targets --all-features -- -D warnings" \
  "test "

run_case \
  "baseline-fail-coverage" \
  "llvm-cov" \
  1 \
  "QUALITY_GATE_RESULT=FAIL CLASS=quality_failure EXIT_CODE=1" \
  "metadata --format-version 1 --no-deps --offline" \
  "check --workspace --all-targets --all-features" \
  "deny check" \
  "audit " \
  "cyclonedx --help" \
  "cyclonedx --manifest-path Cargo.toml --format json --all-features --spec-version 1.5 --override-filename docir-quality-sbom" \
  "fmt --all --check" \
  "clippy --all-targets --all-features -- -D warnings" \
  "test " \
  "llvm-cov --workspace --all-features --summary-only --fail-under-lines ${coverage_threshold}"

run_case \
  "baseline-fail-sbom" \
  "cyclonedx" \
  1 \
  "QUALITY_GATE_RESULT=FAIL CLASS=quality_failure EXIT_CODE=1" \
  "metadata --format-version 1 --no-deps --offline" \
  "check --workspace --all-targets --all-features" \
  "deny check" \
  "audit " \
  "cyclonedx --help" \
  "cyclonedx --manifest-path Cargo.toml --format json --all-features --spec-version 1.5 --override-filename docir-quality-sbom"

run_case \
  "baseline-fail-deny" \
  "deny" \
  1 \
  "QUALITY_GATE_RESULT=FAIL CLASS=quality_failure EXIT_CODE=1" \
  "metadata --format-version 1 --no-deps --offline" \
  "check --workspace --all-targets --all-features" \
  "deny check"

run_case \
  "baseline-fail-audit" \
  "audit" \
  1 \
  "QUALITY_GATE_RESULT=FAIL CLASS=quality_failure EXIT_CODE=1" \
  "metadata --format-version 1 --no-deps --offline" \
  "check --workspace --all-targets --all-features" \
  "deny check" \
  "audit "

echo "quality_gate_baseline_commands: OK"
