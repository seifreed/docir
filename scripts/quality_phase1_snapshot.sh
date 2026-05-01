#!/usr/bin/env bash
set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
TS="$(date -u +"%Y%m%dT%H%M%SZ")"
RUN_DIR="${REPO_ROOT}/target/quality-phase1/${TS}"
mkdir -p "${RUN_DIR}"

LEDGER_FILE="${REPO_ROOT}/docs/quality-debt-dashboard.md"
MIN_DUP_COUNT=3
MIN_DUP_LINES=12

FMT_CHECK_STATUS="NOT_RAN"
API_HYGIENE_STATUS="NOT_RAN"
LAYER_POLICY_STATUS="NOT_RAN"
DEPENDENCY_CYCLES_STATUS="NOT_RAN"
PIPELINE_CONTRACTS_STATUS="NOT_RAN"
PRESENTATION_BOUNDARY_STATUS="NOT_RAN"
BASELINE_STATUS="NOT_RAN"
DUPLICATE_PATTERNS_STATUS="NOT_RAN"
DASHBOARD_FRESHNESS_STATUS="NOT_RAN"

set_stage_status() {
  local name="$1"
  local status="$2"
  case "${name}" in
    fmt_check) FMT_CHECK_STATUS="${status}" ;;
    api_hygiene) API_HYGIENE_STATUS="${status}" ;;
    layer_policy) LAYER_POLICY_STATUS="${status}" ;;
    dependency_cycles) DEPENDENCY_CYCLES_STATUS="${status}" ;;
    pipeline_contracts) PIPELINE_CONTRACTS_STATUS="${status}" ;;
    presentation_boundary) PRESENTATION_BOUNDARY_STATUS="${status}" ;;
    phase1_baseline) BASELINE_STATUS="${status}" ;;
    duplicate_patterns) DUPLICATE_PATTERNS_STATUS="${status}" ;;
    dashboard_freshness) DASHBOARD_FRESHNESS_STATUS="${status}" ;;
    *) ;;
  esac
}

run_stage() {
  local name="$1"
  local log_file="$2"
  shift 2

  set +e
  "$@" >"${log_file}" 2>&1
  local status=$?
  set -e

  if [ "${status}" -eq 0 ]; then
    set_stage_status "${name}" "PASS"
  else
    set_stage_status "${name}" "FAIL"
  fi
}

cd "${REPO_ROOT}"

run_stage "fmt_check" "${RUN_DIR}/fmt_check.log" cargo fmt --all --check
run_stage "api_hygiene" "${RUN_DIR}/quality_api_hygiene.log" bash scripts/quality_api_hygiene.sh
run_stage "layer_policy" "${RUN_DIR}/quality_layer_policy.log" bash scripts/quality_layer_policy.sh
run_stage "dependency_cycles" "${RUN_DIR}/quality_dependency_cycles.log" bash scripts/quality_dependency_cycles.sh
run_stage "pipeline_contracts" "${RUN_DIR}/quality_parser_pipeline_contracts.log" bash scripts/quality_parser_pipeline_contracts.sh
run_stage "presentation_boundary" "${RUN_DIR}/quality_presentation_boundary.log" bash scripts/quality_presentation_boundary.sh
run_stage "phase1_baseline" "${RUN_DIR}/quality_phase1_baseline.log" bash scripts/quality_phase1_baseline.sh
run_stage "duplicate_patterns" "${RUN_DIR}/quality_duplicate_patterns.log" bash scripts/quality_duplicate_patterns.sh "${MIN_DUP_COUNT}" "${MIN_DUP_LINES}"
cp "${RUN_DIR}/quality_duplicate_patterns.log" "${RUN_DIR}/quality_duplicate_patterns.md"

BASELINE_REPORT=""
if [ -f "${RUN_DIR}/quality_phase1_baseline.log" ]; then
  BASELINE_REPORT="$(awk '/^Output:/ {print $2}' "${RUN_DIR}/quality_phase1_baseline.log" | tail -n 1)"
fi

if [ -z "${BASELINE_REPORT}" ]; then
  BASELINE_STATUS="FAIL"
  BASELINE_REPORT="${REPO_ROOT}/target/quality-baseline/quality-baseline-${TS}.md"
fi

DUP_PATH="${RUN_DIR}/quality_duplicate_patterns.md"
if [ -f "${DUP_PATH}" ]; then
  DUP_SUMMARY="$(python3 - "${DUP_PATH}" <<'PY'
import re
import sys

text = open(sys.argv[1], encoding="utf-8").read()
match = re.search(r"Found (\\d+) duplicate", text)
print(match.group(1) if match else "0")
PY
)"
else
  DUP_SUMMARY="0"
fi

if [ -f "${BASELINE_REPORT}" ]; then
  python3 - "${BASELINE_REPORT}" > "${RUN_DIR}/phase1_metrics.tsv" <<'PY'
import re
import sys
from pathlib import Path

report = Path(sys.argv[1]).read_text(encoding="utf-8")

def extract(label):
    match = re.search(r"\|\s*" + re.escape(label) + r"\s*\|\s*([^|]+)\|", report)
    return match.group(1).strip() if match else "n/a"

items = [
    ("rust_files", "Rust files (src)"),
    ("prod_files", "Production files (src)"),
    ("total_loc", "Total LOC"),
    ("files_over_800", "Files over 800 LOC"),
    ("functions_over_100", "Functions over 100 LOC (heuristic)"),
    ("panic_like", "Panic/unwrap/expect calls in production"),
    ("arch_viol", "Architecture dependency violations"),
]

for key, label in items:
    print(f"{key}={extract(label)}")
PY
fi

RUST_FILES="n/a"
PROD_FILES="n/a"
TOTAL_LOC="n/a"
FILES_OVER_800="n/a"
FN_OVER_100="n/a"
PANIC_LIKE="n/a"
ARCH_VIOL="n/a"

if [ -f "${RUN_DIR}/phase1_metrics.tsv" ]; then
  while IFS='=' read -r key value; do
    case "${key}" in
      rust_files) RUST_FILES="${value}" ;;
      prod_files) PROD_FILES="${value}" ;;
      total_loc) TOTAL_LOC="${value}" ;;
      files_over_800) FILES_OVER_800="${value}" ;;
      functions_over_100) FN_OVER_100="${value}" ;;
      panic_like) PANIC_LIKE="${value}" ;;
      arch_viol) ARCH_VIOL="${value}" ;;
      *) ;;
    esac
  done < "${RUN_DIR}/phase1_metrics.tsv"
fi

cat > "${RUN_DIR}/quality_phase1_snapshot.md" <<EOF
# Fase 1 — Snapshot semanal

Generated: ${TS}
Commit: $(git rev-parse --short HEAD)
Baseline report: ${BASELINE_REPORT}

## Estado de checks obligatorios

| Check | Status | Log |
|---|---|---|
| cargo fmt --check | ${FMT_CHECK_STATUS} | ${RUN_DIR}/fmt_check.log |
| quality_api_hygiene | ${API_HYGIENE_STATUS} | ${RUN_DIR}/quality_api_hygiene.log |
| quality_layer_policy | ${LAYER_POLICY_STATUS} | ${RUN_DIR}/quality_layer_policy.log |
| quality_dependency_cycles | ${DEPENDENCY_CYCLES_STATUS} | ${RUN_DIR}/quality_dependency_cycles.log |
| quality_parser_pipeline_contracts | ${PIPELINE_CONTRACTS_STATUS} | ${RUN_DIR}/quality_parser_pipeline_contracts.log |
| quality_presentation_boundary | ${PRESENTATION_BOUNDARY_STATUS} | ${RUN_DIR}/quality_presentation_boundary.log |
| quality_phase1_baseline | ${BASELINE_STATUS} | ${RUN_DIR}/quality_phase1_baseline.log |
| quality_duplicate_patterns | ${DUPLICATE_PATTERNS_STATUS} | ${RUN_DIR}/quality_duplicate_patterns.log |
| quality_dashboard_freshness | ${DASHBOARD_FRESHNESS_STATUS} | ${RUN_DIR}/quality_dashboard_freshness.log |

## Métricas de deuda base

| Métrica | Valor |
|---|---:|
| Rust files (src) | ${RUST_FILES} |
| Production files (src) | ${PROD_FILES} |
| Total LOC | ${TOTAL_LOC} |
| Files > 800 LOC | ${FILES_OVER_800} |
| Functions > 100 LOC | ${FN_OVER_100} |
| unwrap/expect/panic in production | ${PANIC_LIKE} |
| Duplicados funcionales (>= ${MIN_DUP_COUNT} apariciones, >= ${MIN_DUP_LINES} LOC) | ${DUP_SUMMARY} |
| Violaciones arquitectónicas de dependencia | ${ARCH_VIOL} |

## Artículos de deuda residual

- Snapshot reproducible en:
  - target/quality-phase1/${TS}/quality_api_hygiene.log
  - target/quality-phase1/${TS}/quality_layer_policy.log
  - target/quality-phase1/${TS}/quality_dependency_cycles.log
  - target/quality-phase1/${TS}/quality_parser_pipeline_contracts.log
  - target/quality-phase1/${TS}/quality_presentation_boundary.log
  - target/quality-phase1/${TS}/quality_phase1_baseline.log
  - target/quality-phase1/${TS}/quality_duplicate_patterns.log
  - target/quality-phase1/${TS}/quality_duplicate_patterns.md
  - target/quality-phase1/${TS}/quality_dashboard_freshness.log

EOF

if [ ! -f "${LEDGER_FILE}" ]; then
  cat > "${LEDGER_FILE}" <<'EOF'
# Dashboard de deuda técnica de Fase 1

| Semana | Commit | fmt_check | api_hygiene | layer_policy | dependency_cycles | parser_pipeline_contracts | presentation_boundary | Files > 800 LOC | Funciones > 100 LOC | unwrap/expect/panic | Duplicados (grupos) | Baseline report |
|---|---|---|---|---|---|---|---|---|---|---|---|
EOF
fi

WEEK="$(date -u +"%G-%V")"
COMMIT="$(git rev-parse --short HEAD)"
printf "| %s | %s | %s | %s | %s | %s | %s | %s | %s | %s | %s | %s | %s |\n" \
  "${WEEK}" \
  "${COMMIT}" \
  "${FMT_CHECK_STATUS}" \
  "${API_HYGIENE_STATUS}" \
  "${LAYER_POLICY_STATUS}" \
  "${DEPENDENCY_CYCLES_STATUS}" \
  "${PIPELINE_CONTRACTS_STATUS}" \
  "${PRESENTATION_BOUNDARY_STATUS}" \
  "${FILES_OVER_800}" \
  "${FN_OVER_100}" \
  "${PANIC_LIKE}" \
  "${DUP_SUMMARY}" \
  "${BASELINE_REPORT}" \
  >> "${LEDGER_FILE}"

run_stage "dashboard_freshness" "${RUN_DIR}/quality_dashboard_freshness.log" bash scripts/quality_dashboard_freshness.sh

if [ "${FMT_CHECK_STATUS}" != "PASS" ] || \
   [ "${API_HYGIENE_STATUS}" != "PASS" ] || \
   [ "${LAYER_POLICY_STATUS}" != "PASS" ] || \
   [ "${DEPENDENCY_CYCLES_STATUS}" != "PASS" ] || \
   [ "${PIPELINE_CONTRACTS_STATUS}" != "PASS" ] || \
   [ "${PRESENTATION_BOUNDARY_STATUS}" != "PASS" ] || \
   [ "${BASELINE_STATUS}" != "PASS" ] || \
   [ "${DUPLICATE_PATTERNS_STATUS}" != "PASS" ] || \
   [ "${DASHBOARD_FRESHNESS_STATUS}" != "PASS" ]; then
  echo "Fase 1 snapshot: FAIL"
  echo "See ${RUN_DIR}/quality_phase1_snapshot.md for details."
  echo "LEDGER updated: ${LEDGER_FILE}"
  exit 1
fi

echo "Fase 1 snapshot: PASS"
echo "Snapshot: ${RUN_DIR}/quality_phase1_snapshot.md"
echo "LEDGER: ${LEDGER_FILE}"
echo "Duplicate pattern report: ${RUN_DIR}/quality_duplicate_patterns.md"
