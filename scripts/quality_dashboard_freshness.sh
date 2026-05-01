#!/usr/bin/env bash
set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
LEDGER_FILE="${REPO_ROOT}/docs/quality-debt-dashboard.md"
BASELINE_DIR="${REPO_ROOT}/target/quality-baseline"

if [ ! -f "${LEDGER_FILE}" ]; then
  echo "quality_dashboard_freshness: FAIL (missing ${LEDGER_FILE})"
  exit 1
fi

latest_report="$(ls -1t "${BASELINE_DIR}"/quality-baseline-*.md 2>/dev/null | head -n 1 || true)"
if [ -z "${latest_report}" ]; then
  echo "quality_dashboard_freshness: FAIL (no baseline report in ${BASELINE_DIR})"
  exit 1
fi

last_dashboard_report="$(
  awk -F'|' '/\| .*quality-baseline-.*\.md/ {gsub(/^ +| +$/, "", $14); report=$14} END {print report}' "${LEDGER_FILE}"
)"

if [ -z "${last_dashboard_report}" ]; then
  echo "quality_dashboard_freshness: FAIL (no baseline path found in dashboard table)"
  exit 1
fi

if [ "${last_dashboard_report}" != "${latest_report}" ]; then
  echo "quality_dashboard_freshness: FAIL"
  echo "  dashboard: ${last_dashboard_report}"
  echo "  latest:    ${latest_report}"
  exit 1
fi

echo "quality_dashboard_freshness: PASS"
