#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "${SCRIPT_DIR}/.." && pwd)"
OUT_FILE="${1:-${REPO_ROOT}/reports/quality/residual_dashboard.md}"
LARGE_FILE_THRESHOLD="${LARGE_FILE_THRESHOLD:-650}"

mkdir -p "$(dirname "${OUT_FILE}")"

tmp_inventory="$(mktemp)"
tmp_large="$(mktemp)"
tmp_dups="$(mktemp)"
trap 'rm -f "${tmp_inventory}" "${tmp_large}" "${tmp_dups}"' EXIT

(
  cd "${REPO_ROOT}"
  QUALITY_NO_UNWRAP_INVENTORY_REPORT="${tmp_inventory}" \
    bash "${SCRIPT_DIR}/quality_no_unwrap_expect_in_production.sh" inventory >/dev/null
)

(
  cd "${REPO_ROOT}"
  while IFS= read -r file; do
    case "${file}" in
      */tests/*|*/test/*|*tests.rs|*_tests.rs|*/tests_*|*/test_*)
        continue
        ;;
    esac
    loc="$(wc -l < "${file}" | tr -d ' ')"
    if [ "${loc}" -ge "${LARGE_FILE_THRESHOLD}" ]; then
      printf "%s:%s\n" "${file}" "${loc}"
    fi
  done < <(rg --files crates/*/src)
) | sort -t: -k2,2nr > "${tmp_large}"

(
  cd "${REPO_ROOT}"
  bash "${SCRIPT_DIR}/quality_duplicate_patterns.sh" 3 12
) > "${tmp_dups}" || true

{
  echo "# Quality Residual Dashboard"
  echo
  echo "- Generated: $(date -u +"%Y-%m-%dT%H:%M:%SZ")"
  echo "- Large file threshold: ${LARGE_FILE_THRESHOLD} LOC"
  echo
  echo "## 1) panic/unwrap/expect/unreachable (production inventory)"
  echo
  cat "${tmp_inventory}"
  echo
  echo "## 2) Large production files"
  echo
  if [ -s "${tmp_large}" ]; then
    echo "| File | LOC |"
    echo "|---|---:|"
    while IFS=: read -r file loc; do
      echo "| \`${file}\` | ${loc} |"
    done < "${tmp_large}"
  else
    echo "None"
  fi
  echo
  echo "## 3) Duplicate pattern groups"
  echo
  cat "${tmp_dups}"
} > "${OUT_FILE}"

echo "Residual dashboard written to ${OUT_FILE}"
