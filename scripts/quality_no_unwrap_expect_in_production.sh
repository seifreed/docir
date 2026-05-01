#!/usr/bin/env bash
set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "${REPO_ROOT}"

MODE="${1:-baseline}"
BASE_REF="${QUALITY_NO_UNWRAP_BASE:-origin/main}"
TARGET_REF="${2:-HEAD}"
ALLOW_LIST="${REPO_ROOT}/scripts/quality_no_unwrap_expect_in_production.allow"

if [ ! -f "${ALLOW_LIST}" ]; then
  : > "${ALLOW_LIST}"
fi

is_target_file() {
  case "$1" in
    crates/docir-parser/src/*|crates/docir-app/src/*|crates/docir-diff/src/*|crates/docir-security/src/*)
      ;;
    *)
      return 1
      ;;
  esac

  case "$1" in
    */tests/*|*/tests.rs|*"/tests_"*|*_tests.rs|*/test.rs)
      return 1
      ;;
    *)
      return 0
      ;;
  esac
}

collect_changed_files() {
  local scopes=(
    crates/docir-parser/src
    crates/docir-app/src
    crates/docir-diff/src
    crates/docir-security/src
  )

  case "$MODE" in
    baseline)
      git diff --name-only "${BASE_REF}...${TARGET_REF}" -- "${scopes[@]}"
      ;;
    working)
      git diff --name-only HEAD -- "${scopes[@]}"
      ;;
    inventory)
      rg --files "${scopes[@]}"
      ;;
    *)
      echo "Unknown mode: ${MODE}. Use 'baseline', 'working', or 'inventory'." >&2
      exit 2
      ;;
  esac
}

scan_file_for_violations() {
  local source="$1"
  local file="$2"

  if [ "$source" = "working" ]; then
    if [ ! -f "$file" ]; then
      return 0
    fi
    awk -v file="$file" -v allow_file="$ALLOW_LIST" '
      function load_allow_patterns() {
        while ((getline line < allow_file) > 0) {
          if (line ~ /^[[:space:]]*$/ || line ~ /^[[:space:]]*#/) {
            continue
          }
          allow_patterns[++allow_count] = line
        }
      }

      function brace_delta(text,    i, c, n) {
        n = 0
        for (i = 1; i <= length(text); i++) {
          c = substr(text, i, 1)
          if (c == "{") {
            n++
          } else if (c == "}") {
            n--
          }
        }
        return n
      }

      function is_allowed(line, i) {
        for (i = 1; i <= allow_count; i++) {
          if (line ~ allow_patterns[i]) {
            return 1
          }
        }
        return 0
      }

      function emit_violation(line, text) {
        if (is_allowed(text)) {
          return
        }
        printf "%s:%d:%s\n", file, line, text
      }

      BEGIN {
        load_allow_patterns()
        in_tests = 0
        pending_cfg = 0
        pending_open = 0
        test_depth = 0
      }

      {
        if ($0 ~ /^[[:space:]]*(pub[[:space:]]+)?mod[[:space:]]+tests[[:space:]]*\{/) {
          in_tests = 1
          pending_cfg = 0
          pending_open = 0
          test_depth = brace_delta($0)
          next
        }
        if ($0 ~ /^[[:space:]]*(pub[[:space:]]+)?mod[[:space:]]+tests[[:space:]]*$/) {
          in_tests = 1
          pending_cfg = 0
          pending_open = 1
          test_depth = 0
          next
        }

        if (pending_cfg) {
          if ($0 ~ /^[[:space:]]*mod[[:space:]]+tests[[:space:]]*\{/) {
            in_tests = 1
            pending_cfg = 0
            pending_open = 0
            test_depth = brace_delta($0)
            next
          }
          if ($0 ~ /^[[:space:]]*mod[[:space:]]+tests[[:space:]]*$/) {
            in_tests = 1
            pending_cfg = 0
            pending_open = 1
            test_depth = 0
            next
          }
          if ($0 !~ /^[[:space:]]*$/) {
            pending_cfg = 0
            pending_open = 0
          }
        }

        if (in_tests) {
          if (pending_open) {
            if ($0 !~ /\{/) {
              next
            }
            pending_open = 0
            if (brace_delta($0) <= 0) {
              in_tests = 0
              test_depth = 0
              next
            }
            test_depth = brace_delta($0)
            next
          }
          test_depth += brace_delta($0)
          if (test_depth <= 0) {
            in_tests = 0
            test_depth = 0
          }
          next
        }

        if ($0 ~ /^[[:space:]]*#\[cfg\(test\)\][[:space:]]*$/) {
          pending_cfg = 1
          next
        }

        if ($0 ~ /\b(panic|unreachable)!/ || $0 ~ /\.(unwrap|expect)[[:space:]]*\(/) {
          emit_violation(NR, $0)
        }
      }
    ' "$file"
    return 0
  fi

  if ! git cat-file -e "${source}:${file}" 2>/dev/null; then
    return 0
  fi
  git show "${source}:${file}" | scan_file_for_violations working "$file"
}

tmp_changed="$(mktemp)"
tmp_target="$(mktemp)"
tmp_base="$(mktemp)"
tmp_new="$(mktemp)"
tmp_resolved="$(mktemp)"
tmp_working="$(mktemp)"
tmp_new_agg="$(mktemp)"
tmp_resolved_agg="$(mktemp)"
tmp_working_agg="$(mktemp)"
tmp_inventory="$(mktemp)"
tmp_inventory_files="$(mktemp)"
tmp_inventory_modules="$(mktemp)"
trap 'rm -f "${tmp_changed}" "${tmp_target}" "${tmp_base}" "${tmp_new}" "${tmp_resolved}" "${tmp_working}" "${tmp_new_agg}" "${tmp_resolved_agg}" "${tmp_working_agg}" "${tmp_inventory}" "${tmp_inventory_files}" "${tmp_inventory_modules}"' EXIT

collect_changed_files > "${tmp_changed}" || true

if [ "${MODE}" = "inventory" ]; then
  inventory_total=0

  while IFS= read -r file; do
    is_target_file "$file" || continue
    > "${tmp_target}"
    scan_file_for_violations working "$file" > "${tmp_target}"
    count="$(wc -l < "${tmp_target}")"
    if [ "${count}" -le 0 ]; then
      continue
    fi

    printf '%s:%s\n' "${file}" "${count}" >> "${tmp_inventory_files}"
    inventory_total=$((inventory_total + count))

    module="${file%%/src/*}"
    if [ -z "${module}" ] || [ "${module}" = "${file}" ]; then
      module="${file%/*}"
    fi
    printf '%s:%s\n' "${module}" "${count}" >> "${tmp_inventory_modules}"
  done < "${tmp_changed}"

  {
    echo "Production unwrap/expect/panic/unreachable inventory (tests excluded):"
    echo
    if [ -s "${tmp_inventory_files}" ]; then
      echo "Top files:"
      sort -t: -k2,2nr "${tmp_inventory_files}"
      echo
      echo "By module:"
      awk -F: '{sum[$1]+=$2} END {for (k in sum) printf \"%s:%d\n\", k, sum[k]}' "${tmp_inventory_modules}" | sort -t: -k2,2nr
      echo
      echo "Summary: total=${inventory_total}"
    else
      echo "No production unwrap/expect/panic/unreachable usages detected."
      echo "Summary: total=0"
    fi
    echo "Approved exceptions file: ${ALLOW_LIST}"
  } > "${tmp_inventory}"

  cat "${tmp_inventory}"

  if [ -n "${QUALITY_NO_UNWRAP_INVENTORY_REPORT:-}" ]; then
    mkdir -p "$(dirname "${QUALITY_NO_UNWRAP_INVENTORY_REPORT}")"
    cp "${tmp_inventory}" "${QUALITY_NO_UNWRAP_INVENTORY_REPORT}"
    echo "Inventory report written to ${QUALITY_NO_UNWRAP_INVENTORY_REPORT}"
  fi

  if [ -n "${QUALITY_NO_UNWRAP_INVENTORY_MAX:-}" ]; then
    if ! [[ "${QUALITY_NO_UNWRAP_INVENTORY_MAX}" =~ ^[0-9]+$ ]]; then
      echo "Invalid QUALITY_NO_UNWRAP_INVENTORY_MAX='${QUALITY_NO_UNWRAP_INVENTORY_MAX}' (expected integer)." >&2
      exit 2
    fi
    if [ "${inventory_total}" -gt "${QUALITY_NO_UNWRAP_INVENTORY_MAX}" ]; then
      echo "Inventory total ${inventory_total} exceeds QUALITY_NO_UNWRAP_INVENTORY_MAX=${QUALITY_NO_UNWRAP_INVENTORY_MAX}" >&2
      exit 1
    fi
  fi

  exit 0
fi

if [ ! -s "${tmp_changed}" ]; then
  echo "No production files changed in ${MODE} scope."
  exit 0
fi

total_new=0
total_resolved=0
total_existing=0
working_total=0

while IFS= read -r file; do
  is_target_file "$file" || continue

  > "${tmp_target}"
  scan_file_for_violations working "$file" > "${tmp_target}"

  > "${tmp_base}"
  if [ "${MODE}" = "baseline" ]; then
    scan_file_for_violations "${BASE_REF}" "$file" > "${tmp_base}"
  else
    cat "${tmp_target}" > "${tmp_base}"
  fi

  if [ "${MODE}" = "baseline" ]; then
    comm -13 <(sort "${tmp_base}") <(sort "${tmp_target}") > "${tmp_new}"
    comm -23 <(sort "${tmp_base}") <(sort "${tmp_target}") > "${tmp_resolved}"
    comm -12 <(sort "${tmp_base}") <(sort "${tmp_target}") > "${tmp_working}"

    new_count="$(wc -l < "${tmp_new}")"
    resolved_count="$(wc -l < "${tmp_resolved}")"
    existing_count="$(wc -l < "${tmp_working}")"

    if [ "${new_count}" -gt 0 ]; then
      echo "File: ${file}" >> "${tmp_new_agg}"
      sed "s/^/  /" "${tmp_new}" >> "${tmp_new_agg}"
      echo >> "${tmp_new_agg}"
    fi

    if [ "${resolved_count}" -gt 0 ]; then
      echo "File: ${file}" >> "${tmp_resolved_agg}"
      sed "s/^/  /" "${tmp_resolved}" >> "${tmp_resolved_agg}"
      echo >> "${tmp_resolved_agg}"
    fi

    total_new=$((total_new + new_count))
    total_resolved=$((total_resolved + resolved_count))
    total_existing=$((total_existing + existing_count))
    continue
  fi

  working_count="$(wc -l < "${tmp_target}")"
  if [ "${working_count}" -gt 0 ]; then
    echo "File: ${file}" >> "${tmp_working_agg}"
    sed "s/^/  /" "${tmp_target}" >> "${tmp_working_agg}"
    echo >> "${tmp_working_agg}"
  fi
  working_total=$((working_total + working_count))
done < "${tmp_changed}"

if [ "${MODE}" = "baseline" ]; then
  if [ "${total_new}" -gt 0 ]; then
    echo "Production unwrap/expect/panic/unreachable additions detected:"
    cat "${tmp_new_agg}"
    echo
    echo "Resolved in this change:"
    if [ -s "${tmp_resolved_agg}" ]; then
      cat "${tmp_resolved_agg}"
    else
      echo "  none"
    fi
    echo
    echo "Summary: new=${total_new}, resolved=${total_resolved}, existing=${total_existing}"
    echo "Use Result-based error handling with ParserError/AppError/DiffError before merging."
    echo "Approved exceptions file: ${ALLOW_LIST}"
    exit 1
  fi

  echo "No new unwrap/expect/panic/unreachable production additions found."
  echo "Resolved in this change:"
  if [ -s "${tmp_resolved_agg}" ]; then
    cat "${tmp_resolved_agg}"
  else
    echo "  none"
  fi
  echo
  echo "Summary: new=${total_new}, resolved=${total_resolved}, existing=${total_existing}"
  echo "Approved exceptions file: ${ALLOW_LIST}"
  exit 0
fi

if [ "${MODE}" = "working" ]; then
  if [ "${working_total}" -gt 0 ]; then
    echo "Production unwrap/expect/panic/unreachable usages detected in working diff:"
    cat "${tmp_working_agg}"
    echo
    echo "Approved exceptions file: ${ALLOW_LIST}"
    exit 1
  fi
  echo "No unwrap/expect/panic/unreachable production usages in working diff."
  echo "Approved exceptions file: ${ALLOW_LIST}"
  exit 0
fi
