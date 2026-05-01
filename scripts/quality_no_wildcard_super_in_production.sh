#!/usr/bin/env bash
set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "${REPO_ROOT}"

MODE="${1:-working}"
BASE_REF="${QUALITY_NO_WILDCARD_BASE:-origin/main}"
TARGET_REF="${2:-HEAD}"
INVENTORY_FAIL="${QUALITY_NO_WILDCARD_INVENTORY_FAIL:-0}"
SCOPES=(
  crates/docir-parser/src
  crates/docir-core/src
  crates/docir-app/src
  crates/docir-diff/src
  crates/docir-security/src
  crates/docir-rules/src
)

is_target_file() {
  case "$1" in
    crates/docir-parser/src/*|crates/docir-core/src/*|crates/docir-app/src/*|crates/docir-diff/src/*|crates/docir-security/src/*|crates/docir-rules/src/*)
      ;;
    *)
      return 1
      ;;
  esac

  case "$1" in
    */tests/*|*/test/*|*tests.rs|*_tests.rs|*/tests_*|*/test_*)
      return 1
      ;;
    *)
      return 0
      ;;
  esac
}

scan_diff_added_violations() {
  local diff_cmd=()
  case "$MODE" in
    working)
      diff_cmd=(git diff -U0 HEAD -- "${SCOPES[@]}")
      ;;
    baseline)
      diff_cmd=(git diff -U0 "${BASE_REF}...${TARGET_REF}" -- "${SCOPES[@]}")
      ;;
    *)
      echo "Unknown mode: ${MODE}. Use 'working', 'baseline', or 'inventory'." >&2
      exit 2
      ;;
  esac

  "${diff_cmd[@]}" | awk '
    function is_target_file(file) {
      if (file !~ /^crates\/(docir-parser|docir-core|docir-app|docir-diff|docir-security|docir-rules)\/src\//) return 0
      if (file ~ /\/tests\//) return 0
      if (file ~ /\/test\//) return 0
      if (file ~ /tests\.rs$/) return 0
      if (file ~ /_tests\.rs$/) return 0
      if (file ~ /\/tests_/) return 0
      if (file ~ /\/test_/) return 0
      return 1
    }

    /^diff --git / {
      file = ""
      next
    }

    /^\+\+\+ b\// {
      file = substr($0, 7)
      in_target = is_target_file(file)
      next
    }

    /^@@ / {
      if (!in_target) next
      line = $0
      sub(/^@@ -[0-9]+(,[0-9]+)? \+/, "", line)
      sub(/ .*/, "", line)
      split(line, parts, ",")
      new_line = parts[1] + 0
      next
    }

    {
      if (!in_target) next
      if ($0 ~ /^\+/ && $0 !~ /^\+\+\+/) {
        code = substr($0, 2)
        if (code ~ /^use[[:space:]]+super::\*[[:space:]]*;/) {
          printf "%s:%d:%s\n", file, new_line, code
        }
        new_line++
        next
      }
      if ($0 !~ /^-/) {
        new_line++
      }
    }
  '
}

scan_file() {
  local file="$1"
  [ -f "$file" ] || return 0

  awk -v file="$file" '
    function brace_delta(text,    i, c, n) {
      n = 0
      for (i = 1; i <= length(text); i++) {
        c = substr(text, i, 1)
        if (c == "{") n++
        else if (c == "}") n--
      }
      return n
    }

    BEGIN {
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

      if ($0 ~ /^use[[:space:]]+super::\*[[:space:]]*;/) {
        printf "%s:%d:%s\n", file, NR, $0
      }
    }
  ' "$file"
}

tmp_violations="$(mktemp)"
trap 'rm -f "${tmp_violations}"' EXIT

if [ "${MODE}" = "inventory" ]; then
  while IFS= read -r file; do
    is_target_file "$file" || continue
    scan_file "$file" >> "${tmp_violations}"
  done < <(rg --files "${SCOPES[@]}")
else
  scan_diff_added_violations > "${tmp_violations}"
fi

count="$(wc -l < "${tmp_violations}" | tr -d '[:space:]')"

if [ "${MODE}" = "inventory" ]; then
  echo "Wildcard imports inventory (use super::*) in production scope:"
  if [ "${count}" -gt 0 ]; then
    cat "${tmp_violations}"
  else
    echo "No production wildcard imports detected."
  fi
  echo "Summary: total=${count}"
  if [ "${INVENTORY_FAIL}" = "1" ] && [ "${count}" -gt 0 ]; then
    echo "quality_no_wildcard_super_in_production: FAIL (inventory strict mode)"
    exit 1
  fi
  exit 0
fi

if [ "${count}" -gt 0 ]; then
  echo "quality_no_wildcard_super_in_production: FAIL"
  cat "${tmp_violations}"
  exit 1
fi

echo "quality_no_wildcard_super_in_production: PASS"
