#!/usr/bin/env bash
set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
OUT_DIR="${REPO_ROOT}/target/quality-baseline"
mkdir -p "${OUT_DIR}"
source "${REPO_ROOT}/scripts/lib/dependency_allowlist.sh"

MAX_FILE_LOC=800
MAX_FN_LOC=100
MODE_FAIL=0

for arg in "$@"; do
  case "${arg}" in
    --fail-on-violations)
      MODE_FAIL=1
      ;;
    -h|--help)
      cat <<'USAGE'
Usage: ./scripts/quality_phase1_baseline.sh [--fail-on-violations]

Collect quality baseline metrics for clean code / clean architecture phase 1:
- LOC per file
- Files over 800 LOC
- Production functions over 100 LOC (heuristic)
- unwrap/expect/panic/unreachable usage in production files
- Quick architecture sanity checks (dependency direction and infra leakage in core)
USAGE
      exit 0
      ;;
    *)
      echo "Unknown argument: ${arg}" >&2
      echo "Use --help for usage." >&2
      exit 2
      ;;
  esac
done

TS="$(date -u +"%Y%m%dT%H%M%SZ")"
REPORT="${OUT_DIR}/quality-baseline-${TS}.md"
TMP_DIR="$(mktemp -d)"
trap 'rm -rf "${TMP_DIR}"' EXIT

CRATE_SUMMARY="${TMP_DIR}/crate_summary.tsv"
FILES_OVER="${TMP_DIR}/files_over.tsv"
FN_OVER="${TMP_DIR}/functions_over.tsv"
PANIC_FILES="${TMP_DIR}/panic_files.tsv"
ARCH_DEP_VIOL="${TMP_DIR}/arch_dep_violations.tsv"
ARCH_INFRA="${TMP_DIR}/arch_infra_core.txt"

touch "${CRATE_SUMMARY}" "${FILES_OVER}" "${FN_OVER}" "${PANIC_FILES}" "${ARCH_DEP_VIOL}" "${ARCH_INFRA}"

TOTAL_FILES=0
TOTAL_LOC=0
TOTAL_PROD_FILES=0
TOTAL_FILES_OVER=0
TOTAL_FN_OVER=0
TOTAL_PANIC=0
TOTAL_ARCH_DEP_VIOL=0

is_production_file() {
  case "$1" in
    */crates/docir-parser/src/*|*/crates/docir-app/src/*|*/crates/docir-diff/src/*|*/crates/docir-security/src/*)
      ;;
    *)
      return 1
      ;;
  esac

  case "$1" in
    *"/tests/"*|*"\\tests\\"*|*/target/*|*/tests.rs|*/test.rs|*_tests.rs|*_test.rs|*/tests_*|*/test_*)
      return 1
      ;;
    *)
      return 0
      ;;
  esac
}

scan_panic_like_count() {
  local file="$1"
  awk '
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
      violations = 0
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
        violations++
      }
    }

    END {
      print violations
    }
  ' "$file"
}

scan_long_functions() {
  local file="$1"
  local awk_script
  awk_script='
function count_braces(line) {
  return gsub(/\{/, "", line) - gsub(/\}/, "", line)
}

function extract_fn_name(line,   name) {
  name = line
  sub(/^.*\bfn[[:space:]]+/, "", name)
  sub(/[[:space:]]*\(.*/, "", name)
  return name
}

BEGIN {
  in_fn = 0
  seen_open = 0
  depth = 0
  start = 0
  count = 0
}

{
  if (!in_fn && $0 ~ /^[[:space:]]*(pub([[:space:]]+[a-zA-Z0-9_]+)?[[:space:]]+)?(async[[:space:]]+)?(const[[:space:]]+)?(unsafe[[:space:]]+)?fn[[:space:]]+[A-Za-z_][A-Za-z0-9_]*[[:space:]]*(<[^>]*>)?[[:space:]]*\(/) {
    in_fn = 1
    seen_open = 0
    depth = 0
    start = NR
    count = 0
    fn_name = extract_fn_name($0)
  }

  if (!in_fn) {
    next
  }

  count++
  delta = count_braces($0)

  if (!seen_open) {
    if (delta > 0) {
      seen_open = 1
      depth = delta
    } else if ($0 ~ /;[[:space:]]*$/) {
      in_fn = 0
    }
  } else {
    depth += delta
  }

  if (seen_open && depth <= 0) {
    if (count >= 100) {
      printf "%s|%d|%d|%s\n", file, count, NR, fn_name
    }
    in_fn = 0
  }
}'
  awk -v file="$file" "$awk_script" "$file"
}

for crate_toml in "${REPO_ROOT}/crates"/*/Cargo.toml; do
  crate_dir="$(dirname "${crate_toml}")"
  crate_name="$(basename "${crate_dir}")"
  src_dir="${crate_dir}/src"

  if [ ! -d "${src_dir}" ]; then
    continue
  fi

  crate_files=0
  crate_prod_files=0
  crate_loc=0
  crate_over_files=0
  crate_long_fn=0
  crate_panic=0

  while IFS= read -r -d '' file; do
    file_rel="${file#${REPO_ROOT}/}"
    ((crate_files += 1))
    ((TOTAL_FILES += 1))

    loc="$(wc -l < "${file}")"
    ((crate_loc += loc))
    ((TOTAL_LOC += loc))

    if (( loc > MAX_FILE_LOC )); then
      printf "%s\t%d\t%s\n" "${file_rel}" "${loc}" "${crate_name}" >> "${FILES_OVER}"
      ((crate_over_files += 1))
      ((TOTAL_FILES_OVER += 1))
    fi

    if is_production_file "${file}"; then
      ((crate_prod_files += 1))
      ((TOTAL_PROD_FILES += 1))

      # shellcheck disable=SC2312
      panic_count="$(scan_panic_like_count "${file}")"
      if [ "${panic_count}" -gt 0 ]; then
        printf "%s\t%d\n" "${file_rel}" "${panic_count}" >> "${PANIC_FILES}"
        ((crate_panic += panic_count))
        ((TOTAL_PANIC += panic_count))
      fi

      while IFS='|' read -r fn_file fn_len fn_end fn_name; do
        printf "%s\t%s\t%d\t%s\t%s\n" "${crate_name}" "${fn_file}" "${fn_len}" "${fn_end}" "${fn_name}" >> "${FN_OVER}"
        ((crate_long_fn += 1))
        ((TOTAL_FN_OVER += 1))
      done < <(scan_long_functions "${file}")
    fi
  done < <(rg --files -g '*.rs' -0 "${src_dir}")

  printf "%s\t%s\t%s\t%s\t%s\t%s\t%s\n" \
    "${crate_name}" \
    "${crate_files}" \
    "${crate_prod_files}" \
    "${crate_loc}" \
    "${crate_over_files}" \
    "${crate_long_fn}" \
    "${crate_panic}" >> "${CRATE_SUMMARY}"

done

while IFS= read -r -d '' toml_file; do
  crate_dir="$(dirname "${toml_file}")"
  crate_name="$(basename "${crate_dir}")"

  while IFS= read -r dep; do
    if ! is_allowed_dependency "${crate_name}" "${dep}"; then
      printf "%s\t%s\n" "${crate_name}" "${dep}" >> "${ARCH_DEP_VIOL}"
      ((TOTAL_ARCH_DEP_VIOL += 1))
    fi
  done < <(
    awk '
      /^\[dependencies\]/ {in_deps=1; next}
      /^\[[^\]]+\]/ {in_deps=0}
      in_deps && $0 ~ /^[[:space:]]*[A-Za-z0-9_-]+[[:space:]]*=/ {
        dep=$0
        gsub(/^[[:space:]]*/, "", dep)
        sub(/[[:space:]]*=.*/, "", dep)
        print dep
      }
    ' "${toml_file}"
  )
done < <(find "${REPO_ROOT}/crates" -maxdepth 2 -name Cargo.toml -print0)

if rg -n -e "use quick-xml|use quick_xml|use zip::|use flate2::|use calamine::|use pyo3::|use clap::|use env_logger::|std::fs::" \
  "${REPO_ROOT}/crates/docir-core/src" > "${ARCH_INFRA}" 2>/dev/null; then
  :
fi

cat > "${REPORT}" <<EOF
# Quality Baseline – Phase 1

Generated: ${TS}
Output: ${REPORT}

## Thresholds
- Max file LOC: ${MAX_FILE_LOC}
- Max function LOC (heuristic): ${MAX_FN_LOC}
- Production ban: unwrap/expect/panic/unreachable

## Global summary

| Metric | Value |
|---|---:|
| Rust files (src) | ${TOTAL_FILES} |
| Production files (src) | ${TOTAL_PROD_FILES} |
| Total LOC | ${TOTAL_LOC} |
| Files over 800 LOC | ${TOTAL_FILES_OVER} |
| Functions over 100 LOC (heuristic) | ${TOTAL_FN_OVER} |
| Panic/unwrap/expect calls in production | ${TOTAL_PANIC} |
| Architecture dependency violations | ${TOTAL_ARCH_DEP_VIOL} |

## Crate summary

| Crate | Files | Prod files | LOC | Files > 800 LOC | Functions > 100 LOC | Panic-like calls in prod |
|---|---:|---:|---:|---:|---:|---:|
EOF

while IFS=$'\t' read -r crate_name crate_files crate_prod_files crate_loc crate_over_files crate_long_fn crate_panic; do
  printf "| %s | %s | %s | %s | %s | %s | %s |\n" \
    "${crate_name}" "${crate_files}" "${crate_prod_files}" "${crate_loc}" \
    "${crate_over_files}" "${crate_long_fn}" "${crate_panic}" >> "${REPORT}"
done < "${CRATE_SUMMARY}"

cat >> "${REPORT}" <<'EOF'

## Files exceeding 800 LOC
EOF
if [ -s "${FILES_OVER}" ]; then
  printf "| File | LOC | Crate |\n|---|---:|---|\n" >> "${REPORT}"
  sort -t$'\t' -k2,2nr "${FILES_OVER}" | while IFS=$'\t' read -r file loc crate; do
    printf "| %s | %s | %s |\n" "${file}" "${loc}" "${crate}" >> "${REPORT}"
  done
else
  printf "None\n" >> "${REPORT}"
fi

cat >> "${REPORT}" <<'EOF'

## Production functions over 100 LOC (heuristic)
EOF
if [ -s "${FN_OVER}" ]; then
  printf "| Crate | File | LOC | Start line | Function |\n|---|---|---:|---:|---|\n" >> "${REPORT}"
  sort -t$'\t' -k3,3nr "${FN_OVER}" | while IFS=$'\t' read -r crate file loc end_line fn_name; do
    printf "| %s | %s | %s | %s | %s |\n" "${crate}" "${file}" "${loc}" "${end_line}" "${fn_name}" >> "${REPORT}"
  done
else
  printf "None\n" >> "${REPORT}"
fi

cat >> "${REPORT}" <<'EOF'

## Production panic/unwrap/expect/unreachable usage
EOF
if [ -s "${PANIC_FILES}" ]; then
  printf "| File | Count |\n|---|---:|\n" >> "${REPORT}"
  sort -t$'\t' -k2,2nr "${PANIC_FILES}" | while IFS=$'\t' read -r file count; do
    printf "| %s | %s |\n" "${file}" "${count}" >> "${REPORT}"
  done
else
  printf "None\n" >> "${REPORT}"
fi

cat >> "${REPORT}" <<'EOF'

## Architecture dependency checks (Cargo.toml [dependencies] only)
EOF
if [ -s "${ARCH_DEP_VIOL}" ]; then
  printf "| Crate | Dependency |\n|---|---|\n" >> "${REPORT}"
  sort -t$'\t' -k1,1 "${ARCH_DEP_VIOL}" | while IFS=$'\t' read -r crate dep; do
    printf "| %s | %s |\n" "${crate}" "${dep}" >> "${REPORT}"
  done
else
  printf "None\n" >> "${REPORT}"
fi

cat >> "${REPORT}" <<'EOF'

## Infrastructure leakage signals inside docir-core
EOF
if [ -s "${ARCH_INFRA}" ]; then
  printf "Potential leaks found:\n\n" >> "${REPORT}"
  sed -n '1,200p' "${ARCH_INFRA}" >> "${REPORT}"
else
  printf "None\n" >> "${REPORT}"
fi

cat "${REPORT}"

if [ "${MODE_FAIL}" -eq 1 ]; then
  if [ "${TOTAL_FILES_OVER}" -gt 0 ] || [ "${TOTAL_FN_OVER}" -gt 0 ] || [ "${TOTAL_PANIC}" -gt 0 ] || [ "${TOTAL_ARCH_DEP_VIOL}" -gt 0 ]; then
    echo "Phase-1 baseline violations found: failing by request --fail-on-violations" >&2
    exit 1
  fi
fi

exit 0
