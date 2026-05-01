#!/usr/bin/env bash
set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "${REPO_ROOT}"

THRESHOLD="${CC13_COMPLEXITY_THRESHOLD:-10}"
MAX_FILE_LOC="${API_HYGIENE_MAX_FILE_LOC:-800}"
MAX_PUBLIC_FN_LOC="${API_HYGIENE_MAX_PUBLIC_FN_LOC:-100}"
ENFORCE_PR_RULES="${API_HYGIENE_ENFORCE_PR_RULES:-0}"
FAIL=0

if ! command -v python3 >/dev/null 2>&1; then
  echo "api_hygiene policy: FAIL"
  echo "  Required tool missing: python3"
  exit 1
fi

if [ ! -d "crates" ]; then
  echo "api_hygiene policy: FAIL"
  echo "  Missing expected workspace directory: crates"
  exit 1
fi

RUSTFLAGS_WITH_DENIES="${RUSTFLAGS:+${RUSTFLAGS} }--deny dead_code --deny unused_imports"
if ! RUSTFLAGS="${RUSTFLAGS_WITH_DENIES}" cargo check --workspace --all-targets --all-features; then
  echo "api_hygiene policy: FAIL"
  echo "  Rust check reported dead code or unused imports."
  FAIL=1
fi

python3 - "${THRESHOLD}" "${MAX_FILE_LOC}" "${MAX_PUBLIC_FN_LOC}" "${ENFORCE_PR_RULES}" <<'PY'
#!/usr/bin/env python3
from __future__ import annotations

import re
import sys
from pathlib import Path
from typing import List, Set, Tuple


THRESHOLD = int(sys.argv[1])
MAX_FILE_LOC = int(sys.argv[2])
MAX_PUBLIC_FN_LOC = int(sys.argv[3])
ENFORCE_PR_RULES = int(sys.argv[4]) == 1
ROOT = Path("crates")
CRITICAL_API_CRATES: Set[str] = {
    "docir-parser",
    "docir-app",
    "docir-diff",
    "docir-security",
}

pub_fn_pattern = re.compile(
    r"^\s*pub\s+(?:async\s+)?(?:const\s+)?(?:unsafe\s+)?(?:extern\s+\"C\"\s+)?fn\s+([A-Za-z_][A-Za-z0-9_]*)\b"
)
pub_type_pattern = re.compile(
    r"^\s*pub(?:\([^)]+\))?\s+(?:unsafe\s+)?(?:async\s+)?(?:const\s+)?(struct|enum|trait|type|union)\s+([A-Za-z_][A-Za-z0-9_]*)\b"
)
pub_const_pattern = re.compile(
    r"^\s*pub(?:\([^)]+\))?\s+(?:unsafe\s+)?(?:const|static)\s+([A-Za-z_][A-Za-z0-9_]*)\b"
)

test_mod_open_pattern = re.compile(r"^\s*mod\s+tests\b.*\{?\s*$")
cfg_test_pattern = re.compile(r"^\s*#\[\s*cfg\s*\(\s*test\s*\)\s*\]\s*$")
doc_comment_line = re.compile(r"^\s*//!|^\s*///|^\s*/\*\*")


def is_skippable_file(path: Path) -> bool:
    if "/tests/" in str(path):
        return True
    if str(path).endswith("tests.rs") or str(path).endswith("_tests.rs") or str(path).endswith("/test.rs"):
        return True
    if "/src/" not in str(path):
        return True
    return False


def is_critical_api_file(path: Path) -> bool:
    if path.name != "lib.rs":
        return False

    parts = path.parts
    try:
        crate_root = parts.index("crates")
    except ValueError:
        return False
    if crate_root + 1 >= len(parts):
        return False
    return parts[crate_root + 1] in CRITICAL_API_CRATES


def find_files() -> List[Path]:
    return sorted(set(ROOT.glob("*/src/**/*.rs")) | set(ROOT.glob("*/src/*.rs")))


def start_raw_string(line: str, i: int) -> Tuple[int, int] | None:
    if line.startswith("br", i):
        j = i + 2
        hash_count = 0
        while j < len(line) and line[j] == "#":
            hash_count += 1
            j += 1
        if j < len(line) and line[j] == "\"":
            return j + 1, hash_count

    if line.startswith("rb", i):
        j = i + 2
        hash_count = 0
        while j < len(line) and line[j] == "#":
            hash_count += 1
            j += 1
        if j < len(line) and line[j] == "\"":
            return j + 1, hash_count

    if line.startswith("r", i):
        j = i + 1
        hash_count = 0
        while j < len(line) and line[j] == "#":
            hash_count += 1
            j += 1
        if j < len(line) and line[j] == "\"":
            return j + 1, hash_count

    return None


def strip_code_lines(lines: List[str]) -> List[str]:
    normalized = []
    in_block_comment = False
    in_string = False
    in_char = False
    in_raw_string = False
    raw_hashes = 0

    for line in lines:
        src = line.rstrip("\n")
        out = []
        i = 0

        while i < len(src):
            ch = src[i]

            if in_block_comment:
                close_idx = src.find("*/", i)
                if close_idx == -1:
                    i = len(src)
                    continue
                i = close_idx + 2
                in_block_comment = False
                continue

            if in_string:
                if ch == "\\":
                    i += 2
                    continue
                if ch == "\"":
                    in_string = False
                out.append(" ")
                i += 1
                continue

            if in_char:
                if ch == "\\":
                    i += 2
                    continue
                if ch == "'":
                    in_char = False
                out.append(" ")
                i += 1
                continue

            if in_raw_string:
                end = "\""
                if raw_hashes > 0:
                    end = "\"" + ("#" * raw_hashes)
                if src.startswith(end, i):
                    out.append(" " * len(end))
                    in_raw_string = False
                    raw_hashes = 0
                    i += len(end)
                    continue
                out.append(" ")
                i += 1
                continue

            if src.startswith("//", i):
                break

            if src.startswith("/*", i):
                in_block_comment = True
                i += 2
                continue

            if ch == "\"":
                in_string = True
                out.append(" ")
                i += 1
                continue

            if ch == "'":
                in_char = True
                out.append(" ")
                i += 1
                continue

            raw_start = start_raw_string(src, i)
            if raw_start is not None:
                i = raw_start[0]
                in_raw_string = True
                raw_hashes = raw_start[1]
                out.append(" ")
                continue

            if ch in "{};":
                out.append(ch)
            else:
                out.append(ch)
            i += 1

        normalized.append("".join(out))

    return normalized


def complexity_delta(text: str) -> int:
    normalized = text
    total = 0
    total += normalized.count("else if")
    normalized = normalized.replace("else if", " ")
    for keyword in ("if", "match", "while", "for", "loop", "&&", "||"):
        if keyword in ("if", "match", "while", "for", "loop"):
            total += len(re.findall(rf"\b{re.escape(keyword)}\b", normalized))
        else:
            total += normalized.count(keyword)
    return total


def scan_file(path: Path) -> Tuple[int, int, int, int]:
    with path.open("r", encoding="utf-8", errors="replace") as source:
        raw_lines = source.readlines()

    cleaned = strip_code_lines(raw_lines)

    missing_docs = 0
    complexity_violations = 0
    fn_loc_violations = 0
    file_loc_violation = 0

    if len(raw_lines) > MAX_FILE_LOC:
        file_loc_violation = 1
        print(
            f"CC-14: file size threshold exceeded: {path}:{len(raw_lines)} > {MAX_FILE_LOC}"
        )

    track_public_type_docs = is_critical_api_file(path)
    in_tests_block = False
    pending_cfg_test = False
    test_depth = 0

    had_doc = False
    in_fn = False
    fn_depth = 0
    fn_complexity = 1
    fn_name = ""
    fn_start = 0
    fn_length = 0

    for lineno, (raw_line, clean_line) in enumerate(zip(raw_lines, cleaned), start=1):
        raw = raw_line.rstrip("\n")
        code = clean_line.rstrip("\n")
        stripped = raw.strip()

        if pending_cfg_test:
            if cfg_test_pattern.match(raw) is not None:
                continue
            if test_mod_open_pattern.match(raw) is not None:
                in_tests_block = True
                pending_cfg_test = False
                test_depth = raw.count("{") - raw.count("}")
                if test_depth <= 0:
                    in_tests_block = False
                    test_depth = 0
                continue
            if not stripped:
                continue
            pending_cfg_test = False

        if in_tests_block:
            test_depth += raw.count("{") - raw.count("}")
            if test_depth <= 0:
                in_tests_block = False
                test_depth = 0
            continue

        if cfg_test_pattern.match(raw) is not None:
            pending_cfg_test = True
            continue

        if raw.startswith("#[") and "cfg(test)" in raw:
            pending_cfg_test = True
            continue

        pub_fn_match = pub_fn_pattern.match(code)
        if pub_fn_match:
            fn_name = pub_fn_match.group(1)
            if not had_doc:
                missing_docs += 1
                print(f"CC-12: missing documentation: {path}:{lineno}:{fn_name}")

            fn_start = lineno
            fn_complexity = 1 + complexity_delta(code)
            fn_length = 1
            in_fn = True
            fn_depth = 0
            if "{" in code:
                fn_depth += code.count("{") - code.count("}")
                if fn_depth <= 0:
                    if fn_complexity > THRESHOLD:
                        complexity_violations += 1
                        print(
                            f"CC-13: complexity threshold exceeded: "
                            f"{path}:{lineno}:{fn_name} complexity={fn_complexity} > {THRESHOLD}"
                        )
                    if fn_length > MAX_PUBLIC_FN_LOC:
                        fn_loc_violations += 1
                        print(
                            f"CC-14: function LOC threshold exceeded: "
                            f"{path}:{lineno}:{fn_name} loc={fn_length} > {MAX_PUBLIC_FN_LOC}"
                        )
                    in_fn = False
            had_doc = False
            continue

        if track_public_type_docs:
            pub_type_match = pub_type_pattern.match(code)
            if pub_type_match:
                decl_name = pub_type_match.group(2)
                if not had_doc:
                    missing_docs += 1
                    print(f"CC-12: missing documentation: {path}:{lineno}:{decl_name}")
                had_doc = False
                continue

            pub_const_match = pub_const_pattern.match(code)
            if pub_const_match:
                decl_name = pub_const_match.group(1)
                if not had_doc:
                    missing_docs += 1
                    print(f"CC-12: missing documentation: {path}:{lineno}:{decl_name}")
                had_doc = False
                continue

        if raw.startswith("#["):
            if raw.startswith("#[doc"):
                had_doc = True
            continue

        if doc_comment_line.match(raw):
            had_doc = True
            continue

        if stripped == "":
            had_doc = False
            continue

        if in_fn:
            fn_complexity += complexity_delta(code)
            fn_length += 1
            fn_depth += code.count("{") - code.count("}")
            if fn_depth <= 0:
                if fn_complexity > THRESHOLD:
                    complexity_violations += 1
                    print(
                        f"CC-13: complexity threshold exceeded: {path}:{lineno}:{fn_name} "
                        f"starts at {fn_start} complexity={fn_complexity} > {THRESHOLD}"
                    )
                if fn_length > MAX_PUBLIC_FN_LOC:
                    fn_loc_violations += 1
                    print(
                        f"CC-14: function LOC threshold exceeded: {path}:{fn_start}:{fn_name} "
                        f"loc={fn_length} > {MAX_PUBLIC_FN_LOC}"
                    )
                in_fn = False
            continue

        if stripped.startswith(("pub ", "pub(", "pub\t")):
            had_doc = False
        elif raw.startswith("//") or raw.startswith("/*") or raw.startswith("*/"):
            had_doc = False
        elif not raw.startswith("#"):
            had_doc = False

    if in_fn:
        print(
            f"CC-13: unclosed public function body: {path}:{fn_start}:{fn_name} "
            f"complexity={fn_complexity}"
        )
        complexity_violations += 1
        if fn_length > MAX_PUBLIC_FN_LOC:
            fn_loc_violations += 1
            print(
                f"CC-14: function LOC threshold exceeded: {path}:{fn_start}:{fn_name} "
                f"loc={fn_length} > {MAX_PUBLIC_FN_LOC}"
            )

    return missing_docs, complexity_violations, fn_loc_violations, file_loc_violation


missing = 0
complexity = 0
fn_loc = 0
file_loc = 0
for file in find_files():
    if is_skippable_file(file):
        continue
    m, c, fl, fc = scan_file(file)
    missing += m
    complexity += c
    fn_loc += fl
    file_loc += fc

print(f"CC-12 count: {missing}")
print(f"CC-13 count: {complexity}")
print(f"CC-14-public-fn-loc count: {fn_loc}")
print(f"CC-14-lib-file-loc count: {file_loc}")
if not ENFORCE_PR_RULES and (fn_loc != 0 or file_loc != 0):
    print(
        "API hygiene guidance: set API_HYGIENE_ENFORCE_PR_RULES=1 to fail PRs "
        "when module/function size limits are exceeded."
    )
if missing != 0:
    raise SystemExit(1)
if complexity != 0:
    raise SystemExit(1)
if ENFORCE_PR_RULES and (fn_loc != 0 or file_loc != 0):
    raise SystemExit(1)
PY

if [ "${FAIL}" -eq 0 ]; then
  echo "api_hygiene policy: PASS"
  exit 0
fi

echo "api_hygiene policy: FAIL"
exit 1
