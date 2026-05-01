#!/usr/bin/env bash
set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"

MIN_COUNT="${1:-3}"
MIN_LINES="${2:-12}"

if ! [[ "${MIN_COUNT}" =~ ^[0-9]+$ && "${MIN_LINES}" =~ ^[0-9]+$ ]]; then
  echo "Usage: ./scripts/quality_duplicate_patterns.sh [min_count] [min_lines]" >&2
  echo "  min_count default: 3" >&2
  echo "  min_lines default: 12" >&2
  exit 2
fi

TS="$(date -u +"%Y%m%dT%H%M%SZ")"

python3 - "${REPO_ROOT}" "${MIN_COUNT}" "${MIN_LINES}" <<'PY'
from __future__ import annotations

import hashlib
import re
from collections import defaultdict
from pathlib import Path
import sys

repo_root = Path(sys.argv[1])
min_count = int(sys.argv[2])
min_lines = int(sys.argv[3])

RUST_FILES = sorted(repo_root.glob("crates/*/src/**/*.rs")) + sorted(repo_root.glob("crates/*/src/*.rs"))

keywords = {
    "as", "async", "await", "break", "const", "continue", "crate", "dyn", "else", "enum", "extern",
    "false", "fn", "for", "if", "impl", "in", "let", "loop", "match", "move", "mut", "pub", "return",
    "self", "Self", "static", "struct", "super", "trait", "true", "type", "union", "use", "where", "while",
    "use", "use", "unsafe", "dyn", "ref", "self", "super", "crate", "vec", "Some", "None", "Ok", "Err",
}

fn_sig_re = re.compile(
    r"^\s*(?:pub(?:\([^\)]*\))?\s+)?(?:async\s+)?(?:const\s+)?(?:unsafe\s+)?fn\s+([A-Za-z_][A-Za-z0-9_]*)\b"
)


def is_production_file(path: Path) -> bool:
    p = str(path)
    if "/tests/" in p:
        return False
    if path.name in {"tests.rs", "test.rs"}:
        return False
    if path.name.endswith("_tests.rs") or path.name.endswith("_test.rs"):
        return False
    return True


def strip_comments_and_strings(text: str) -> str:
    # Remove block comments first (best effort).
    text = re.sub(r"/\*.*?\*/", " ", text, flags=re.S)

    cleaned_lines = []
    for line in text.splitlines():
        line = line.split("//", 1)[0]
        cleaned_lines.append(line)

    text = "\n".join(cleaned_lines)
    text = re.sub(r"r#*\".*?\"#*", '""', text, flags=re.S)
    text = re.sub(r'"(\\.|[^"\\])*"', '""', text)
    text = re.sub(r"'([^'\\]|\\.)*'", "''", text)
    return text


def normalize_snippet(text: str) -> str:
    text = strip_comments_and_strings(text)
    text = re.sub(r"\b\d+\b", "N", text)

    def normalize_token(match: re.Match[str]) -> str:
        token = match.group(0)
        if token in keywords or len(token) <= 1:
            return token
        return "ID"

    text = re.sub(r"[A-Za-z_][A-Za-z0-9_]*", normalize_token, text)
    text = re.sub(r"\s+", " ", text)
    text = text.strip()
    return text


def extract_functions(path: Path):
    text = path.read_text(encoding="utf-8")
    lines = text.splitlines()

    i = 0
    n = len(lines)
    while i < n:
        m = fn_sig_re.match(lines[i])
        if not m:
            i += 1
            continue

        fn_name = m.group(1)
        start_line = i + 1
        fn_lines = []
        brace_depth = 0
        opened = False
        j = i

        while j < n:
            current = lines[j]
            fn_lines.append(current)
            opens = current.count("{")
            closes = current.count("}")

            if not opened:
                brace_depth += opens - closes
                if brace_depth > 0:
                    opened = True
                elif opens == 0:
                    j += 1
                    continue
                elif brace_depth <= 0:
                    break
            else:
                brace_depth += opens - closes
                if brace_depth <= 0:
                    break

            j += 1

        if opened and brace_depth <= 0:
            end_line = j + 1
            snippet = "\n".join(fn_lines)
            loc = len(snippet.splitlines())
            yield path, fn_name, start_line, end_line, loc, snippet

        i = j + 1


signature_groups = defaultdict(list)

for file in RUST_FILES:
    if not is_production_file(file):
        continue

    for file_path, fn_name, start_line, end_line, loc, snippet in extract_functions(file):
        if loc < min_lines:
            continue

        signature = normalize_snippet(snippet)
        if not signature:
            continue

        digest = hashlib.sha1(signature.encode("utf-8")).hexdigest()
        signature_groups[digest].append(
            {
                "file": str(file_path),
                "function": fn_name,
                "start_line": start_line,
                "end_line": end_line,
                "loc": loc,
                "signature": signature,
            }
        )

dup_groups = []
for group in signature_groups.values():
    if len(group) >= min_count:
        dup_groups.append(group)

dup_groups.sort(key=lambda g: (-(len(g)), -sum(item["loc"] for item in g) / len(g)))

print(f"Found {len(dup_groups)} duplicate function-pattern groups with threshold >= {min_count} occurrences")
print()
if not dup_groups:
    print("No duplicate function patterns detected under the configured threshold.")
    raise SystemExit(0)

print("| Count | Avg LOC | Files | Sample function(s) | Locations |")
print("|---|---:|---|---|---|")

for group in dup_groups:
    locations = []
    function_names = set()
    file_names = set()
    for item in group:
        locations.append(f"{item['file']}:{item['start_line']}")
        file_names.add(item["file"])
        function_names.add(item["function"])

    avg_loc = round(sum(item["loc"] for item in group) / len(group), 1)
    sample_function = next(iter(function_names))
    locations_text = "<br>".join(locations[:8])
    if len(locations) > 8:
        locations_text += "<br>..."

    print(
        f"| {len(group)} | {avg_loc} | {len(file_names)} | "
        f"{sample_function} (+{len(function_names)-1}) | {locations_text} |"
    )
PY
