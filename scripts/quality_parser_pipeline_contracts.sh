#!/usr/bin/env bash
set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "${REPO_ROOT}"

status=0

check_contract() {
  local parser="$1"
  local file="$2"

if ! python3 - "$parser" "$file" <<'PY'
import re
import sys
from pathlib import Path

parser = sys.argv[1]
path = Path(sys.argv[2])
text = path.read_text(encoding="utf-8", errors="replace")

if not re.search(rf"impl\s+ParseStage\s+for\s+{re.escape(parser)}\b", text):
    raise SystemExit(1)

if not re.search(
    rf"impl\s+{re.escape(parser)}\s*{{[\s\S]*?pub\s+fn\s+parse_reader<R:\s*Read\s*\+\s*Seek>\s*\([^)]*\)\s*->\s*Result<ParsedDocument,\s*ParseError>\s*{{[\s\S]*?run_parser_pipeline\(self,\s*reader\)",
    text,
    re.MULTILINE,
):
    raise SystemExit(2)

if not re.search(
    rf"\bimpl\s+ParseStage\s+for\s+{re.escape(parser)}\b[\s\S]*?fn\s+parse_stage<",
    text,
    re.MULTILINE,
):
    raise SystemExit(3)
PY
  then
    echo "parser pipeline contract missing or invalid: ${parser} (${file})"
    status=1
  fi
}

check_contract "DocumentParser" "crates/docir-parser/src/parser/document.rs"
check_contract "OoxmlParser" "crates/docir-parser/src/parser/ooxml.rs"
check_contract "RtfParser" "crates/docir-parser/src/rtf/parser.rs"
check_contract "OdfParser" "crates/docir-parser/src/odf/builder.rs"
check_contract "HwpParser" "crates/docir-parser/src/hwp/builder.rs"
check_contract "HwpxParser" "crates/docir-parser/src/hwp/builder.rs"

if rg -q "parser::contracts::" crates/docir-parser/src; then
  echo "direct parser::contracts imports detected; use crate::parser reexports"
  status=1
fi

if rg -q "^pub\\(crate\\)\\s+mod\\s+contracts;" crates/docir-parser/src/parser.rs; then
  echo "contracts module should remain private in parser.rs"
  status=1
fi

if [ "${status}" -ne 0 ]; then
  echo "parser pipeline contracts: FAIL"
  exit "${status}"
fi

echo "parser pipeline contracts: PASS"
exit 0
