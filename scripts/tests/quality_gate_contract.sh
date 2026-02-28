#!/usr/bin/env bash
set -euo pipefail

canonical="scripts/quality_gate.sh"

if [ ! -f "${canonical}" ]; then
  echo "Missing canonical quality gate at ${canonical}"
  exit 1
fi

if [ ! -x "${canonical}" ]; then
  echo "Canonical quality gate is not executable: ${canonical}"
  exit 1
fi

alt_found=0
while IFS= read -r file; do
  [ "$file" = "${canonical}" ] && continue

  if [ -x "$file" ]; then
    echo "Alternate executable gate-like script detected: ${file}"
    alt_found=1
  fi
done < <(find scripts -maxdepth 2 -type f | rg '(gate|quality|check)')

if [ "$alt_found" -ne 0 ]; then
  exit 1
fi

echo "quality_gate_contract: OK"
