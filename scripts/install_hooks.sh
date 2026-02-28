#!/usr/bin/env bash
set -euo pipefail

expected_path=".githooks"

git config core.hooksPath "${expected_path}"
actual_path="$(git config --get core.hooksPath || true)"

if [ "${actual_path}" != "${expected_path}" ]; then
  echo "Failed to configure core.hooksPath. Expected '${expected_path}', got '${actual_path:-<unset>}'"
  exit 1
fi

echo "Configured core.hooksPath=${actual_path}"
