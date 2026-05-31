#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "${SCRIPT_DIR}/.." && pwd)"
SBOM_DIR="${REPO_ROOT}/target/sbom"
RAW_NAME="docir-quality-sbom"

cd "${REPO_ROOT}"

missing_tools=()
for tool in python3 sbomqs sbom-tools; do
  if ! command -v "${tool}" >/dev/null 2>&1; then
    missing_tools+=("${tool}")
  fi
done

if ! cargo cyclonedx --help >/dev/null 2>&1; then
  missing_tools+=("cargo-cyclonedx")
fi

if [ "${#missing_tools[@]}" -gt 0 ]; then
  printf 'SBOM quality gate skipped; missing tooling: %s\n' "${missing_tools[*]}"
  exit 0
fi

mkdir -p "${SBOM_DIR}"
rm -f crates/*/"${RAW_NAME}.json"
trap 'rm -f crates/*/"${RAW_NAME}.json"' EXIT

cargo cyclonedx \
  --manifest-path Cargo.toml \
  --format json \
  --all-features \
  --spec-version 1.5 \
  --override-filename "${RAW_NAME}"

shopt -s nullglob
raw_sboms=(crates/*/"${RAW_NAME}.json")
if [ "${#raw_sboms[@]}" -eq 0 ]; then
  echo "SBOM quality gate failed: cargo-cyclonedx produced no crate SBOMs"
  exit 1
fi

for raw_sbom in "${raw_sboms[@]}"; do
  crate_name="$(basename "$(dirname "${raw_sbom}")")"
  enriched_sbom="${SBOM_DIR}/${crate_name}.cdx.json"

  python3 "${SCRIPT_DIR}/quality_sbom_enrich.py" "${raw_sbom}" "${enriched_sbom}"
  sbom-tools validate "${enriched_sbom}" --standard ntia -q

  score_line="$(sbomqs score --profile ntia --basic "${enriched_sbom}")"
  score="$(awk '{print $1}' <<<"${score_line}")"
  grade="$(awk '{print $2}' <<<"${score_line}")"

  if [ "${grade}" != "A" ] || ! awk -v score="${score}" 'BEGIN { exit !(score >= 10.0) }'; then
    printf 'SBOM quality gate failed for %s: %s\n' "${crate_name}" "${score_line}"
    exit 1
  fi

  printf 'SBOM quality gate passed for %s: %s\n' "${crate_name}" "${score_line}"
done
