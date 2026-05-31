#!/usr/bin/env python3
"""Add deterministic supplier metadata required by the SBOM quality gate."""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any


PROJECT_SUPPLIER = "docir project"
DEFAULT_CARGO_SUPPLIER = "crates.io"


def supplier_for(component: dict[str, Any]) -> dict[str, str]:
    purl = str(component.get("purl", ""))
    name = str(component.get("name", ""))
    author = str(component.get("author", "")).strip()

    if purl.startswith("pkg:cargo/docir-") or name.startswith("docir-"):
        return {"name": PROJECT_SUPPLIER}
    if author:
        return {"name": author}
    return {"name": DEFAULT_CARGO_SUPPLIER}


def enrich_component(component: dict[str, Any]) -> None:
    component.setdefault("supplier", supplier_for(component))
    if not component.get("author"):
        component["author"] = component["supplier"]["name"]

    for child in component.get("components", []) or []:
        if isinstance(child, dict):
            enrich_component(child)


def main() -> int:
    if len(sys.argv) != 3:
        print("usage: quality_sbom_enrich.py <input.cdx.json> <output.cdx.json>", file=sys.stderr)
        return 2

    input_path = Path(sys.argv[1])
    output_path = Path(sys.argv[2])

    with input_path.open(encoding="utf-8") as handle:
        sbom = json.load(handle)

    metadata = sbom.setdefault("metadata", {})
    metadata.setdefault("supplier", {"name": PROJECT_SUPPLIER})

    primary_component = metadata.get("component")
    if isinstance(primary_component, dict):
        enrich_component(primary_component)

    for component in sbom.get("components", []) or []:
        if isinstance(component, dict):
            enrich_component(component)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8") as handle:
        json.dump(sbom, handle, indent=2, sort_keys=True)
        handle.write("\n")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
