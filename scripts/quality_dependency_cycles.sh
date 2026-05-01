#!/usr/bin/env bash
set -euo pipefail
if ! command -v python3 >/dev/null 2>&1; then
  echo "dependency cycle policy: FAIL"
  echo "  Required tool missing: python3"
  exit 1
fi

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "${REPO_ROOT}"

tmp_metadata="$(mktemp)"
trap 'rm -f "${tmp_metadata}"' EXIT

if ! cargo metadata --format-version 1 --no-deps --offline >"${tmp_metadata}"; then
  echo "dependency cycle policy: FAIL"
  echo "  Failed to collect cargo metadata"
  exit 1
fi

python3 - "${tmp_metadata}" <<'PY'
import json
import sys

def dep_kinds(dependency):
    kind = dependency.get("kind")
    if kind is None:
        return {"normal"}
    if isinstance(kind, list):
        return set(kind)
    return {kind}


def main(metadata_path):
    with open(metadata_path, "r", encoding="utf-8") as metadata_file:
        data = json.load(metadata_file)
    members = set(data.get("workspace_members", []))
    packages = [p for p in data.get("packages", []) if p.get("id") in members]
    id_to_name = {p["id"]: p["name"] for p in packages}
    name_to_id = {p["name"]: p["id"] for p in packages}
    name_set = set(name_to_id.keys())

    graph = {name: set() for name in name_set}

    for p in packages:
        src_name = p["name"]
        for dependency in p.get("dependencies", []):
            target_name = dependency.get("name")
            if target_name not in name_set or target_name == src_name:
                continue

            kinds = dep_kinds(dependency)
            if "normal" not in kinds:
                continue

            graph[src_name].add(target_name)

    state = {}
    stack = []
    cycles = []

    def dfs(node):
        state[node] = 1  # visiting
        stack.append(node)
        for dep in sorted(graph.get(node, [])):
            if dep not in state:
                dfs(dep)
            elif state[dep] == 1:
                idx = stack.index(dep)
                cycles.append(stack[idx:] + [dep])
        state[node] = 2  # done
        stack.pop()

    for node in sorted(graph):
        if node not in state:
            dfs(node)

    if not cycles:
        print("dependency cycle policy: PASS")
        return

    print("dependency cycle policy: FAIL")
    for cycle in cycles:
        print("  cycle:", " -> ".join(cycle))
    raise SystemExit(1)


if __name__ == "__main__":
    main(sys.argv[1])
PY
