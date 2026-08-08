"""check_deps.py — Check module dependencies before deletion.

Usage:
    python scripts/check_deps.py <module_name>
    python scripts/check_deps.py enterprise/precision_renderer.py
    python scripts/check_deps.py precision_renderer

Prints:
    - Who imports this module (downstream impact)
    - What this module imports (upstream dependencies)
    - Risk level (CORE / ENTRY / INTERNAL / ISOLATED)
    - Deletion recommendation
"""
from __future__ import annotations

import ast
import os
import sys
from collections import defaultdict

INTERNAL_PREFIX = "ppt_pro_max."
SRC_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "src", "ppt_pro_max")


def _normalize(name: str) -> str:
    name = name.replace("/", ".").replace("\\", ".")
    if name.endswith(".py"):
        name = name[:-3]
    if not name.startswith("ppt_pro_max."):
        name = "ppt_pro_max." + name
    return name


def _collect_modules() -> dict[str, str]:
    modules = {}
    for root, _dirs, files in os.walk(SRC_DIR):
        for f in files:
            if f.endswith(".py") and not f.startswith("__"):
                rel = os.path.relpath(os.path.join(root, f), os.path.dirname(SRC_DIR))
                mod = rel.replace(os.sep, ".")[:-3]
                modules[mod] = os.path.join(root, f)
    return modules


def _parse_deps(modules: dict[str, str]) -> tuple[dict[str, set[str]], dict[str, set[str]]]:
    dep_graph = defaultdict(set)
    for mod, path in modules.items():
        try:
            tree = ast.parse(open(path, encoding="utf-8").read())
            for node in ast.walk(tree):
                if isinstance(node, ast.ImportFrom):
                    if node.module and node.module.startswith(INTERNAL_PREFIX):
                        dep_graph[mod].add(node.module)
                elif isinstance(node, ast.Import):
                    for alias in node.names:
                        if alias.name.startswith(INTERNAL_PREFIX):
                            dep_graph[mod].add(alias.name)
        except Exception:
            pass

    reverse = defaultdict(set)
    for mod, deps in dep_graph.items():
        for d in deps:
            reverse[d].add(mod)

    return dep_graph, reverse


def _classify_risk(mod: str, reverse: dict[str, set[str]]) -> str:
    importers = len(reverse.get(mod, set()))

    entry_modules = {
        "ppt_pro_max.enterprise.pipeline",
        "ppt_pro_max.renderer.ppt_renderer",
        "ppt_pro_max.build_helpers",
        "ppt_pro_max.cli",
        "ppt_pro_max.analyze_template",
        "ppt_pro_max.enterprise.proposal_generator",
    }
    if mod in entry_modules:
        return "ENTRY"
    if importers >= 3:
        return "CORE"
    if importers >= 1:
        return "INTERNAL"
    return "ISOLATED"


def check(module_name: str) -> None:
    modules = _collect_modules()
    dep_graph, reverse = _parse_deps(modules)

    mod = _normalize(module_name)

    if mod not in modules and mod not in reverse and mod not in dep_graph:
        print(f"Module not found: {mod}")
        print(f"Available modules (partial): {sorted(modules.keys())[:20]}...")
        sys.exit(1)

    risk = _classify_risk(mod, reverse)
    importers = reverse.get(mod, set())
    deps = dep_graph.get(mod, set())

    print(f"Module: {mod}")
    print(f"Risk:   {risk}")
    print()

    if importers:
        print(f"Imported BY ({len(importers)} modules — deletion breaks these):")
        for i in sorted(importers):
            i_risk = _classify_risk(i, reverse)
            print(f"  <- {i}  [{i_risk}]")
    else:
        print("Imported BY: (none — no internal module imports this)")

    print()

    if deps:
        print(f"Imports FROM ({len(deps)} modules — these become unused if caller deleted):")
        for d in sorted(deps):
            d_importers = reverse.get(d, set()) - {mod}
            still_used = "still used" if d_importers else "BECOMES ORPHAN"
            print(f"  -> {d}  [{still_used}]")
    else:
        print("Imports FROM: (none — this is a leaf module)")

    print()
    if risk == "CORE":
        print(">>> NEVER DELETE — Cascade failure will break {} modules".format(len(importers)))
    elif risk == "ENTRY":
        print(">>> NEVER DELETE — Entry point called by __init__.py or CLI")
    elif risk == "INTERNAL":
        print(">>> CAUTION — Update {} caller(s) before deleting".format(len(importers)))
    else:
        print(">>> SAFE to delete — No internal dependencies")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python scripts/check_deps.py <module_name>")
        print("Example: python scripts/check_deps.py precision_renderer")
        sys.exit(1)
    check(sys.argv[1])
