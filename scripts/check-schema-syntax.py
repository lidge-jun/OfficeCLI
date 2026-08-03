#!/usr/bin/env python3
"""Syntax + local $ref check for schemas/interfaces.

Deliberately limited: this parses JSON and resolves file-local $ref targets.
It does NOT validate documents against the schemas -- that needs a real
resolver and belongs in the wp6 test suite with JsonSchema.Net.

The point here is to catch the cheap failure early: a schema that $refs a file
nobody restored. error-result.v1.schema.json is referenced by two other
schemas, and omitting it was a real audit finding.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SCHEMA_DIR = ROOT / "schemas" / "interfaces"


EXPECTED_MIN = 12  # the restored interface-schema set; guards against a silent shrink


def refs(node) -> list[str]:
    found: list[str] = []
    if isinstance(node, dict):
        for k, v in node.items():
            if k == "$ref" and isinstance(v, str):
                found.append(v)
            else:
                found.extend(refs(v))
    elif isinstance(node, list):
        for item in node:
            found.extend(refs(item))
    return found


def main() -> int:
    if not SCHEMA_DIR.is_dir():
        print(f"no schema dir: {SCHEMA_DIR}", file=sys.stderr)
        return 2

    files = sorted(SCHEMA_DIR.glob("*.json"))
    if not files:
        print(f"no schemas found in {SCHEMA_DIR}", file=sys.stderr)
        return 2
    if len(files) < EXPECTED_MIN:
        # "all present schemas parsed" is trivially true of a directory someone
        # emptied. Require the set not to shrink.
        print(
            f"FAIL only {len(files)} schema(s) in {SCHEMA_DIR}, expected >= {EXPECTED_MIN}",
            file=sys.stderr,
        )
        return 1

    bad = 0
    for path in files:
        try:
            doc = json.loads(path.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError) as exc:
            print(f"FAIL parse {path.name}: {exc}")
            bad += 1
            continue

        for ref in refs(doc):
            if ref.startswith("#"):
                continue  # internal pointer
            target = (path.parent / ref.split("#", 1)[0]).resolve()
            if not target.exists():
                print(f"FAIL {path.name}: $ref -> {ref} (missing)")
                bad += 1

    if bad:
        print(f"{bad} schema problem(s)")
        return 1

    print(f"schemas OK: {len(files)} parsed, all local $ref targets present")
    return 0


if __name__ == "__main__":
    sys.exit(main())
