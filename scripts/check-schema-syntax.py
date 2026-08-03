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


# The exact restored set. A count threshold is not enough: dropping a required
# schema and adding any unrelated JSON file would still satisfy ">= 12".
EXPECTED = {
    "capability-result.v1.schema.json",
    "compatibility-corpus.v1.schema.json",
    "diff-result.v1.schema.json",
    "edit-result.v1.schema.json",
    "error-result.v1.schema.json",
    "expected-capabilities.v1.schema.json",
    "rhwp-provider-capabilities.v1.schema.json",
    "rhwp-sidecar-request.v1.schema.json",
    "rhwp-sidecar-response.v1.schema.json",
    "save-policy.v1.schema.json",
    "save-transaction.v1.schema.json",
    "validation-result.v1.schema.json",
}


def resolve_pointer(doc, pointer: str):
    """Walk a JSON Pointer fragment; return None when any step is missing."""
    node = doc
    for raw in pointer.lstrip("#").strip("/").split("/"):
        if not raw:
            continue
        key = raw.replace("~1", "/").replace("~0", "~")
        if isinstance(node, dict):
            if key not in node:
                return None
            node = node[key]
        elif isinstance(node, list):
            try:
                node = node[int(key)]
            except (ValueError, IndexError):
                return None
        else:
            return None
    return node


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
    present = {p.name for p in files}
    missing = EXPECTED - present
    if missing:
        for name in sorted(missing):
            print(f"FAIL required schema missing: {name}")
        return 1

    parsed: dict[Path, object] = {}
    bad = 0
    for path in files:
        try:
            doc = json.loads(path.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError) as exc:
            print(f"FAIL parse {path.name}: {exc}")
            bad += 1
            continue
        parsed[path] = doc

    for path, doc in parsed.items():
        for ref in refs(doc):
            file_part, _, fragment = ref.partition("#")
            if ref.startswith("#"):
                # Internal pointers were previously skipped, so deleting a $defs
                # entry still reported "all targets present".
                if resolve_pointer(doc, ref) is None:
                    print(f"FAIL {path.name}: internal $ref -> {ref} (unresolvable)")
                    bad += 1
                continue
            target = (path.parent / file_part).resolve()
            if not target.is_file():
                print(f"FAIL {path.name}: $ref -> {ref} (missing)")
                bad += 1
                continue
            if fragment:
                other = parsed.get(target)
                if other is None:
                    try:
                        other = json.loads(target.read_text(encoding="utf-8"))
                    except (json.JSONDecodeError, OSError) as exc:
                        print(f"FAIL {path.name}: $ref target {file_part} unreadable: {exc}")
                        bad += 1
                        continue
                if resolve_pointer(other, fragment) is None:
                    print(f"FAIL {path.name}: $ref -> {ref} (fragment unresolvable)")
                    bad += 1

    if bad:
        print(f"{bad} schema problem(s)")
        return 1

    print(f"schemas OK: {len(files)} parsed, all local $ref targets present")
    return 0


if __name__ == "__main__":
    sys.exit(main())
