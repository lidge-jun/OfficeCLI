#!/usr/bin/env python3
"""Parse VSTest TRX results with a real XML parser.

Console output does not guarantee emitting passing test names, and grep
pipelines return non-zero on an empty match -- which trips `set -o pipefail`
and turns "zero skipped tests" (a good outcome) into a gate failure.
"""
from __future__ import annotations

import argparse
import sys
import xml.etree.ElementTree as ET

NS = {"t": "http://microsoft.com/schemas/VisualStudio/TeamTest/2010"}


def results(path: str):
    root = ET.parse(path).getroot()
    for r in root.iter():
        if r.tag.endswith("UnitTestResult"):
            name = r.get("testName")
            if name:
                yield name, (r.get("outcome") or "")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--file", required=True)
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--list-all", action="store_true")
    g.add_argument("--list-skipped", action="store_true")
    g.add_argument("--list-failed", action="store_true")
    g.add_argument("--count-matching", metavar="SUBSTR")
    ap.add_argument("--summary", action="store_true")
    a = ap.parse_args()

    try:
        rows = list(results(a.file))
    except (ET.ParseError, OSError) as exc:
        print(f"trx: cannot read {a.file}: {exc}", file=sys.stderr)
        return 2

    if a.list_all:
        for n, _ in rows:
            print(n)
    elif a.list_skipped:
        # VSTest reports skips as NotExecuted.
        for n, o in rows:
            if o == "NotExecuted":
                print(n)
    elif a.list_failed:
        for n, o in rows:
            if o == "Failed":
                print(n)
    else:
        print(len({n for n, _ in rows if a.count_matching in n}))

    if a.summary:
        passed = sum(1 for _, o in rows if o == "Passed")
        failed = sum(1 for _, o in rows if o == "Failed")
        skipped = sum(1 for _, o in rows if o == "NotExecuted")
        print(
            f"total={len(rows)} passed={passed} failed={failed} skipped={skipped}",
            file=sys.stderr,
        )
    return 0


if __name__ == "__main__":
    sys.exit(main())
