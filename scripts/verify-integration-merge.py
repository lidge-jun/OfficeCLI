#!/usr/bin/env python3
"""Verify the fork-main integration against its path decision ledger."""

from __future__ import annotations

import argparse
import csv
import json
import re
import subprocess
import sys
from pathlib import Path


def git(repo: Path, *args: str, check: bool = True) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        ["git", "-C", str(repo), *args],
        check=check,
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )


def fail(message: str) -> None:
    print(f"FAIL integration-ledger: {message}", file=sys.stderr)
    raise SystemExit(1)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--ledger", required=True, type=Path)
    parser.add_argument("--repo", required=True, type=Path)
    parser.add_argument("--base-feature", required=True)
    parser.add_argument("--head", default="HEAD")
    parser.add_argument("--checks", required=True, type=Path)
    args = parser.parse_args()

    repo = args.repo.resolve()
    rows = list(csv.DictReader(args.ledger.open(newline=""), delimiter="\t"))
    required_columns = {"status", "path", "decision", "reason"}
    if not rows or set(rows[0]) != required_columns:
        fail(f"unexpected ledger columns: {set(rows[0]) if rows else 'empty'}")

    by_path: dict[str, dict[str, str]] = {}
    for row in rows:
        path = row["path"]
        if path in by_path:
            fail(f"duplicate path: {path}")
        if row["status"] not in {"A", "M"}:
            fail(f"unsupported status {row['status']} for {path}")
        if row["decision"] not in {"adapt", "allow-add", "preserve-feature", "reject"}:
            fail(f"unsupported decision {row['decision']} for {path}")
        by_path[path] = row

    diff = git(repo, "diff", "--name-status", "--no-renames", f"{args.base_feature}..{args.head}").stdout
    actual: dict[str, str] = {}
    for line in diff.splitlines():
        status, path = line.split("\t", 1)
        if path in actual:
            fail(f"duplicate diff path: {path}")
        actual[path] = status
        if path not in by_path:
            fail(f"changed path is absent from ledger: {status} {path}")

    checks: dict[str, str] = json.loads(args.checks.read_text())
    adapted = {path for path, row in by_path.items() if row["decision"] == "adapt"}
    if set(checks) != adapted:
        missing = sorted(adapted - set(checks))
        extra = sorted(set(checks) - adapted)
        fail(f"adapt-check mismatch; missing={missing}, extra={extra}")

    for path, row in by_path.items():
        status = row["status"]
        decision = row["decision"]
        actual_status = actual.get(path)

        if decision == "adapt":
            if actual_status not in {"A", "M"}:
                fail(f"adapt path was not integrated: {path} (actual={actual_status})")
            content = git(repo, "show", f"{args.head}:{path}").stdout
            if not re.search(checks[path], content, re.MULTILINE | re.DOTALL):
                fail(f"activation pattern missing for adapted path: {path}")
        elif decision == "allow-add":
            if status != "A" or actual_status != "A":
                fail(f"allow-add path must be newly added: {path} (ledger={status}, actual={actual_status})")
        elif decision == "preserve-feature":
            if actual_status is not None:
                fail(f"preserve-feature path changed from feature base: {path} ({actual_status})")
        elif decision == "reject":
            if actual_status is not None:
                fail(f"rejected path changed from feature base: {path} ({actual_status})")
            if status == "A":
                exists = git(repo, "cat-file", "-e", f"{args.head}:{path}", check=False)
                if exists.returncode == 0:
                    fail(f"rejected merge-added artifact exists in head: {path}")

    markers = git(
        repo,
        "grep",
        "-n",
        "-E",
        "^(<{7} |={7}$|>{7} )",
        args.head,
        "--",
        check=False,
    )
    if markers.returncode == 0:
        fail(f"conflict markers found:\n{markers.stdout}")
    if markers.returncode not in {0, 1}:
        fail(f"git grep failed: {markers.stderr.strip()}")

    print(
        "PASS integration-ledger: "
        f"{len(rows)} decisions, {len(actual)} integrated paths, "
        f"{len(adapted)} adapted activation checks"
    )


if __name__ == "__main__":
    main()
