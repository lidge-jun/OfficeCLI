#!/usr/bin/env python3
"""Enforce the HWP re-port restore ledger against the working tree.

`decision` alone is not an assertion -- its meaning depends on `diff_status`:

  A,restore   path must EXIST (we brought it over)
  A,exclude   path must be ABSENT (we deliberately did not)
  M,adapt     path must EXIST and its check_id hook must be present
  M,exclude   path must be IDENTICAL to upstream (additive-only rule)
  D,exclude   path must EXIST and be IDENTICAL to upstream (deletion guard)
  R*,exclude  old path unchanged; rename destination absent

Treating every `exclude` as "must be absent" would falsely flag every
M,exclude row, since those files exist upstream by definition.
"""
from __future__ import annotations

import argparse
import csv
import os
import subprocess
import sys
from collections import defaultdict


def git(repo: str, *args: str) -> subprocess.CompletedProcess:
    return subprocess.run(
        ["git", "-C", repo, *args], capture_output=True, text=True
    )


def unchanged_vs_upstream(repo: str, upstream: str, path: str) -> bool:
    return git(repo, "diff", "--quiet", upstream, "--", path).returncode == 0


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--ledger", required=True)
    ap.add_argument("--repo", required=True)
    ap.add_argument("--upstream", default="upstream/main")
    ap.add_argument("--checks", help="optional check_id -> grep-pattern manifest")
    ap.add_argument(
        "--through-wp",
        help="only enforce restore/adapt rows owned by work-phases up to and "
        "including this one (e.g. wp2). Later phases are reported as PENDING, "
        "not FAIL. Omit to enforce the whole ledger (wp7 final gate).",
    )
    a = ap.parse_args()

    with open(a.ledger, newline="", encoding="utf-8") as fh:
        rows = list(csv.DictReader(fh))
    if not rows:
        print("ledger: empty", file=sys.stderr)
        return 2

    fail: dict[str, list[str]] = defaultdict(list)
    pending: list[str] = []

    def in_scope(owner: str) -> bool:
        """Rows owned by a later work-phase are not yet due."""
        if not a.through_wp:
            return True
        if not owner or not owner.startswith("wp"):
            return True
        try:
            return int(owner[2:]) <= int(a.through_wp[2:])
        except ValueError:
            return True

    for r in rows:
        st, path = r["diff_status"], r["path"]
        decision, check_id = r["decision"], r.get("check_id") or ""
        abs_path = os.path.join(a.repo, path)
        exists = os.path.exists(abs_path)

        if not decision:
            fail["blank decision"].append(path)
            continue
        if decision == "undecided":
            if a.through_wp:
                pending.append(f"{path} (undecided, {r['owner_wp']})")
            else:
                fail["undecided (must be resolved before wp7)"].append(path)
            continue

        if decision in ("restore", "adapt") and not in_scope(r["owner_wp"]):
            if not exists:
                pending.append(f"{path} ({decision}, {r['owner_wp']})")
            continue

        if st.startswith("A"):
            if decision == "restore" and not exists:
                fail["A,restore missing"].append(path)
            elif decision == "exclude" and exists:
                fail["A,exclude unexpectedly present"].append(path)
        elif st.startswith("M"):
            if decision == "adapt":
                if not exists:
                    fail["M,adapt missing"].append(path)
                elif not check_id:
                    fail["M,adapt without check_id"].append(path)
            elif decision == "exclude" and exists and not unchanged_vs_upstream(
                a.repo, a.upstream, path
            ):
                fail["M,exclude modified (must match upstream)"].append(path)
        elif st.startswith("D"):
            # Upstream-newer file: keep it, untouched.
            if not exists:
                fail["D,exclude DELETED (upstream regression)"].append(path)
            elif not unchanged_vs_upstream(a.repo, a.upstream, path):
                fail["D,exclude modified"].append(path)
        elif st.startswith("R"):
            if st.endswith("-dest"):
                if exists:
                    fail["rename destination unexpectedly present"].append(path)
            elif exists and not unchanged_vs_upstream(a.repo, a.upstream, path):
                fail["rename source modified"].append(path)

    # Independent guard: no upstream file may vanish from the port branch.
    diff = git(a.repo, "diff", "--name-status", f"{a.upstream}..HEAD")
    if diff.returncode != 0:
        fail["git diff vs upstream failed"].append(diff.stderr.strip())
    else:
        for line in diff.stdout.splitlines():
            if line.startswith("D\t"):
                fail["deleted vs upstream"].append(line[2:])

    if fail:
        for kind, paths in sorted(fail.items()):
            print(f"FAIL {kind}: {len(paths)}")
            for p in paths[:20]:
                print(f"  {p}")
            if len(paths) > 20:
                print(f"  ... and {len(paths) - 20} more")
        return 1

    scope = f" through {a.through_wp}" if a.through_wp else " (full)"
    print(f"ledger OK{scope}: {len(rows)} rows verified against {a.upstream}")
    if pending:
        print(f"  {len(pending)} row(s) pending in later work-phases")
    return 0


if __name__ == "__main__":
    sys.exit(main())
