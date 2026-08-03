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
import json
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
    ap.add_argument(
        "--backup",
        default="backup/pre-officecli-main-restore-20260626_112245",
        help="ref the ledger was derived from; used to prove the ledger is complete",
    )
    ap.add_argument(
        "--checks",
        help="check_id -> regex manifest (JSON). REQUIRED once any adapt row is in scope.",
    )
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

    checks_manifest: dict[str, str] = {}
    if a.checks:
        with open(a.checks, encoding="utf-8") as fh:
            checks_manifest = json.load(fh)

    def tracked(path: str) -> bool:
        # os.path.exists is not enough: an untracked or ignored file would
        # satisfy a "restored" row without ever entering the port.
        return git(a.repo, "ls-files", "--error-unmatch", "--", path).returncode == 0

    # Ledger completeness: the CSV must describe exactly the real diff, or a
    # truncated ledger silently stops being a completeness oracle.
    ns = git(a.repo, "diff", "--name-status", f"{a.upstream}..{a.backup}")
    if ns.returncode != 0:
        fail["cannot derive diff for completeness check"].append(ns.stderr.strip())
    else:
        real: set[tuple[str, str]] = set()
        for line in ns.stdout.splitlines():
            parts = line.split("\t")
            if len(parts) >= 2:
                real.add((parts[0], parts[1]))
        ledger_keys = {
            (r["diff_status"], r["path"])
            for r in rows
            if not r["diff_status"].endswith("-dest")
        }
        for missing in sorted(real - ledger_keys):
            fail["path in diff but NOT in ledger"].append(f"{missing[0]}\t{missing[1]}")
        for extra in sorted(ledger_keys - real):
            fail["ledger row not present in diff (stale)"].append(
                f"{extra[0]}\t{extra[1]}"
            )
        # Every rename must carry its synthetic destination row.
        dests = {r["path"] for r in rows if r["diff_status"].endswith("-dest")}
        for st, path in sorted(real):
            if st.startswith("R"):
                line = next(
                    (l for l in ns.stdout.splitlines() if l.startswith(f"{st}\t{path}\t")),
                    None,
                )
                dest = line.split("\t")[2] if line and len(line.split("\t")) > 2 else None
                if dest and dest not in dests:
                    fail["rename destination row missing from ledger"].append(dest)

    ALLOWED = {
        ("A", "restore"), ("A", "exclude"), ("A", "undecided"),
        ("M", "adapt"), ("M", "exclude"), ("M", "undecided"),
        # D,adapt is legitimate: the file is newer than our backup (so it reads
        # as "deleted" from the backup's perspective) but we still add a hook to
        # it. DocumentLimits.cs is the case -- upstream added it after the fork
        # point, and the zip-bomb recovery bounds belong there.
        ("D", "exclude"), ("D", "adapt"), ("R", "exclude"),
    }

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
        if (st[0], decision) not in ALLOWED:
            fail["unsupported diff_status/decision combination"].append(
                f"{st},{path},{decision}"
            )
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
            elif decision == "restore" and not tracked(path):
                fail["A,restore present but UNTRACKED"].append(path)
            elif decision == "exclude" and exists:
                fail["A,exclude unexpectedly present"].append(path)
        elif st.startswith("M"):
            if decision == "adapt":
                if not exists:
                    fail["M,adapt missing"].append(path)
                elif not check_id:
                    fail["M,adapt without check_id"].append(path)
                elif not checks_manifest:
                    fail["M,adapt in scope but no --checks manifest supplied"].append(
                        f"{path} ({check_id})"
                    )
                elif check_id not in checks_manifest:
                    fail["check_id not in manifest"].append(f"{path} ({check_id})")
                else:
                    # Prove the hook is actually present, not merely that the
                    # file exists. An untouched upstream file must FAIL here.
                    rc = subprocess.run(
                        ["grep", "-qE", checks_manifest[check_id], abs_path],
                        capture_output=True,
                    ).returncode
                    if rc != 0:
                        fail["adapt hook NOT found in file"].append(
                            f"{path} ({check_id})"
                        )
            elif decision == "exclude" and exists and not unchanged_vs_upstream(
                a.repo, a.upstream, path
            ):
                fail["M,exclude modified (must match upstream)"].append(path)
        elif st.startswith("D"):
            if decision == "adapt":
                if not exists:
                    fail["D,adapt missing"].append(path)
                elif not check_id:
                    fail["D,adapt without check_id"].append(path)
                elif not checks_manifest:
                    fail["D,adapt in scope but no --checks manifest supplied"].append(
                        f"{path} ({check_id})"
                    )
                elif check_id not in checks_manifest:
                    fail["check_id not in manifest"].append(f"{path} ({check_id})")
                else:
                    rc = subprocess.run(
                        ["grep", "-qE", checks_manifest[check_id], abs_path],
                        capture_output=True,
                    ).returncode
                    if rc != 0:
                        fail["adapt hook NOT found in file"].append(
                            f"{path} ({check_id})"
                        )
            # Upstream-newer file we did not touch: keep it, untouched.
            elif not exists:
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
