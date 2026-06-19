# Branch Strategy

## ⚠️ Upstream rewrites history — sync by RE-FORK, not rebase (2026-06-20)

`iOfficeAI/OfficeCLI` **force-rewrites its history** (verified: the old upstream tip `8735b01a` is NOT
an ancestor of the new `b5fad943` / v1.0.115). The legacy ff-only "Sync Procedure" below therefore
**breaks** — `git merge --ff-only upstream/main` is impossible, and a classic `git rebase upstream/main`
would replay ~3626 commits (≈3543 of them patch-duplicates) across ~982 overlapping files.

**Canonical sync = re-fork + replay the fork-unique delta:**
```bash
git fetch upstream
git branch backup/pre-refork-$(date +%Y%m%d) HEAD              # safety: preserve current tip
# find ONLY genuinely fork-unique commits (patch-id; excludes rewritten upstream dups):
git cherry upstream/main main | grep '^+' | awk '{print $2}' > /tmp/fork_delta.txt   # ~56 commits
git checkout -b refork/<upstream-version> upstream/main
git cherry-pick --empty=drop -x $(cat /tmp/fork_delta.txt)    # resolve the seams below
# run the verification gate, THEN (with owner approval) cut over main + push --force-with-lease
```

The fork delta is clean + additive (HWPX handlers, rhwp Rust+.NET bridge, CJK, fork installer/fixtures).
The 2026-06-20 re-fork onto **v1.0.115** replayed all **56** fork-unique commits with **zero conflicts**
and a green **269-test** suite. rhwp engine pinned at **v0.7.12** (`docs/hwp-source-inventory.md`).

**Integration seams** (resolve keeping BOTH sides — upstream plugin/PDF/minimal-doc **and** fork native
`.hwpx`/`.hwp` bridge + CJK): `CommandBuilder*.cs`, `DocumentHandlerFactory.cs`, `BlankDocCreator.cs`,
`ResidentServer.cs`, `WordHandler.Add.Text.cs`, `PowerPointHandler.*.cs`, `McpServer.cs`, `officecli.csproj`.

**Verification gate:** `git ls-files 'src/rhwp-field-bridge/target/*' | wc -l`==0 ·
`dotnet build officecli.slnx` · `cargo build --manifest-path src/rhwp-field-bridge/Cargo.toml` · `dotnet test`.

**Cutover is DESTRUCTIVE** (replace `main` + `push --force-with-lease`) → requires explicit owner
approval. `backup/pre-refork-*` is the recovery point.

> The "Branch Roles" / "Sync Procedure" below are **legacy** — valid only while upstream fast-forwards.

## Current State (2026-05-06, legacy)
- **upstream/main**: tracks `iOfficeAI/OfficeCLI` main. Pull-only, never push.
- **main**: local mirror of `upstream/main` plus compliance docs.
- **feat/hwpx**: active working branch on `origin`. Phase 36 HWP/HWPX evidence work (compatibility corpus, round-trip catalog, visual thresholds, provider matrix, release gate, safe-save) lives here.
- **docs/structure-init-2026-05-06**: short-lived doc branch that introduced `structure/`, fast-forwarded into `feat/hwpx`.

## Branch Roles
- **upstream/main**: read-only upstream reference.
- **main**: clean mirror; rebased from `upstream/main`.
- **feat/hwpx**: primary working branch. All cli-jaw HWP/HWPX changes land here.
- **feature/\*** / **docs/\***: optional short-lived branches. Merge back into `feat/hwpx` (typically `--ff-only`) and delete.

## Sync Procedure
```bash
git fetch upstream
git checkout main
git merge --ff-only upstream/main
git checkout feat/hwpx
git rebase main          # or: git merge main --no-ff, depending on history needs
```

## Rebase Conflict Notes

### 2026-05-14: upstream/main 1.0.91 rebase

`upstream/main` moved from the 1.0.68-era base to 1.0.91 and introduced core plugin/exporter work while `feat/hwpx` carried native HWPX and experimental HWP bridge work. Resolve these conflicts by preserving both sides:

- `BlankDocCreator.cs`: keep upstream `--minimal`/plugin-create support and keep native `.hwpx` creation before the plugin fallback. Unsupported-type text should list `.hwpx` plus plugin-served formats.
- `CommandBuilder.Import.cs`: keep upstream `--minimal` for DOCX and keep HWPX `--from-markdown`/`--align`; call `BlankDocCreator.Create(file, locale, minimal)` before optional HWPX markdown import.
- `CommandBuilder.cs`: register both upstream `BuildPluginsCommand` and fork `BuildCompareCommand`.
- `DocumentHandlerFactory.cs`: route `.hwpx` to `HwpxHandler`, keep `.hwp` bridge guidance, and fall back to upstream plugin handlers for unknown extensions.
- `CommandBuilder.View.cs` / `ResidentServer.cs`: keep upstream `pdf`/plugin forms support and HWP/HWPX modes (`forms`, `tables`, `markdown`, `objects`, `styles`, `fields`, `field`). Keep `BuildViewDescription()` when HWP bridge help is present.
- `WordHandler.Add.Text.cs`: keep upstream `sym=font:hex` handling so dump/batch round-trips do not duplicate symbol glyph text.

After resolving, continue the rebase and verify:

```bash
git ls-files 'src/rhwp-field-bridge/target/*' | wc -l  # must be 0
dotnet build officecli.slnx
cargo build --manifest-path src/rhwp-field-bridge/Cargo.toml
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --filter FullyQualifiedName~HwpBridge --no-build
dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --no-build
```

## Working with `structure/`
- See [`structure/INDEX.md`](structure/INDEX.md) for the doc map.
- Update `structure/*.md` whenever the corresponding source/schema/test surface changes (sync checklist in `structure/INDEX.md`).
- Doc-only changes should not modify `.cs`, `.ts`, `.js`, `.py`, or other source files.

## Rules
- Never push to `upstream`.
- All cli-jaw changes go on `feat/hwpx` (or short-lived branches that merge into it).
- HWP/HWPX claims must be evidence-gated per `docs/qa/phase-36-release-gate.md` — no DOCX parity language.
- Tag releases as `cjk-v{version}` on `feat/hwpx`.
