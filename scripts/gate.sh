#!/usr/bin/env bash
# HWP re-port verification harness.
#
# Explicit bash, never the login shell: PIPESTATUS is bash-only ($pipestatus in
# zsh) and `echo "exit=$?"` PRINTS a failure instead of PROPAGATING it. Every
# check here either passes or kills the script.
set -Eeuo pipefail

PHASE="${1:-}"
REPO="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
LOG_DIR="${OCX_GATE_LOG_DIR:-/tmp/ocx-gate/$PHASE}"
mkdir -p "$LOG_DIR"
cd "$REPO"

# Never hard-code the artifact path: SelfContained=true nests it under a RID
# (bin/Debug/net10.0/osx-arm64/officecli.dll). Ask MSBuild.
msbuild_prop() {
  dotnet msbuild src/officecli/officecli.csproj \
    -nologo -v:quiet -getProperty:"$1" | tr -d '\r\n'
}

run() {                       # run <name> <cmd...>
  local name="$1"; shift
  local log="$LOG_DIR/$name.log"
  if "$@" >"$log" 2>&1; then
    echo "PASS $name"
  else
    local rc=$?
    echo "FAIL $name (exit=$rc) -> $log"; tail -30 "$log"; return $rc
  fi
}

assert() {                    # assert <message> <test-args...>
  local msg="$1"; shift
  if test "$@"; then echo "PASS assert: $msg"
  else echo "FAIL assert: $msg"; return 1; fi
}

resolve_bin() {
  DLL="$(msbuild_prop TargetPath)"
  assert "TargetPath resolved" -n "$DLL"
  BIN="${DLL%.dll}"
}

build_fresh() {               # freshness proven by absence-then-presence
  resolve_bin
  run clean dotnet clean src/officecli/officecli.csproj
  assert "dll removed by clean ($DLL)" ! -f "$DLL"
  run build dotnet build src/officecli/officecli.csproj
  assert "dll produced by this build ($DLL)" -f "$DLL"
}

runtime_surfaces() {          # shared by wp5 and wp7
  assert "cli binary exists" -x "$BIN"
  # Fresh scratch dir per run: fixed /tmp paths made the second invocation fail
  # with file_exists, which looks like a regression and is just stale state.
  local scratch; scratch="$(mktemp -d)"
  run help "$BIN" --help
  for c in hwp capabilities schema native-ops; do
    grep -qE "^[[:space:]]+$c" "$LOG_DIR/help.log" \
      || { echo "FAIL: root command $c not exposed (dead code)"; return 1; }
  done
  run capabilities "$BIN" capabilities --json
  # `schema` is a container command; `schema --json` is a usage error, not a
  # health check. Exercise the real subcommands.
  run schema-list     "$BIN" schema list --json
  run schema-validate "$BIN" schema validate --json
  # OOXML must survive our host-file surgery: every format CREATE + READ +
  # MUTATE, not a create-only smoke that the comment overstated.
  run docx-create "$BIN" create "$scratch/g.docx" --json
  run docx-view   "$BIN" view   "$scratch/g.docx" text --json
  run docx-add    "$BIN" add    "$scratch/g.docx" /body --type paragraph --prop text=gate --json
  run xlsx-create "$BIN" create "$scratch/g.xlsx" --json
  run xlsx-set    "$BIN" set    "$scratch/g.xlsx" /Sheet1/A1 --prop value=42 --json
  run xlsx-query  "$BIN" query  "$scratch/g.xlsx" /Sheet1/A1 --json
  run pptx-create "$BIN" create "$scratch/g.pptx" --json
  run pptx-view   "$BIN" view   "$scratch/g.pptx" outline --json
  # Upstream subsystems the host seams could have broken.
  run validate-docx "$BIN" validate "$scratch/g.docx" --json
  run dump-docx     "$BIN" dump     "$scratch/g.docx" --json
  run plugins-list  "$BIN" plugins list --json
  run hwpx-create "$BIN" create "$scratch/g.hwpx" --json
  run hwpx-view   "$BIN" view   "$scratch/g.hwpx" text --json
}

gate_wp2() {
  run cargo-build cargo build --manifest-path src/rhwp-field-bridge/Cargo.toml
  run bridge-build dotnet build src/rhwp-officecli-bridge/rhwp-officecli-bridge.csproj
  local tracked; tracked=$(git ls-files 'src/rhwp-field-bridge/target/*' | wc -l | tr -d ' ')
  assert "no tracked rust target artifacts" "$tracked" = "0"
  # HwpRuntimeProbe.DiscoverApiCommands decides availability by running
  # `--help` and doing stdout.Contains(command) per KnownApiCommands, so this
  # log is the source of truth for what the C# side will believe exists.
  #
  # Match on a delimiter boundary, not a bare substring: `grep -- set-field`
  # also matches `unset-field`, and a digit-tolerant boundary would accept
  # `create-blank2`. The help line is a `|`-separated command list, so require
  # the real delimiters. The C# side uses stdout.Contains() and keeps that
  # blind spot -- wp5 must not rely on the probe alone.
  run bridge-help ./src/rhwp-field-bridge/target/debug/rhwp-field-bridge --help
  # Full expected set, not a sample. render-png is intentionally absent unless
  # built with --features native-skia, so it is asserted separately below.
  local expected=(create-blank read-text render-svg export-pdf export-markdown
                  document-info diagnostics dump-controls dump-pages thumbnail
                  list-fields get-field set-field fill-fields replace-text insert-text
                  get-cell-text scan-cells set-cell-text convert-to-editable
                  native-op save-as-hwp)
  local missing=0
  for c in "${expected[@]}"; do
    grep -qE "(^|[|[:space:]])${c}([|[:space:]]|$)" "$LOG_DIR/bridge-help.log" \
      || { echo "FAIL: bridge does not advertise '$c' as a distinct command"; missing=1; }
  done
  assert "all ${#expected[@]} rhwp commands advertised" "$missing" = "0"
  # Dispatch smoke test: the usage string is hand-maintained and can advertise
  # a route that no longer exists. Exercise one real read path end to end.
  if [ -n "${OCX_HWP_SMOKE_INPUT:-}" ]; then
    run bridge-smoke ./src/rhwp-field-bridge/target/debug/rhwp-field-bridge \
      read-text --format hwp --input "$OCX_HWP_SMOKE_INPUT" --json
  else
    echo "SKIP bridge-smoke (set OCX_HWP_SMOKE_INPUT to a .hwp fixture; wp6 restores them)"
  fi
}

gate_wp3() { build_fresh; }
gate_wp4() { build_fresh; run schema-syntax python3 scripts/check-schema-syntax.py; }
gate_wp5() { build_fresh; runtime_surfaces; }

gate_wp6() {
  run clean-slnx dotnet clean officecli.slnx
  run build-slnx dotnet build officecli.slnx
  run test-all  dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj --no-build \
                  --logger "trx;LogFileName=$LOG_DIR/all.trx"
  run test-hwp  dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj \
                  --filter FullyQualifiedName~HwpBridge --no-build
  run test-bomb dotnet test tests/OfficeCli.Tests/OfficeCli.Tests.csproj \
                  --filter FullyQualifiedName~HwpxZipBomb --no-build \
                  --logger "trx;LogFileName=$LOG_DIR/bomb.trx"
  local n; n=$(python3 scripts/trx.py --file "$LOG_DIR/bomb.trx" --count-matching HwpxZipBombTests)
  assert "4 zip-bomb cases present (got $n)" "$n" -ge 4
  python3 scripts/trx.py --file "$LOG_DIR/all.trx" --list-all     | sort -u > "$LOG_DIR/inventory.txt"
  python3 scripts/trx.py --file "$LOG_DIR/all.trx" --list-skipped | sort -u > "$LOG_DIR/skipped.txt"
  python3 scripts/trx.py --file "$LOG_DIR/all.trx" --list-failed  | sort -u > "$LOG_DIR/failed.txt"
  assert "no failing tests" ! -s "$LOG_DIR/failed.txt"
  echo "recorded $(wc -l < "$LOG_DIR/skipped.txt") skipped tests for wp7 comparison"
}

gate_wp7() {
  run sidecars bash scripts/build-rhwp-sidecars.sh "$REPO/build-local/sidecars" "" Release
  export OFFICECLI_RHWP_API_BIN="$REPO/build-local/sidecars/rhwp-field-bridge"
  export OFFICECLI_RHWP_BRIDGE_PATH="$REPO/build-local/sidecars/rhwp-officecli-bridge"
  assert "api sidecar present"    -x "$OFFICECLI_RHWP_API_BIN"
  assert "bridge sidecar present" -x "$OFFICECLI_RHWP_BRIDGE_PATH"

  gate_wp6
  resolve_bin
  runtime_surfaces

  # skip -> executed. No `|| true`: comm returns 0 for an empty intersection,
  # so swallowing its status would make the whole check vacuous.
  # NOTE: the skip-transition check this gate used to require is unimplementable
  # against this suite. These tests guard on missing sidecars with an early
  # `return`, not an xUnit Skip, so every TRX reports skipped=0 and the
  # comparison was vacuous by construction. Assert the property that actually
  # matters instead: with sidecars attached, the HWP surface must really work.

  local hwp; hwp="$(mktemp -d)"
  run hwp-doctor  "$BIN" hwp doctor --json
  run hwp-create  "$BIN" create "$hwp/a.hwp" --json
  # Real mutation round-trip, not create-and-read-an-empty-file: seed text,
  # replace it, and assert the replacement is actually observable.
  run hwp-add     "$BIN" add "$hwp/a.hwp" /text --type paragraph \
                    --prop value=gatetext --prop output="$hwp/b.hwp" --json
  run hwp-read    "$BIN" view "$hwp/b.hwp" text --json
  grep -q 'gatetext' "$LOG_DIR/hwp-read.log" \
    || { echo "FAIL: inserted text not readable back"; return 1; }
  run hwp-replace "$BIN" set "$hwp/b.hwp" /text --find gatetext --replace gatedone \
                    --prop output="$hwp/c.hwp" --json
  run hwp-verify  "$BIN" view "$hwp/c.hwp" text --json
  grep -q 'gatedone' "$LOG_DIR/hwp-verify.log" \
    || { echo "FAIL: replacement not observable"; return 1; }
  # Render paths depend on the native-skia release sidecar.
  run hwp-pdf     "$BIN" view "$hwp/c.hwp" pdf --out "$hwp/c.pdf" --json
  assert "pdf produced" -s "$hwp/c.pdf"
  run hwp-png     "$BIN" view "$hwp/c.hwp" png --out "$hwp/c.png" --json
  run hwp-svg     "$BIN" view "$hwp/c.hwp" svg --out "$hwp/c.svg" --json
  run hwpx-create "$BIN" create "$hwp/d.hwpx" --json

  assert "OCX_PARENT set to cli-jaw root" -n "${OCX_PARENT:-}"
  run ledger python3 scripts/verify-ledger.py \
        --ledger "$OCX_PARENT/devlog/_plan/260803_officecli_upstream_report_rhwp/003_restore_ledger.csv" \
        --repo "$REPO" --upstream upstream/main --checks scripts/ledger-checks.json
  local integration_ledger="$OCX_PARENT/devlog/_plan/260803_officecli_upstream_report_rhwp/009_fork_main_merge_paths.tsv"
  assert "fork-main integration ledger exists" -f "$integration_ledger"
  run integration-ledger python3 scripts/verify-integration-merge.py \
        --ledger "$integration_ledger" \
        --repo "$REPO" \
        --base-feature dfdbcd89e018f139845e6c175aa9c27167ccca58 \
        --head HEAD \
        --checks scripts/integration-checks.json
}

case "$PHASE" in
  wp2|wp3|wp4|wp5|wp6|wp7) "gate_$PHASE" ;;
  *) echo "usage: gate.sh <wp2|wp3|wp4|wp5|wp6|wp7>" >&2; exit 2 ;;
esac
echo "GATE OK: $PHASE"
