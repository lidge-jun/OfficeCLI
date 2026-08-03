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
  run help "$BIN" --help
  for c in hwp capabilities schema native-ops; do
    grep -qE "^[[:space:]]+$c" "$LOG_DIR/help.log" \
      || { echo "FAIL: root command $c not exposed (dead code)"; return 1; }
  done
  run capabilities "$BIN" capabilities --json
  run schema-cmd   "$BIN" schema --json
  # OOXML must survive our host-file surgery: create AND read AND mutate.
  run docx-create "$BIN" create /tmp/ocx-gate.docx --json
  run docx-view   "$BIN" view   /tmp/ocx-gate.docx text --json
  run xlsx-create "$BIN" create /tmp/ocx-gate.xlsx --json
  run xlsx-set    "$BIN" set    /tmp/ocx-gate.xlsx /Sheet1/A1 --prop value=42 --json
  run pptx-create "$BIN" create /tmp/ocx-gate.pptx --json
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
                  list-fields get-field set-field replace-text insert-text
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
  local prev="${OCX_WP6_SKIPPED:-}"
  assert "wp6 skip inventory provided" -n "$prev"
  assert "wp6 skip inventory exists"   -f "$prev"
  assert "wp6 skip inventory non-empty (else comparison is vacuous)" -s "$prev"
  comm -23 <(sort -u "$prev") <(sort -u "$LOG_DIR/inventory.txt") > "$LOG_DIR/vanished.txt"
  assert "no wp6-skipped test vanished from wp7 discovery" ! -s "$LOG_DIR/vanished.txt"
  comm -12 <(sort -u "$prev") <(sort -u "$LOG_DIR/skipped.txt") > "$LOG_DIR/still-skipped.txt"
  assert "sidecar-dependent tests now execute" ! -s "$LOG_DIR/still-skipped.txt"

  run hwp-doctor  "$BIN" hwp doctor --json
  run hwp-create  "$BIN" create /tmp/ocx-gate.hwp --json
  run hwp-read    "$BIN" view /tmp/ocx-gate.hwp text --json
  run hwpx-create "$BIN" create /tmp/ocx-gate.hwpx --json

  assert "OCX_PARENT set to cli-jaw root" -n "${OCX_PARENT:-}"
  run ledger python3 scripts/verify-ledger.py \
        --ledger "$OCX_PARENT/devlog/_plan/260803_officecli_upstream_report_rhwp/003_restore_ledger.csv" \
        --repo "$REPO" --upstream upstream/main
}

case "$PHASE" in
  wp2|wp3|wp4|wp5|wp6|wp7) "gate_$PHASE" ;;
  *) echo "usage: gate.sh <wp2|wp3|wp4|wp5|wp6|wp7>" >&2; exit 2 ;;
esac
echo "GATE OK: $PHASE"
