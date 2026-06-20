# HWP / rhwp Source Inventory

Provenance of the Korean HWP/HWPX support in this fork. (Created 2026-06-20 during the upstream
re-fork onto v1.0.115; see the cli-jaw devlog `260620_officecli_refork_and_rhwp_wiring`.)

## rhwp engine crate (pinned)

| Field | Value |
|---|---|
| Crate | `rhwp` (external) |
| Source | `https://github.com/edwardkim/rhwp.git` |
| **Pinned rev** | `de02159ab4d2c5d165d6e25568bad3f8af5ef6cb` (**tag v0.7.16**) |
| Pinned at | `src/rhwp-field-bridge/Cargo.toml` |
| Previous pin | `1899ef9b…` (v0.7.12) — bumped 2026-06-20 |

**Why v0.7.16:** it is the latest release tag confirmed on 2026-06-20 and keeps the
5 newly-wired methods introduced by v0.7.12. The v0.7.16 `insert_picture_native`
signature adds `cell_path` plus optional paper offsets; OfficeCLI preserves the old
`insert-picture` behavior by passing an empty cell path and `None` offsets from
`src/rhwp-field-bridge/src/ops_native_support.rs`. v0.7.10/v0.7.11 lack the
newly-wired methods.

## Bridge layers (call chain)

```
OfficeCLI (.NET handlers, src/officecli/Handlers/Hwp/**)
  → rhwp-officecli-bridge (C#, src/rhwp-officecli-bridge/Program.cs)   [sidecar 1: dispatcher]
    → rhwp-field-bridge (Rust, src/rhwp-field-bridge/src/**)           [sidecar 2: engine wrapper]
      → rhwp crate (rhwp::wasm_api::HwpDocument)                       [the HWP/HWPX engine]
```

Build: `scripts/build-rhwp-sidecars.sh` (local-RID only). Runtime discovery: `HwpRuntimeProbe.cs`
(env `OFFICECLI_RHWP_BIN` / `OFFICECLI_RHWP_API_BIN`, then PATH/app-dir). Capability gating is
command-driven (`NativeOpAvailable` = bridge advertises `native-op`).

## Native ops wired from v0.7.12+ (new, 2026-06-20)

Reflected from upstream rhwp `devel`/`v0.7.12` and retained in `v0.7.16`. Each is a `native-op` sub-operation (the `.NET`
side forwards `--prop key=value` generically, so no .NET change was required):

| Op string | rhwp method | Args | Family file |
|---|---|---|---|
| `search-all-text` | `search_all_text_native` | `--query`, `--case-sensitive`(def false), `--include-cells`(def true) | `ops_native_text.rs` |
| `insert-new-number` | `insert_new_number_native` | `--section --paragraph --offset --start-num` | `ops_native_text.rs` |
| `get-page-overlay-images` | `get_page_overlay_images_native` | `--page-num` | `ops_native_objects.rs` |
| `get-hf-picture-properties` | `get_header_footer_picture_properties_native` | `--section --outer-para --outer-control --inner-para --inner-control` | `ops_native_header_footer.rs` |
| `set-hf-picture-properties` | `set_header_footer_picture_properties_native` | …same 5 indices + `--props-json` | `ops_native_header_footer.rs` |

Smoke-verified (`search-all-text`, `get-page-overlay-images` return clean JSON); all 5 compile and
dispatch. The pre-existing ~64 native ops are unchanged.
