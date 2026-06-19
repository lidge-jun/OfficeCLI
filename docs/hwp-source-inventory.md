# HWP / rhwp Source Inventory

Provenance of the Korean HWP/HWPX support in this fork. (Created 2026-06-20 during the upstream
re-fork onto v1.0.115; see the cli-jaw devlog `260620_officecli_refork_and_rhwp_wiring`.)

## rhwp engine crate (pinned)

| Field | Value |
|---|---|
| Crate | `rhwp` (external) |
| Source | `https://github.com/edwardkim/rhwp.git` |
| **Pinned rev** | `1899ef9bc2dfd1c6c0c4d18b192d253a2d0a1fb5` (**tag v0.7.12**) |
| Pinned at | `src/rhwp-field-bridge/Cargo.toml` |
| Previous pin | `62a458aa…` (v0.7.10) — bumped 2026-06-20 |

**Why v0.7.12:** it is the earliest tag exposing the 5 newly-wired methods (see below).
v0.7.10/v0.7.11 lack them; v0.7.16 was rejected because its `insert_picture` signature change breaks
the bridge (E0061/E0308) — revisit when intentionally chasing the latest rhwp.

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

## Native ops wired in v0.7.12 (new, 2026-06-20)

Reflected from upstream rhwp `devel`/`v0.7.12`. Each is a `native-op` sub-operation (the `.NET`
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
