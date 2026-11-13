# Known Issues

## Active Known Issues

*(No active critical defects currently blocking production execution).*

---

## Recently Resolved (Last 30 Days)

### Dual-Environment Manifest Collision & WEF Debug Lock
**Status**: Fixed (2026-11-13)

When running both Local Dev and GitHub Production versions in Excel Desktop, identical GUIDs caused WEF cache corruption, Kanata client command starvation, and Visual Studio debug session expiration locks. Resolved via dedicated Dev GUID `d2c17289-...`, master automation toggle in settings, and WEF cache purge.

> Resolution: See [Dual-Environment Manifest Collision & WEF Debug Lock Fix](issues/dual_env_manifest_collision_and_wef_debug_lock_fix.md)

### Markdown Parser Fallback Heuristic Disabling & Tool Stream Hardening
**Status**: Fixed (2026-10-24)

Aggressive fallback heuristics in markdown parsers mutated streamed JSON envelopes and escaped characters, breaking tool dispatch loops. Resolved by disabling parser fallback heuristics and enforcing byte-exact deterministic streaming.

> Resolution: See [Markdown Parser Fix](issues/markdown_parser_fix.md)

### Dialog Lifecycle Reopen Reliability and Webpack HMR Crash
**Status**: Fixed (2026-10-22)

Webpack dev-server HMR multi-entrypoint collisions threw an Uncaught TypeError in `commands.js` that crashed the Office shared runtime, while concurrent taskpane opening and stale `currentDialog` references prevented the floating Command Palette dialog from reopening cleanly.

> Resolution: See [Dialog Lifecycle Reopen Reliability and Webpack HMR Crash Fix](issues/dialog_lifecycle_and_hmr_crash_fix.md)

### Spotlight Crash Residuals & Multi-Sheet Desync
**Status**: Fixed (2026-10-20)

Prior to the fix: (1) `f.includes('INDEX(')` caused false-positive deletion of user custom conditional formatting formulas; (2) worksheet switching rendered crosshairs on the new sheet without clearing the previous sheet; (3) fast arrow navigation dropped events when an IPC sync was in-flight, leaving crosshairs stuck; (4) startup detox was bypassed on reboot after crashes; and (5) multi-sheet purge caused high IPC round-trip latency.

> Resolution: See [Spotlight Crash and Sheet Desync Fix](issues/spotlight_crash_and_sheet_desync_fix.md)

### Spotlight Formula Precedence and Border Box Defect
**Status**: Fixed (2026-10-18)

In Excel Spotlight, anchor cell rules were placed after the master crosshair formula in the conditional formatting collection, causing the custom accent color to be overridden by the master spotlight fill. Additionally, single-cell anchor rules using `mode: 'row_only'` lacked left/right borders, header-based peer matching defaulted to Column A, and numeric inputs were not clamped against Excel bounds.

> Resolution: See [Spotlight Formula Priority and Border Fix](issues/spotlight_formula_priority_and_border_fix.md)

### Kanata Channel Subscription Mismatch & Missing Imports
**Status**: Fixed (2026-10-16)

Kanata server filtered incoming channel messages and reported "Sending message to 0 clients on channel: excel_websocket_Offic_Add_in" due to mismatched subscription arrays and absence of explicit subscribe handshake. Additionally, `MessageHandler.ts` threw `ReferenceError: KanataProtocol is not defined` when receiving frames due to missing imports.

> Resolution: See [Kanata Channel Subscription and Import Fix](issues/kanata_channel_subscription_and_import_fix.md)

### Missing clientType Field in Register Message
**Status**: Fixed (2026-10-15)

The WebSocket register message was missing the required `clientType` field, causing the Kanata server to reject the connection with a parsing error.

> Resolution: See [Register Message Fix](issues/register_message_fix.md)

### Async Function Not Awaited in updateConnectionUI
**Status**: Fixed (2026-10-14)

The `updateConnectionUI()` function was calling `connectionManager.getSettings()` without `await`, causing a "Cannot read properties of undefined (reading 'join')" error when the Connect button was clicked.

> Resolution: See [UI Update Fix](issues/ui_update_fix.md)

### Log-Level Functionality Removed
**Status**: Fixed (2026-10-14)

Log-level functionality was previously removed from the codebase, causing compilation errors. The Logger class was recreated and integrated throughout the WebSocket components.

> Resolution: See [Logger Integration Fix](issues/logger_integration_fix.md)

### Button Click Event Listeners Not Attaching
**Status**: Fixed (2026-10-12)

Event listeners for WebSocket control buttons were not attaching properly due to timing issues with DOM readiness and connection manager initialization.

> Resolution: See [Event Listener Fix](issues/event_listener_fix.md)

### TypeScript Compilation Errors
**Status**: Fixed (2026-10-10)

Various TypeScript compilation errors occurred due to missing imports and incorrect type usage after log-level removal.

> Resolution: See [TypeScript Compilation Fix](issues/typescript_compilation_fix.md)
