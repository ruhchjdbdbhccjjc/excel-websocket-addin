# Dual-Environment Manifest Collision & WEF Debug Lock Resolution

**Status**: Fixed (2026-11-13)  
**Area**: Office WEF Cache · Manifest GUID Segregation · WebSocket Dual-Client Race Conditions · Developer Debug Lifecycle

## Problem

1. When running both the local development version (`localhost:3000`) and the GitHub-hosted production version (`ruhchjdbdbhccjjc.github.io`) concurrently in Excel Desktop, identical `<Id>` GUIDs caused Office Web Extension Framework (WEF) cache corruption.
2. When both instances connected to Kanata WebSocket Server (`ws://127.0.0.1:8890`), targeted commands (`SendToClient("excel")`) starved the secondary client, while broadcast commands (`SendToChannel("excel_websocket_Offic_Add_in")`) executed twice simultaneously, causing workbook race conditions.
3. Sideloading via `HKCU\Software\Microsoft\Office\16.0\WEF\Developer` resulted in the error:
   > *"This add-in is no longer available: add-ins inserted during development are only available during debugging from Visual Studio. Please open your project in Visual Studio and re-run your application, or deploy your add-in into a valid catalog and re-insert."*
4. Clicking the `[DEV]` button without an active dev-server showed:
   > *"Sorry, we can't load the add-in. Please make sure you have network connectivity."*

## Root Cause (Why)

```text
FAILURE PATH (Shared GUID & Ephemeral Debug Registry)
  manifest.xml (Dev) & manifest.github.xml (Prod) share GUID c1b06178
         |
         +---------------------------------------+
         |                                       |
         v                                       v
  [ Office WEF Cache ]                    [ Kanata Core (:8890) ]
  - Pointers overwrite each other        - Dual clients named "excel"
  - Marked as "VS Debug Session"          - SendToClient picks 1st (starves 2nd)
  - When debugger closes:                 - SendToChannel runs both at once
    "Add-in is no longer available"       - Duplicate cell format & undo corrupt

SUCCESS PATH (Option 4 GUID Separation + Option 1 Automation Guard)
  Dev GUID: d2c17289-...0031              Prod GUID: c1b06178-...9930
  DisplayName: "[DEV] Excel WebSocket"    DisplayName: "Excel WebSocket"
         |                                       |
         v                                       v
  [ Office WEF: Independent Add-ins ]     [ Trusted Network Catalog ]
  - Independent cache partitions          - Permanent installation
  - Distinct ribbon buttons               - Zero debug-session expiration
         |
         v
  [ User Toggles Automation to Disabled ]
  - Instant socket close(1000)
  - Clears all reconnect backoff timers
  - Dims ribbon button (enabled: false)
  - Remaining instance has 100% exclusive access to Kanata (:8890)
```

## Fix (How)

1. **Option 4: Distinct Manifest Identity**:
   - Local Dev Manifest ([`manifest.xml`](../../manifest.xml)): Assigned dedicated GUID `d2c17289-5195-5b1f-902f-8bdeeef20031`, DisplayName `[DEV] Excel WebSocket`, and Ribbon label `[DEV] Excel WS`.
   - Production Manifest ([`manifest.github.xml`](../../manifest.github.xml)): Retained canonical GUID `c1b06178-4084-4a0e-8f1e-7acddec19930`.
   - Automated Build Pipeline ([`scripts/build-manifest.js`](../../scripts/build-manifest.js)): Injects production GUID and clean labels for GitHub Pages deployment automatically.
2. **Option 1: Master Automation Toggle**:
   - Added `automationEnabled?: boolean` in [`ConnectionSettings.ts`](../../src/websocket/storage/ConnectionSettings.ts#L10) and persisted to Office Settings.
   - Added `setAutomationEnabled(enabled: boolean)` in [`ConnectionManager.ts`](../../src/websocket/client/ConnectionManager.ts#L108-L135). Disabling immediately closes the WebSocket, cancels reconnection timers, and dims the ribbon.
   - In [`BackgroundService.ts`](../../src/websocket/service/BackgroundService.ts#L86-L100), `initialize()` skips auto-connecting if `automationEnabled === false`, and `handleRemoteMessage()` discards commands when disabled.
   - Built reactive SolidJS `.toggle-switch-btn` in [`SettingsTab.tsx`](../../src/taskpane/components/settings/SettingsTab.tsx) and modularized [`LogMonitorCard.tsx`](../../src/taskpane/components/settings/LogMonitorCard.tsx) to stay under the 250-line limit.
3. **WEF Developer Debug Lock Purge**:
   - Removed stale debug keys from `HKCU:\Software\Microsoft\Office\16.0\WEF\Developer`.
   - Purged stale cache directories matching `*c1b06178*` and `*d2c17289*` from `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.
   - Synchronized clean manifests to Trusted Network Catalog `C:\Users\Public\OfficeManifests\`.

## Files & Functions

| File | Functions / Role |
| :--- | :--- |
| [`ConnectionSettings.ts`](../../src/websocket/storage/ConnectionSettings.ts#L10) | Added `automationEnabled` schema and default |
| [`ConnectionManager.ts`](../../src/websocket/client/ConnectionManager.ts#L108-L135) | `isAutomationEnabled()`, `setAutomationEnabled()`, instant disconnect & timer disarm |
| [`BackgroundService.ts`](../../src/websocket/service/BackgroundService.ts#L86-L100) | `initialize()` automation guard & `handleRemoteMessage()` execution filter |
| [`SettingsTab.tsx`](../../src/taskpane/components/settings/SettingsTab.tsx#L35-L60) | UI toggle switch handler & card layout |
| [`LogMonitorCard.tsx`](../../src/taskpane/components/settings/LogMonitorCard.tsx) | Extracted log viewer component (< 150 lines) |
| [`scripts/build-manifest.js`](../../scripts/build-manifest.js) | Production GUID injection & label sanitization |
| [`manifest.xml`](../../manifest.xml) | Dedicated Dev GUID `d2c17289-5195-5b1f-902f-8bdeeef20031` |
