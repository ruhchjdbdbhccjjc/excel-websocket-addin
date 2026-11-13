# Dual-Environment Isolation & Master Automation Toggle

This feature enables seamless, collision-free coexistence of both the Local Development (`localhost:3000`) and GitHub Production (`ruhchjdbdbhccjjc.github.io`) add-in builds within Microsoft Excel Desktop.

```text
+---------------------------------------------------------------------------------------------------+
|               DUAL-ENVIRONMENT TOPOLOGY & CONFLICT-FREE ARBITRATION                               |
+---------------------------------------------------------------------------------------------------+

   [ LOCAL DEV ADD-IN ]                                  [ GITHUB PROD ADD-IN ]
   - GUID: d2c17289-...0031                               - GUID: c1b06178-...9930
   - Group: "[DEV] Kanata WebSocket"                      - Group: "Kanata WebSocket"
   - Button: "[DEV] Excel WS"                             - Button: "Excel WebSocket"
   - Origin: https://localhost:3000                       - Origin: https://ruhchjdbdbhccjjc.github.io
              |                                                     |
              v                                                     v
   [ Taskpane Settings: Automation Toggle ]               [ Taskpane Settings: Automation Toggle ]
   - Option: Enabled / Disabled                           - Option: Enabled / Disabled
   - Isolated localStorage partition                      - Isolated localStorage partition
              |                                                     |
              +--------------------------+--------------------------+
                                         |
                                         v
   +------------------------------------------------------------------------------------------------+
   |                       ARBITRATED KANATA WEBSOCKET SERVER (:8890)                                |
   |                                                                                                |
   | - When one instance is Disabled:                                                               |
   |     * Socket closed via ws.close(1000)                                                        |
   |     * Background reconnect retry timer disarmed                                                |
   |     * Ribbon control dimmed (Office.ribbon.requestUpdate)                                      |
   | - Active instance holds 100% EXCLUSIVE control:                                                |
   |     * SendToClient("excel") delivers without starvation                                        |
   |     * SendToChannel("excel_websocket_Offic_Add_in") runs once without race conditions          |
   +------------------------------------------------------------------------------------------------+
```

---

## 1. Capabilities & Workflow

### 1.1 Manifest Identity Segregation (Option 4)
- **Local Dev Manifest** ([`manifest.xml`](../../manifest.xml)):
  - Dedicated GUID: `d2c17289-5195-5b1f-902f-8bdeeef20031`
  - DisplayName: `[DEV] Excel WebSocket`
  - Ribbon Label: `[DEV] Excel WS`
  - Group Label: `[DEV] Kanata WebSocket`
- **Production Manifest** ([`manifest.github.xml`](../../manifest.github.xml)):
  - Canonical GUID: `c1b06178-4084-4a0e-8f1e-7acddec19930`
  - DisplayName: `Excel WebSocket`
  - Ribbon Label: `Excel WebSocket`
  - Group Label: `Kanata WebSocket`
- **Build Pipeline Invariant**:
  [`scripts/build-manifest.js`](../../scripts/build-manifest.js) deterministically substitutes the production GUID and production URLs into `dist/manifest.xml` and `manifest.github.xml` during `npm run build:gh-pages`.

### 1.2 Master Automation Toggle (Option 1)
- **Zero-Latency Toggle**: Toggling the switch takes effect in **< 10ms** without requiring an Excel reload or document restart.
- **Immediate Socket Termination**:
  - Sets `isManualDisconnect = true`.
  - Clears reconnection backoff timers (`clearReconnectTimer`).
  - Clears heartbeat timers (`clearHeartbeatTimer`).
  - Issues `ws.close(1000, 'Manual disconnect')`.
- **Persistent State**: Saved to both `Office.context.document.settings` and `localStorage.setItem('excel_addin_ws_automation_enabled', 'false')`.
- **Headless Execution Guard**:
  - [`BackgroundService.initialize`](../../src/websocket/service/BackgroundService.ts#L86-L100) respects `automationEnabled !== false`.
  - [`BackgroundService.handleRemoteMessage`](../../src/websocket/service/BackgroundService.ts#L180-L195) silently drops incoming remote actions when automation is disabled.

---

## 2. Key Functions Reference

| Function / Component | Location | Role |
| :--- | :--- | :--- |
| [`setAutomationEnabled`](../../src/websocket/client/ConnectionManager.ts#L125) | `ConnectionManager.ts` | Saves state, disconnects socket, and disarms retry timers |
| [`isAutomationEnabled`](../../src/websocket/client/ConnectionManager.ts#L119) | `ConnectionManager.ts` | Synchronous state check for automation activity |
| [`initialize`](../../src/websocket/service/BackgroundService.ts#L86) | `BackgroundService.ts` | Evaluates persistent setting before auto-connecting |
| [`handleRemoteMessage`](../../src/websocket/service/BackgroundService.ts#L180) | `BackgroundService.ts` | Guards remote execution against disabled automation |
| [`SettingsTab`](../../src/taskpane/components/settings/SettingsTab.tsx#L45) | `SettingsTab.tsx` | Reactive toggle switch UI pill |
| [`LogMonitorCard`](../../src/taskpane/components/settings/LogMonitorCard.tsx#L20) | `LogMonitorCard.tsx` | Modularized log console component |
