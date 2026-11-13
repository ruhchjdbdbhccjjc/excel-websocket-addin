# Excel WebSocket Add-in System Design Document

**Document Version**: 1.0  
**Project**: `Excel WebSocket` (`ExcelShortcut`)  
**Target Host**: Microsoft Excel Desktop (Windows / Office 365 / Excel 2016+)  
**Integration Target**: Kanata Keyboard Controller (`ws://localhost:8890`)

---

## 1. Executive Summary & Design Goals

The **Excel WebSocket** add-in bridges Microsoft Excel with the Kanata keyboard layout manager. It transforms Excel from an isolated spreadsheet application into a fully responsive automation endpoint capable of receiving real-time keyboard triggers, executing headless spreadsheet manipulations, and exposing connection feedback directly on the native Office Ribbon.

### Core Design Principles
1. **Single-Context Persistence (Shared Runtime)**: Background services and UI must run in a unified, persistent memory space to prevent state desync and eliminate background-load limits.
2. **Minimal Ribbon Footprint**: A single, clean Home tab button (`Excel WebSocket`) serves as both an interactive control panel launcher and a dynamic visual indicator of Kanata's connection status.
3. **Headless-First Operation**: The add-in must start automatically when Excel opens (`Office.StartupBehavior.load`) and process remote commands without requiring the taskpane sidebar to remain visible.
4. **Infinite Reconnect Resilience**: The WebSocket engine must seamlessly withstand Kanata crashes, live reloads (`lrld`), and process restarts via exponential backoff without requiring user intervention.

---

## 2. System Architecture & Runtime Model

```
+-----------------------------------------------------------------------------------+
|                            Microsoft Excel Desktop Host                           |
|                                                                                   |
|  +-----------------------------------------------------------------------------+  |
|  |                           Excel Home Tab Ribbon                             |  |
|  |                     Group: [Excel WebSocket]                                |  |
|  |                     Control: [btnExcelWebSocket]                            |  |
|  |                     State: (enabled: true = Online | false = Offline)        |  |
|  +---------------------------------------+-------------------------------------+  |
|                                          ^                                        |
|                                          | Office.ribbon.requestUpdate            |
|  +---------------------------------------+-------------------------------------+  |
|  |                       Shared JavaScript Runtime                             |  |
|  |                       (taskpane.html + all bundled chunks)                  |  |
|  |                                                                             |  |
|  |   +---------------------------------------------------------------------+   |  |
|  |   |             BackgroundService (Singleton Headless Daemon)           |   |  |
|  |   |   - Auto-boots via Office.onReady on document open                  |   |  |
|  |   |   - Evaluates remote JSON-RPC / S-expression commands headlessly    |   |  |
|  |   |   - Issues ribbon state updates upon connection lifecycle changes   |   |  |
|  |   +---------------------------------+-----------------------------------+   |  |
|  |                                     |                                       |  |
|  |   +---------------------------------v-----------------------------------+   |  |
|  |   |           ConnectionManager (Multi-Listener Event Hub)              |   |  |
|  |   |   - Maintains array-based observer registries (stateChange, etc.)   |   |  |
|  |   |   - Caches settings & state in Office Settings Storage              |   |  |
|  |   +-------------------+-----------------------------+-------------------+   |  |
|  |                       |                             |                       |  |
|  |                       v                             v                       |  |
|  |   +-----------------------+     +---------------------------------------+   |  |
|  |   |    WebSocketClient    |     |          Taskpane View Controller     |   |  |
|  |   | - Auto-reconnect loop |     | - Settings UI (Port, Client, Channel) |   |  |
|  |   | - ws://127.0.0.1:8890 |     | - Live Diagnostic Event Console       |   |  |
|  |   +-----------+-----------+     +---------------------------------------+   |  |
|  +---------------|-------------------------------------------------------------+  |
+------------------|----------------------------------------------------------------+
                   | Bidirectional TCP WebSocket
                   v
+-----------------------------------------------------------------------------------+
|                        Kanata WebSocket Server (:8890)                            |
|                                                                                   |
|  - Channel: "excel-addin-websocket"                                               |
|  - Client Type: "office-addin", Client Name: "excel"                              |
|  - Bidirectional triggers (`ws-send-channel`, `ws-send-name`, HTTP /api/tcp)      |
+-----------------------------------------------------------------------------------+
```

---

## 3. Detailed Component Design

### 3.1. Shared Runtime & Manifest Pipeline

Standard Office Add-ins run commands and taskpanes in isolated browser processes. To enable dynamic ribbon updates and headless background persistence, this add-in enforces the **Shared JavaScript Runtime**:

- **Manifest XML Declaration** (`manifest.xml`):
  ```xml
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="SharedRuntime" MinVersion="1.1"/>
    </Sets>
  </Requirements>
  ...
  <Host xsi:type="Workbook">
    <Runtimes>
      <Runtime resid="Taskpane.Url" lifetime="long" />
    </Runtimes>
    <DesktopFormFactor>
      <FunctionFile resid="Taskpane.Url"/>
  ```
- **Webpack Bundling Strategy** (`webpack.config.js`):
  Both `taskpane` and `commands` entry points are combined into `taskpane.html`:
  ```javascript
  new HtmlWebpackPlugin({
    filename: "taskpane.html",
    template: "./src/taskpane/taskpane.html",
    chunks: ["polyfill", "taskpane", "commands"],
  })
  ```
- **Automated Manifest Converter** (`scripts/convert-manifest.js`):
  Maintains the Teams Toolkit JSON format (`appPackage/manifest.json`) as the single source of truth for metadata while automatically transforming it into schema-compliant `manifest.xml` during build and debug cycles.

---

### 3.2. Ribbon UX & State Reactivity

Rather than cluttering the Home tab with multiple action buttons, the interface employs a **single multifunctional control**:

```
+----------------------------+
|     [K] Excel WebSocket    |  <-- In Full Color (Connected)
+----------------------------+
+----------------------------+
|     [K] Excel WebSocket    |  <-- Dimmed / Grayed-out (Disconnected)
+----------------------------+
```

1. **Control Definition**:
   - Group: `ExcelWebSocketGroup`
   - Control ID: `btnExcelWebSocket`
   - Label: `Excel WebSocket`
   - Action: `ShowTaskpane` (Opens connection settings and live console).
2. **Reactivity Mechanism**:
   ```typescript
   await Office.ribbon.requestUpdate({
     tabs: [{
       id: 'TabHome',
       groups: [{
         id: 'ExcelWebSocketGroup',
         controls: [{
           id: 'btnExcelWebSocket',
           enabled: isConnected
         }]
       }]
     }]
   });
   ```
3. **Behavior Matrix**:
   - **Kanata Online (`connected`)**: Button is colorful and enabled. Clicking opens the diagnostic sidebar.
   - **Kanata Offline (`disconnected` / `error`)**: Button dims immediately (`enabled: false`), signaling that background automation is offline without interrupting normal spreadsheet work.

---

### 3.3. Multi-Listener Event Hub (`ConnectionManager.ts`)

To avoid event listener overwriting between the background daemon and the taskpane UI, `ConnectionManager` implements an array-based observer registry:

```typescript
// Prevents taskpane initialization from clobbering BackgroundService
private eventHandlers: { [K in keyof ConnectionManagerEvents]?: Array<ConnectionManagerEvents[K]> } = {};

on<K extends keyof ConnectionManagerEvents>(event: K, handler: ConnectionManagerEvents[K]): void {
  if (!this.eventHandlers[event]) {
    this.eventHandlers[event] = [];
  }
  (this.eventHandlers[event] as Array<ConnectionManagerEvents[K]>).push(handler);
}
```

When a connection transition occurs (`setState`), all subscribers (`BackgroundService` for ribbon state, `taskpane.ts` for DOM indicators) are notified in parallel within isolated `try/catch` blocks.

---

### 3.4. Resilient WebSocket Engine (`WebSocketClient.ts`)

Kanata is frequently restarted during development or configuration switching (`lrld`). The client incorporates an unyielding reconnection loop:

- **Backoff Algorithm**:
  $$\text{delay} = \min(\text{initialDelay} \times \text{factor}^{\text{attempt}}, \text{maxDelay})$$
  - Initial delay: 1,000 ms
  - Exponential factor: 2.0
  - Maximum ceiling: 30,000 ms
  - `maxReconnectAttempts`: `Infinity`
- **Initial Connect Safety**: If the initial connection attempt fails because Kanata is not yet running, the client does not crash or throw unhandled exceptions; it transitions immediately into the background retry loop.

---

### 3.5. Network Security & Cross-Origin WSS Transport

While native desktop environments (such as Win32 WebView2) have unrestricted loopback access to `ws://127.0.0.1:8890`, cloud-hosted web runtimes (such as Excel on the Web at `https://excel.office.com`) execute inside a sandboxed browser `<iframe>` guarded by W3C security policies:

```text
+---------------------------------------------------------------------------------+
|               CROSS-PLATFORM WEBSOCKET TRANSPORT TOPOLOGY                       |
+---------------------------------------------------------------------------------+
  Excel Desktop (Win32 WebView2)  ---> ws://127.0.0.1:8890  ---> Kanata Core (:8890)
  Excel on the Web (Office Cloud) ---> wss://localhost:8891 ---> WSS Proxy (:8891)
                                                                       |
                                                                       v
                                                              Kanata Core (:8890)
```

1. **Dual Transport Ports**:
   - **Port `8890` (`ws://`)**: Default plaintext TCP WebSocket server hosted directly by `kanata.exe`. Used by desktop runtimes and browser extensions.
   - **Port `8891` (`wss://`)**: Secure TLS reverse proxy hosted via [`scripts/wss-proxy.js`](../scripts/wss-proxy.js). Terminated with trusted local certificates (`~/.office-addin-dev-certs/localhost.crt`) and forwarded to port `8890`.
2. **Browser Sandbox Mitigation**:
   - **Mixed Active Content**: [`WebSocketClient`](../src/websocket/client/WebSocketClient.ts#L61) detects secure origins or port `8891` targets, automatically requesting `wss://` instead of `ws://` to satisfy browser encryption constraints.
   - **Private Network Access (PNA)**: Proxy and endpoints emit `Access-Control-Allow-Private-Network: true` to prevent browser preflight rejections.
3. **Future Native Termination Blueprint**:
   - Eliminates external node scripts by integrating `tokio-rustls` directly into Kanata's binary, binding dual listeners (`:8890` plaintext and `:8891` TLS) automatically when local development certificates are detected.

> Deep Dive: See [Cross-App WSS & Browser Security Sandbox Guide](features/cross_app_wss_and_browser_sandbox_guide.md)

---

### 3.6. Dual-Environment Identity & Master Automation Toggle

To support local developer debugging alongside production users on the same machine without collisions, the system implements dual-manifest isolation and a master execution switch:

```text
+---------------------------------------------------------------------------------+
|               DUAL-ENVIRONMENT ARBITRATION & AUTOMATION TOGGLE                 |
+---------------------------------------------------------------------------------+
  Local Dev:  GUID d2c17289-...0031 | [DEV] Excel WS    | https://localhost:3000
  Production: GUID c1b06178-...9930 | Excel WebSocket   | https://...github.io
                                     |
                                     v
                 [ Master Toggle: Disabled / Enabled ]
                 - Disabled: Socket closed, retry canceled, ribbon dimmed
                 - Enabled:  Socket connected, full Kanata IPC active
```

1. **Manifest Segregation**:
   - Distinct GUIDs prevent Office Web Extension Framework (WEF) cache collisions.
   - Distinct ribbon labels (`[DEV] Excel WS` vs `Excel WebSocket`) ensure visual clarity.
2. **Master Automation Guard**:
   - [`ConnectionManager.setAutomationEnabled`](../src/websocket/client/ConnectionManager.ts#L125) provides instant (< 10ms) socket disconnect and cancelation of backoff retry timers.
   - [`BackgroundService`](../src/websocket/service/BackgroundService.ts#L86) guards `initialize()` and `handleRemoteMessage()` so dormant instances generate zero network traffic and execute zero commands.

> Deep Dive: See [Dual-Environment Isolation & Master Automation Toggle](features/dual_env_isolation_and_automation_toggle.md)

---

## 4. Headless Command Execution Pipeline

When Kanata transmits an action message over WebSocket, the add-in executes it directly via the Excel JavaScript Object Model:

```
[Kanata ws-send-name "excel"] 
        |
        v JSON Payload: { "type": "command", "data": { "action": "formatCell", "range": "A1:D1", "color": "#E0FFE0" } }
[WebSocketClient.onMessage]
        |
[BackgroundService.handleRemoteMessage]
        |
        +--> Excel.run(async (context) => {
               const sheet = context.workbook.worksheets.getActiveWorksheet();
               const range = sheet.getRange(cmd.range);
               range.format.fill.color = cmd.color;
               await context.sync();
             });
```

Because `BackgroundService` runs in the Shared Runtime, it possesses direct access to `Excel.run` without needing the user to open or focus the taskpane.

---

## 5. Debugging & Deployment Architecture

### 5.1. Desktop XML Sideload Pipeline
Office Desktop on Windows does not support local Teams JSON sideloading for ribbon commands. The development pipeline is configured to use the XML path by default:

- **`npm start` / `npm run start:desktop:excel`**:
  Executes `npm run convert-manifest` $\rightarrow$ copies `manifest.xml` $\rightarrow$ runs `office-addin-debugging start manifest.xml desktop --app excel`.
- **VS Code F5 Integration** (`.vscode/launch.json`):
  `Excel Desktop (Edge Chromium)` is the primary compound configuration. Pressing F5 starts the dev-server and attaches Microsoft Edge Chromium DevTools to the shared runtime webview.

### 5.2. Trusted Network Catalog Synchronization
To ensure seamless sideloading without certificate or administrative blocks, `manifest.xml` is published to a local trusted network share:
- Physical path: `C:\Users\Public\OfficeManifests\manifest.xml`
- UNC Share path: `\\DESKTOP-TP761MQ\Users\Public\OfficeManifests`

---

## 6. Multi-Language Split-File Architecture (i18n)

The add-in implements a strict split-file internationalization architecture:

```text
+-------------------------------------------------------------+
|                      i18n SUBSYSTEM                         |
+-------------------------------------------------------------+
  [src/i18n/types.ts]  <-- Strongly typed TranslationSchema
          ^
          +-------------------------------+
          |                               |
  [locales/zh-CN.ts]              [locales/en-US.ts]
  (Chinese Dictionary)            (English Dictionary)
          |                               |
          +---------------+---------------+
                          |
                          v
                 [src/i18n/index.ts]
                 - useI18n() SolidJS Hook
                 - Multi-tier detection:
                   1. localStorage ('excel_addin_locale')
                   2. Office.context.displayLanguage
                   3. navigator.language
                   4. Fallback: en-US
```

### Design Benefits:
- **Compile-Time Key Validation**: Any missing translation key in any locale triggers a TypeScript compilation error.
- **Granular Locales**: Locales are isolated in separate files (<70 lines each) avoiding giant monolithic JSON files.
- **Instant Reactive Switching**: Powered by SolidJS signals, switching locale updates the entire taskpane UI instantaneously without reloading the page or losing WebSocket state.

---

## 7. Zero-Pollution Crash Recovery & Startup Detox Protocol

Conditional formatting is Excel's native presentation layer. However, if Excel terminates unexpectedly, conditional formatting rules persist in the `.xlsx` file. To ensure 0 permanent document pollution:

```text
+-------------------------------------------------------------+
|               CRASH RECOVERY & DETOX PROTOCOL               |
+-------------------------------------------------------------+

  1. Lifecycle Unload Handlers:
     - window.addEventListener('beforeunload', handleUnload)
     - window.addEventListener('pagehide', handleUnload)
     - Disarms active crosshairs before runtime shutdown.

  2. Startup Auto-Detox (BackgroundService.initialize):
     - Unconditionally executes purgeAllOrphanSpotlightsAcrossWorkbook().
     - Batched 3-phase atomic IPC sync across all sheets in workbook.
     - Strictly isolates spotlight formulas (isSpotlightFormula) without
       touching user conditional formats.
     - If spotlight enabled: re-applies fresh crosshair to active cell only.

  3. Cross-Sheet Ghost Isolation:
     - On selection change, automatically cleans previous worksheet
       (context.workbook.worksheets.getItemOrNullObject(prevSheetName)).

  4. Trailing Execution Slot (pendingSelection):
     - Drains queued selections when IPC sync finishes, preventing
       frozen crosshairs during rapid arrow-key navigation.
```

---

## 8. Unified Diagnostic Logging Hub

- **1000-Entry Ring Buffer**: Fixed in-memory queue capturing timestamped, scoped entries (`SYSTEM`, `WS-IN`, `WS-OUT`, `EXCEL-CMD`, `CF-DETOX`).
- **Global Error Interception**: Automatically catches unhandled promise rejections, `window.onerror`, and overrides `console.log/warn/error` so all background events appear in the export log.
- **Fingerprinted Diagnostic Export**: One-click download of `.log` files with complete platform, host build, port, channels, and stack traces.

---

## 9. Directory & Code Reference

| File | Purpose |
| :--- | :--- |
| [`scripts/convert-manifest.js`](../scripts/convert-manifest.js) | Manifest generation script with Shared Runtime & Ribbon definitions |
| [`scripts/build-manifest.js`](../scripts/build-manifest.js) | Build script generating production GitHub Pages manifest (`manifest.github.xml`) |
| [`scripts/wss-proxy.js`](../scripts/wss-proxy.js) | Zero-dependency local TLS reverse proxy forwarding `wss://localhost:8891` to `ws://127.0.0.1:8890` |
| [`docs/features/cross_app_wss_and_browser_sandbox_guide.md`](features/cross_app_wss_and_browser_sandbox_guide.md) | Architectural guide on browser sandbox, Mixed Content, PNA, and native Kanata WSS |
| [`docs/features/dual_env_isolation_and_automation_toggle.md`](features/dual_env_isolation_and_automation_toggle.md) | Architectural guide on dual-environment isolation and master automation toggle |
| [`manifest.xml`](../manifest.xml) | Dedicated local dev Office XML manifest sideloaded into Excel |
| [`manifest.github.xml`](../manifest.github.xml) | Production Office XML manifest hosted on GitHub Pages CDN |
| [`src/websocket/service/BackgroundService.ts`](../src/websocket/service/BackgroundService.ts) | Headless daemon managing startup behavior, command execution, and ribbon updates |
| [`src/websocket/service/ExcelSpotlightService.ts`](../src/websocket/service/ExcelSpotlightService.ts) | Zero-destruction spotlight engine with adaptive selection and crash detox |
| [`src/websocket/service/ExcelPositionSlotManager.ts`](../src/websocket/service/ExcelPositionSlotManager.ts) | Fast position slot memory and safe selection tracking |
| [`src/websocket/logging/Logger.ts`](../src/websocket/logging/Logger.ts) | Singleton ring-buffer logging hub with global console and error hooks |
| [`src/i18n/index.ts`](../src/i18n/index.ts) | Reactive i18n hook, multi-tier locale detection, and storage persistence |
| [`src/websocket/client/ConnectionManager.ts`](../src/websocket/client/ConnectionManager.ts) | Multi-listener connection manager and settings storage bridge |
| [`src/websocket/client/WebSocketClient.ts`](../src/websocket/client/WebSocketClient.ts) | Resilient WebSocket transport client with infinite auto-reconnection |
| [`src/taskpane/components/settings/SettingsTab.tsx`](../src/taskpane/components/settings/SettingsTab.tsx) | Connection settings UI with reactive master automation switch (< 175 lines) |
| [`src/taskpane/components/settings/LogMonitorCard.tsx`](../src/taskpane/components/settings/LogMonitorCard.tsx) | Modular diagnostic log console with export and copy (< 150 lines) |
| [`src/taskpane/theme/ThemeManager.ts`](../src/taskpane/theme/ThemeManager.ts) | Theme manager detecting Excel theme and system dark mode |
| [`src/data/CommandSearchEngine.ts`](../src/data/CommandSearchEngine.ts) | High-speed bilingual search engine with Pinyin and acronym indexing |
| [`src/data/excel_mso_bilingual.json`](../src/data/excel_mso_bilingual.json) | 2,243 official Excel commands with bilingual labels and explanations |
| [`src/taskpane/components/CommandPalette.tsx`](../src/taskpane/components/CommandPalette.tsx) | Instant command palette with live supertip preview and keyboard navigation |
| [`src/taskpane/App.tsx`](../src/taskpane/App.tsx) | Main SolidJS shell with header language pill, theme pill, and segmented navigation |
| [`webpack.config.js`](../webpack.config.js) | Webpack configuration bundling taskpane and commands into a shared runtime |

---

## 10. Official Bilingual Command Hub & Pinyin Search Engine

### 10.1. Ground Truth Extraction via Native Excel COM
Rather than relying on generic translation files (such as `CHINESE (SIMPLIFIED).tbx`), all 2,243 official command names (`GetLabelMso`) and in-depth explanations (`GetSupertipMso`) are extracted directly from Excel's installed MUI resource libraries (`1033` English & `2052` Chinese).

```text
+---------------------------------------------------------------------------------+
|                       BILINGUAL SEARCH ENGINE PIPELINE                          |
|                                                                                 |
|   1. Query: "djcg" / "冻结" / "fp" / "freeze"                                   |
|   2. Decoupled LRU Query Cache (Keyed purely by query string):                  |
|        - Pre-filters top candidates with base scoring (0.01ms latency)          |
|        - Never evicted or invalidated by executing recent commands              |
|   3. Multi-tier Matching:                                                       |
|        - Exact ID / Label match: +3500 - +4000                                  |
|        - Exact Pinyin initials ("djcg" -> 冻结窗格): +3200                      |
|        - Exact English acronym ("fp" -> Freeze Panes): +2600                    |
|        - Multi-word prefix cohesion: +600                                       |
|        - Actionable control bonus (buttons): +50                                |
|   4. Dynamic MRU Recent Boost (+400 to +25):                                    |
|        - Applied dynamically on top cached candidates                           |
|        - Only verified native MSO commands are stored (custom typos excluded)   |
+---------------------------------------------------------------------------------+
```

### 10.2. Dual-Layer Hotkey Architecture (Ctrl+Shift+P & Ctrl+Shift+K)
- **Primary Hotkey (`Ctrl + Shift + P`)**:
  - Matches the universal muscle memory of VS Code, DevTools, and modern editors.
  - Intercepted by Kanata at the driver level when Excel is active, swallowing the key event and preventing Excel's native Format Cells (Font) modal dialog.
  - Automatically sends `{"command": "open_command_palette"}` over WebSocket `:8890`.
- **Default Alias Hotkey (`Ctrl + Shift + K`)**:
  - Mapped identically to `open_command_palette` for users who prefer to retain Excel's native font shortcut without conflicts.
- **In-Taskpane DOM Fallback**:
  - `window.addEventListener('keydown', ...)` active in the add-in webview catches both `Ctrl+Shift+P` and `Ctrl+Shift+K` directly when focused.
  - Automatically calls `Office.addin.showAsTaskpane()` to bring the palette to the foreground if closed or minimized in Shared Runtime.

---

## 11. Theme & Language Synchronization

### 11.1. Automatic Theme Synchronization
- **Dynamic Luminance Calculation**: Rather than fragile hardcoded hex string matches, theme background colors are evaluated using ITU-R BT.601 perceived luminance `(0.299*R + 0.587*G + 0.114*B) < 128` to classify any Microsoft 365 dark variant (`#262626`, `#201F1E`, `#1E1E1E`, `#333333`).
- **Safe Deferred Office Lifecycle Listener**: Theme change handlers (`addHandlerAsync`) are registered strictly inside `Office.onReady()` to eliminate race conditions before Office.js initialization.
- **OS Theme Fallback**: Uses `window.matchMedia('(prefers-color-scheme: dark)')`.
- **User Override**: Users can cycle `Auto` | `Light` | `Dark` directly from the header pill, saved to `localStorage`.

### 11.2. Reactive Language Switching
- **Auto-Detection**: Senses `Office.context.displayLanguage` and browser locale.
- **Dynamic UI Flip**: Toggling `中文` / `EN` in the header immediately translates tab labels, button text, search placeholders, and flips the primary title in the Command Palette between Chinese and English without reloading the page.


