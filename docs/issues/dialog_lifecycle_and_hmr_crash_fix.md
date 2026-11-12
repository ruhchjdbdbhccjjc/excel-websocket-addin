# Dialog Lifecycle Reopen Reliability and Webpack HMR Crash Resolution

**Status**: Fixed (2026-10-24)  
**Area**: Excel Office Add-in · Office.js Dialog API · Webpack HMR · Shared Runtime  

---

## 1. Problem Description

1. **Webpack HMR Crash**: When editing files while the dev server was running, Webpack HMR attempted to hot-reload `taskpane.js` and `commands.js`. Because both chunks were included on the same HTML page without a shared runtime chunk, `self.webpackHotUpdateenter` was overwritten by `commands.js`, throwing `Uncaught TypeError: Cannot set properties of undefined (setting './src/websocket/service/ExcelCommandExecutor.ts')`. This uncaught exception killed the embedded WebView2 execution context and disconnected the dev server (`ECONNREFUSED ::1:3000`).
2. **Dialog "Only Works Once" Behavior**:
   - `App.tsx` had an unconditioned listener on `connMgr.on('messageReceived')` for `open_command_palette` that called `Office.addin.showAsTaskpane()`. Whenever the floating dialog was requested, `App.tsx` simultaneously forced the taskpane open, stealing focus and creating conflicting window state transitions.
   - When dialogs closed or failed, `currentDialog` was not reliably cleared to `null` if the reference differed or if error 12007 retries exhausted.
   - Any stale `currentDialog` reference caused subsequent calls to invoke toggle-close instead of opening, requiring multiple hotkey presses.
   - `close_dialog` command did not update `lastDialogCloseTime` or reset `isOpeningDialog`.

---

## 2. Architecture & Execution Flow

```text
+-----------------------------------------------------------------------------------+
|                         DIALOG LIFECYCLE & HMR SAFEGUARD                          |
+-----------------------------------------------------------------------------------+

                    [ Incoming Command: "open_command_palette" ]
                                         |
               +-------------------------+-------------------------+
               |                                                   |
               v (BackgroundService / Shared Runtime)              v (App.tsx UI Listener)
   [ExcelCommandExecutor.openPaletteDialog()]           [connMgr.on('messageReceived')]
               |                                                   |
   +-----------+-----------+                     +-----------------+-----------------+
   |                       |                     |                                   |
   | Is currentDialog set? |                     | mode === 'taskpane'?              |
   |                       |                     |                                   |
   +-----------+-----------+                     +--------+-----------------+--------+
         |           |                                    |                 |
  (Yes)  |           | (No)                         (Yes) |                 | (No)
         v           v                                    v                 v
   [Close &    [Cooldown Buffer >= 400ms]        [showAsTaskpane()]    [Do NOT force taskpane]
    Reset]           |                                                 (Leave focus for dialog)
                     v
        [displayDialogAsync(dialogUrl)]
                     |
       +-------------+-------------+
       |                           |
   (Succeeded)             (Error 12007 Retry)
       |                           |
       v                           v
  [Track dialog]           [Exponential backoff up to 12 retries]
  [Message / Event         [Reset currentDialog & lastDialogCloseTime on exit]
   unconditionally
   nullify reference]
```

---

## 3. Root Causes & Resolutions

### 3.1. Webpack Multiple Entrypoint Collision & Dev Server HMR
- **Root Cause**: Webpack 5 requires `optimization.runtimeChunk = 'single'` when multiple entry points (`taskpane` and `commands`) exist on a single HTML document (`taskpane.html`). Without this, each entrypoint defines its own runtime global (`self.webpackHotUpdateenter`), corrupting the module cache on HMR updates. Furthermore, HMR in Office.js WebView2 often causes desync.
- **Resolution**:
  - Added `optimization: { runtimeChunk: "single" }` to [webpack.config.js](../../webpack.config.js).
  - Configured `devServer: { hot: false, liveReload: false, client: { overlay: false } }` to eliminate HMR-induced WebView2 crashes.

### 3.2. Taskpane & Dialog Race Condition
- **Root Cause**: `src/taskpane/App.tsx` was listening to `open_command_palette` and unconditionally calling `Office.addin.showAsTaskpane()`. This fought against the floating dialog opened by `ExcelCommandExecutor`.
- **Resolution**: Updated [src/taskpane/App.tsx](../../src/taskpane/App.tsx) so `showAsTaskpane()` is only triggered when explicit taskpane mode is requested (`mode === 'taskpane'`).

### 3.3. Unconditional Dialog State Nullification & Cooldown Expansion
- **Root Cause**: `currentDialog` was conditionally reset only when `this.currentDialog === activeDialog`. On errors or stale states, `currentDialog` stayed truthy, trapping future calls in the "close toggle" branch.
- **Resolution**:
  - Expanded cooldown buffer to 400ms in [src/websocket/service/ExcelCommandExecutor.ts](../../src/websocket/service/ExcelCommandExecutor.ts).
  - Unconditionally set `this.currentDialog = null` and updated `this.lastDialogCloseTime = Date.now()` on `DialogMessageReceived`, `DialogEventReceived`, `close_dialog`, and retry exhaustions.
  - Enhanced `isError12007` to recognize both numeric codes and string message contents.

---

## 4. Verification

1. `npm run build:dev` passed cleanly; `runtime.js` is generated and cleanly shared across all entrypoints.
2. `npm run build` completed production bundling without collision.
3. `npm run validate` passed Office manifest schema verification.
