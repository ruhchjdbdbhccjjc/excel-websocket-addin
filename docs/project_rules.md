# ExcelShortcut Project Rules & Architectural Specifications

This document defines the local project rules, reference implementations, and architectural specifications for the **ExcelShortcut** Office Add-in.

```text
+-----------------------------------------------------------------------------------+
|                        EXCELSHORTCUT ARCHITECTURAL TOPOLOGY                       |
+-----------------------------------------------------------------------------------+

   +--------------------------+               +--------------------------+
   |   KANATA / CLI CLIENT    |               |  OFFICE-AGENTS / BRIDGE  |
   +--------------------------+               +--------------------------+
                |                                          |
                +--------------------+---------------------+
                                     |
                                     | (Local WebSocket / JSON-RPC)
                                     v
                  +--------------------------------------+
                  |         EXCEL SHARED RUNTIME         |
                  |           lifetime: "long"           |
                  +--------------------------------------+
                                     |
               +---------------------+---------------------+
               |                                           |
               v                                           v
+-----------------------------+             +-----------------------------+
|    HEADLESS / BACKGROUND    |             |     TASKPANE (OPTIONAL)     |
+-----------------------------+             +-----------------------------+
| * WebSocket Listener        |             | * Office.addin.showAsTaskpane()
| * Excel.run(...) execution  |             | * Office.addin.hide()       |
| * Ribbon ExecuteFunction    |             | * Configuration UI          |
| * Custom Functions Hooks    |             | * Debugging & Log Monitor   |
+-----------------------------+             +-----------------------------+
```

---

## 1. Upstream & Reference Implementations

All developers and AI agents must consult and reference the following foundational repositories:

1. **Excel Shared Runtime Scenario**:
   - **Repository Link**: [Office-Add-in-samples / Samples / excel-shared-runtime-scenario](https://github.com/ruhchjdbdbhccjjc/Office-Add-in-samples/tree/main/Samples/excel-shared-runtime-scenario)
   - **Local Reference Clone**: Located in [references/excel-shared-runtime-scenario](../references/excel-shared-runtime-scenario)
   - **Purpose**: Demonstrates unified Shared Runtime (`lifetime: "long"`), sharing state between ribbon buttons, custom functions, and the taskpane, and controlling taskpane visibility programmatically.

2. **Office Agents Package Example**:
   - **Repository Link**: [hewliyang/office-agents](https://github.com/hewliyang/office-agents)
   - **Purpose**: Production architecture pattern for modern Office Add-in packages interacting with autonomous agents and tool-calling runtimes.

3. **Office Agents Bridge (WebSocket RPC)**:
   - **Repository Link**: [hewliyang/office-agents / packages / bridge](https://github.com/hewliyang/office-agents/tree/main/packages/bridge)
   - **Purpose**: Demonstrates local HTTPS/WebSocket RPC bridge architecture for inspecting, commanding, and driving live Microsoft Office add-in runtimes from external tools and CLI processes.

4. **Official Microsoft Office.js Documentation (Local Reference)**:
   - **Local Reference Clone**: Located in [references/office-js-docs](../references/office-js-docs/docs)
   - **Scope**: Sparse-checkout of `docs/excel/` (76 guides covering Range, Worksheet, Table, Formula, and Performance) and `docs/develop/` (71 guides covering Shared Runtime, Manifests, Add-in Commands, and Lifecycle).
   - **Purpose**: Authoritative offline source for exact Excel JavaScript APIs, manifest specifications, and Shared Runtime requirements.

---

## 2. Headless & Background Execution Rules

### Rule 2.1: Running Without Opening the Taskpane (Shared Runtime Mandate)
- **Problem**: In standard Office Add-ins, webview code only initializes when the user manually opens the Taskpane. When closed, background tasks or WebSocket connections can terminate.
- **Mandate**: `ExcelShortcut` MUST use a **Shared Runtime** (`<Runtime resid="..." lifetime="long"/>` in XML, or `"lifetime": "long"` in Unified Manifest).
- **Execution Flow**:
  1. The Shared Runtime hosts a single JavaScript context that persists for the lifetime of the Excel session.
  2. The background engine starts upon workbook open (via AutoRun / Custom Function hook) or on the first headless Ribbon command (`ExecuteFunction`).
  3. The add-in executes Excel operations via `Excel.run(...)` silently in the background without needing DOM UI.

### Rule 2.2: Programmatic Taskpane Visibility
- The Taskpane is **optional** and should not be forced open on the user.
- To open the taskpane on demand:
  ```typescript
  await Office.addin.showAsTaskpane();
  ```
- To dismiss the taskpane without terminating the WebSocket or background listeners:
  ```typescript
  await Office.addin.hide();
  ```
- Closing the taskpane via UI MUST NEVER terminate the WebSocket connection.

---

## 3. WebSocket & Bridge Integration Standards

### Rule 3.1: Decoupled Lifecycle
- WebSocket client initialization (`ConnectionManager`) MUST be decoupled from Taskpane DOM elements (`document.getElementById`).
- If DOM elements are absent (running headless), the client must log to internal memory/console and continue operations without throwing DOM errors.

### Rule 3.2: Reconnection & Heartbeat
- The WebSocket client must implement exponential backoff reconnection.
- It must handle network sleep, Excel suspension, and Kanata restarts gracefully without memory leaks.

---

## 4. Code Generation & Maintenance Guardrails

### Rule 4.1: File Line Limits
To prevent context degradation and maintain high readability:
- **TypeScript / JavaScript (.ts / .js)**: Maximum 400 lines per module.
- **UI Components / HTML**: Maximum 250 lines.
- **Styles (.css)**: Maximum 500 lines.
- If a file exceeds these limits, propose refactoring before adding new logic.

### Rule 4.2: Visual Logic Maps
- Lead with an ASCII art logic map for every technical explanation, architecture proposal, and feature documentation.
- Do NOT use Mermaid diagrams in markdown documentation to prevent parser errors in various markdown viewers.

### Rule 4.3: Relative Clickable Links
- All markdown references to source code and sub-documents MUST use relative paths with functional anchor tags (e.g., `[ConnectionManager](../src/websocket/client/ConnectionManager.ts)`).

---

## 5. Feature Implementation & Root Cause Analysis Protocol

When implementing new Office Add-in features or investigating unexpected bugs, follow this 5-stage protocol to prevent runtime misconfigurations and regressions:

```text
+-----------------------------------------------------------------------------------+
|                 OFFICE ADD-IN FEATURE & ROOT CAUSE ANALYSIS FLOW                  |
+-----------------------------------------------------------------------------------+

     [ New Feature / Bug Report ]
                  |
                  v
     [ STAGE 1: Local Docs First ] <--- Search references/office-js-docs/docs/excel/
                  |                     and references/excel-shared-runtime-scenario/
                  v
     [ STAGE 2: Architecture Check] <--- Validate Shared Runtime, <Runtimes>,
                  |                     single Action association, & event arrays
                  v
     [ STAGE 3: Runtime Diagnosis ] <--- Inspect Edge Chromium DevTools / Office console
                  |                     Catch InvalidRibbonDefinition, DuplicatedName
                  v
     [ STAGE 4: WEF Cache Purge ]  <--- Clear %LOCALAPPDATA%\Microsoft\Office\16.0\Wef\
                  |                     whenever manifest or control IDs change
                  v
     [ STAGE 5: Targeted Search ]  <--- Search GitHub issues (OfficeDev/office-js)
                                        ONLY for obscure host quirks or undocumented bugs
```

### Stage 1: Local Documentation First
- Before designing new APIs or commands, perform a fast local `grep_search` in `references/office-js-docs/`:
  - For Excel range, sheet, and formula logic: check `references/office-js-docs/docs/excel/`.
  - For manifest schema, Ribbon commands, and shared runtimes: check `references/office-js-docs/docs/develop/`.
- Cross-reference with the working sample in `references/excel-shared-runtime-scenario/`.

### Stage 2: Architecture & Manifest Validation
- **Shared Runtime Verification**: Ensure `<Requirements>` contains `<Set Name="SharedRuntime" MinVersion="1.1"/>` and `<Host xsi:type="Workbook">` contains `<Runtimes><Runtime resid="Taskpane.Url" lifetime="long" /></Runtimes>`.
- **Single Action Association**: Ensure `Office.actions.associate()` is called in exactly one module. Duplicate registrations in both `commands.ts` and `taskpane.ts` cause fatal `[DuplicatedName] Function=action` collisions.
- **Multi-Listener Events**: Any state manager (e.g., `ConnectionManager`) must store listeners in arrays (`eventHandlers[event] = [...]`), never single function overwrites.

### Stage 3: Runtime Diagnosis
- Office.js compiles cleanly in TypeScript even when manifest or host rules are violated.
- Always inspect the runtime console:
  - Excel Win32: Right-click taskpane -> **Inspect**, or attach Edge DevTools via `F12` / `edge://inspect`.
  - Catch silent failures such as `InvalidRibbonDefinition` or `RichApi.Error`.

### Stage 4: Office WEF Cache Invalidation
- Excel Desktop aggressively caches manifests in `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.
- Whenever you modify `manifest.xml`, ribbon group labels, button IDs, or runtime declarations, close Excel and purge the WEF cache folder to guarantee the changes take effect.

### Stage 5: Targeted Upstream Search
- If a runtime error is not explained by the local documentation, search the official repository issues at `https://github.com/OfficeDev/office-js/issues` or `search_web` targeting the exact error code (e.g., `RichApi.Error InvalidRibbonDefinition`).

---

## 4. Architectural Rules for Core Subsystems

### 4.1. Split-File Internationalization (i18n) Rules
1. **Separation of Contract and Dictionaries**:
   - `src/i18n/types.ts` is the single source of truth for `TranslationSchema`.
   - Every locale file under `src/i18n/locales/` (e.g. `zh-CN.ts`, `en-US.ts`) must implement `TranslationSchema`.
   - No ad-hoc string literals in UI components; all labels, tooltips, and toasts must be indexed via `t().section.key`.
2. **File Size Constraints**:
   - Keep locale dictionaries strictly modular (<100 lines each).
   - Use compile-time type exports (`export type { ... }`) to avoid Webpack/Babel empty module export warnings.
3. **Detection Hierarchy**:
   - Prioritize explicit user selection in `localStorage` -> `Office.context.displayLanguage` -> `navigator.language` -> `en-US` default fallback.

### 4.2. Zero-Destruction Conditional Formatting & Detox Rules
1. **Strict Formula Signature Discrimination**:
   - Never match generic Excel formula functions (such as bare `INDEX(`, `MATCH(`, or `ISNUMBER(`) when finding add-in conditional format rules.
   - Always route formula validation through `isSpotlightFormula()` requiring exact coordinate anchors (`ROW()=`, `COLUMN()=`, `OR(ROW()=`, `AND(ROW()>=`) or specialized peer-match patterns (`INDEX(1:1048576,ROW(),MATCH(`).
2. **Cross-Sheet Ghost Prevention**:
   - Excel conditional formats are worksheet-scoped.
   - When active selection transitions to another sheet, always identify `prevSheetName` and clear `getItemOrNullObject(prevSheetName)` to prevent phantom crosshairs from lingering across sheets.
3. **Trailing Execution Queue for IPC Synchronization**:
   - When an asynchronous `Excel.run` / `context.sync()` call is in-flight, rapid user input (e.g. fast arrow navigation) must NOT be silently dropped.
   - Store incoming selections in `pendingSelection` and drain the queue in the `finally` block once the active sync completes.
4. **Unconditional Startup Auto-Detox**:
   - Add-in startup in `BackgroundService.initialize()` must unconditionally run `purgeAllOrphanSpotlightsAcrossWorkbook()` to scrub leftover rules from previous ungraceful exits or crashes.
   - Multi-sheet purges must batch collection and formula queries into a 3-phase atomic pipeline to avoid $2N$ round-trips over the Office.js bridge.

### 4.3. Bilingual Command Management & Search Rules
1. **Canonical Source of Truth**:
   - The universal anchor key is `idMso` (from `excelcontrols.xlsx`).
   - Official labels and supertips must be extracted directly via Excel Win32 COM (`GetLabelMso`, `GetSupertipMso`) using local language packs (`1033` English, `2052` Chinese). Do not rely on generic termbases (e.g. `CHINESE (SIMPLIFIED).tbx`) which lack `idMso`.
2. **Pinyin and Acronym Pre-Indexing**:
   - Chinese labels must have pre-computed full Pinyin and Pinyin initials (`initials`, e.g. `djcg` for `冻结窗格`).
   - English labels/IDs must have pre-computed acronyms (`acronym`, e.g. `fp` for `FreezePanes`).
3. **Sub-Millisecond Search Latency & Decoupled LRU Caching**:
   - The search engine pre-indexes tokens at import time.
   - The LRU query cache must be keyed strictly by the normalized query string (`q`), completely decoupled from the recent command list (`recentIds`).
   - Dynamic MRU recency boosts must be computed dynamically on top cached candidate items in memory ($O(K)$) so that executing a command never invalidates search cache entries for subsequent keystrokes.
4. **Recent History Pollution Protection**:
   - Only verified official/registered commands (`cmd.type !== 'custom'`) may be persisted to the `recentIds` storage. Arbitrary typo submissions from the custom fallback route must never be persisted.

### 4.4. Theme Detection & Synchronization Rules
1. **Dynamic Perceived Luminance Detection**:
   - Never rely on brittle hardcoded hex comparisons (`#262626`).
   - Evaluate `Office.context.officeTheme.bodyBackgroundColor` using ITU-R BT.601 perceived luminance `((r * 299 + g * 587 + b * 114) / 1000) < 128` to support all current and future Microsoft 365 dark variants (`#333333`, `#201F1E`, `#1E1E1E`, etc.).
   - Fall back to `window.matchMedia('(prefers-color-scheme: dark)')` when `officeTheme` is unavailable.
2. **Safe Deferred Lifecycle Registration**:
   - Theme event handlers (`officeTheme.addHandlerAsync`) must only be registered within `Office.onReady()` to eliminate race conditions before Office.js context is initialized.
3. **User Theme Override Persistence**:
   - Support `auto` (follows Excel / OS), `light`, and `dark`.
   - Store user preference in `localStorage` under `excel_addin_theme_mode`.
   - Apply theme globally to `document.body` via `data-theme="light"` or `data-theme="dark"`.


