# ExcelShortcut Project Rules & Architectural Governance

This document defines the core project rules, reference implementations, and architectural specifications for the **ExcelShortcut** Office Add-in.

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

1. **Excel Shared Runtime Scenario**:
   - Reference Clone: [`references/excel-shared-runtime-scenario`](../references/excel-shared-runtime-scenario)
   - Demonstrates unified Shared Runtime (`lifetime: "long"`), shared state between ribbon buttons, custom functions, and the taskpane.
2. **Office Agents Package & Bridge**:
   - Production pattern for Office Add-in runtimes interacting with autonomous agents and local WebSocket RPC.
3. **Microsoft Office.js Documentation (Local Reference)**:
   - Reference Clone: [`references/office-js-docs/docs`](../references/office-js-docs/docs)
   - Authoritative offline source for exact Excel JavaScript APIs, manifest specifications, and Shared Runtime requirements.

---

## 2. 4-Category Active Project Rules

### 1. Build & Runtime Rules

- **Rule 1.1 (Shared Runtime Mandate)**:
  `ExcelShortcut` MUST use a Shared JavaScript Runtime (`<Runtime resid="..." lifetime="long"/>`). The single JS context persists across workbook lifetime, allowing headless `Excel.run(...)` background automation. Closing the taskpane via UI MUST NEVER terminate the WebSocket connection.
- **Rule 1.2 (Dual-Manifest Identity Segregation)**:
  > **Rule 1.2 (Updated 2026-11)**: Local Dev and GitHub Production MUST maintain distinct manifest GUIDs and visual ribbon controls. *Sharing GUID `c1b06178` across local and prod caused fatal WEF cache clobbering and is strictly forbidden.*
  - **Local Dev Manifest** ([`manifest.xml`](../manifest.xml)): Dedicated GUID `d2c17289-5195-5b1f-902f-8bdeeef20031`, DisplayName `[DEV] Excel WebSocket`, Ribbon label `[DEV] Excel WS`.
  - **Production Manifest** ([`manifest.github.xml`](../manifest.github.xml)): Canonical GUID `c1b06178-4084-4a0e-8f1e-7acddec19930`, DisplayName `Excel WebSocket`. Generated automatically by [`scripts/build-manifest.js`](../scripts/build-manifest.js).
- **Rule 1.3 (Deployment & Catalog Standards)**:
  - Production add-in is hosted statically 24/7 on GitHub Pages CDN (`https://ruhchjdbdbhccjjc.github.io/excel-websocket-addin/`) with 0 local server processes.
  - Sideloading into Excel Desktop must use the Trusted Network Catalog ([`\\DESKTOP-TP761MQ\Users\Public\OfficeManifests`](file:///C:/Users/Public/OfficeManifests)).
- **Rule 1.4 (Cross-Origin WSS Transport)**:
  - Excel on the Web requires secure WebSockets (`wss://localhost:8891`) via [`scripts/wss-proxy.js`](../scripts/wss-proxy.js) emitting `Access-Control-Allow-Private-Network: true`.
  - Excel Desktop connects directly to `ws://127.0.0.1:8890`.
- **Rule 1.5 (WEF Cache Invalidation Protocol)**:
  Excel aggressively caches manifests in `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`. Whenever manifest GUIDs, ribbon group IDs, or URLs change, close Excel and purge the WEF cache.

### 2. TypeScript & Code Standards

- **Rule 2.1 (File Line Limits)**:
  - TypeScript (`.ts`): Maximum 400 lines per module.
  - UI Components (`.tsx`): Maximum 250 lines per component.
  - Styles (`.css`): Maximum 500 lines.
  - Propose modular refactoring before expanding files beyond these limits.
- **Rule 2.2 (Master Automation Toggle Enforcement)**:
  Every WebSocket connection attempt must check [`isAutomationEnabled()`](../src/websocket/client/ConnectionManager.ts#L116). When disabled by user, the client must immediately close the socket (`ws.close(1000)`), cancel reconnect backoff timers, and dim the ribbon button.
- **Rule 2.3 (Multi-Listener Observer Registries)**:
  State managers (e.g. [`ConnectionManager`](../src/websocket/client/ConnectionManager.ts)) must store event listeners in arrays (`eventHandlers[event] = [...]`). Overwriting single callback references is forbidden.
- **Rule 2.4 (Split-File i18n Rules)**:
  [`src/i18n/types.ts`](../src/i18n/types.ts) is the single source of truth for `TranslationSchema`. All locale files under `src/i18n/locales/` must be modular (< 100 lines) and strictly type-checked.
- **Rule 2.5 (Relative Clickable Links)**:
  All markdown references to source code and documentation must use relative paths with functional anchor tags (e.g., `[ConnectionManager](../src/websocket/client/ConnectionManager.ts#L100)`).

### 3. Cross-OS & Path Handling Rules

- **Rule 3.1 (Deterministic URL & Script Resolution)**:
  Node scripts must resolve paths via `path.join(__dirname, ...)` and normalize forward slashes (`/`) for web resources and manifest `<AppDomain>` tags.
- **Rule 3.2 (Trusted Catalog UNC Compatibility)**:
  Windows network shares must be referenced using standard UNC format (`\\<ComputerName>\<Share>`) or mapped local paths (`C:\Users\Public\OfficeManifests`).
- **Rule 3.3 (UTF-8 Encoding Invariant)**:
  All TypeScript, JSON, XML, and localization dictionaries must be saved with UTF-8 encoding without BOM.

### 4. Forbidden Anti-Patterns

- **Anti-Pattern 4.1 (DO NOT Use Ephemeral Developer Debug Keys for Production)**:
  DO NOT insert production add-ins via `HKCU:\Software\Microsoft\Office\16.0\WEF\Developer`. Office treats them as temporary Visual Studio debug sessions and locks ribbon buttons with *"This add-in is no longer available"* once debugging ends.
- **Anti-Pattern 4.2 (DO NOT Duplicate Action Associations)**:
  DO NOT call `Office.actions.associate()` in multiple files. Duplicate associations in `commands.ts` and `taskpane.ts` crash the shared runtime with fatal `[DuplicatedName] Function=action` errors.
- **Anti-Pattern 4.3 (DO NOT Match Generic Formulas for Spotlight Purge)**:
  DO NOT use bare `f.includes('INDEX(')` to identify spotlight conditional formats. It deletes user custom conditional formatting formulas. Use [`isSpotlightFormula()`](../src/websocket/service/ExcelSpotlightService.ts) requiring coordinate signatures.
- **Anti-Pattern 4.4 (DO NOT Use Mermaid Blocks in Docs)**:
  DO NOT use Mermaid diagrams (` ```mermaid `) in documentation. Always use ASCII logic maps (`|`, `+`, `-`, `>`) to guarantee rendering across all editors and terminals.
- **Anti-Pattern 4.5 (DO NOT Speculatively Mutate Tool Envelopes)**:
  DO NOT enable speculative Markdown parser fallback heuristics when parsing agent tool calls or WebSocket envelopes. Maintain byte-exact payload invariants.
