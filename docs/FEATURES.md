# Excel Shortcut Features

## WebSocket Integration with Kanata

The Excel Shortcut add-in now includes WebSocket client functionality to communicate with the Kanata keyboard layout manager. This allows real-time bidirectional communication between the Excel add-in and Kanata.

### Overview

The WebSocket integration enables:
- Real-time connection management with Kanata server
- Message parsing and handling (strings, arrays, objects)
- Automatic reconnection with exponential backoff
- Heartbeat mechanism to keep connections alive
- Configurable connection settings (port, client name, channels)
- Log-level based logging system

### Usage

The WebSocket connection is managed through the task pane UI:
1. Configure connection settings (port, client name, client type, channels)
2. Click "Connect" to establish connection
3. View connection status and received messages in the log viewer
4. Click "Disconnect" to close the connection

### Connection Settings

- **Port**: WebSocket server port (default: 8890)
- **Client Name**: Identifier for the client (default: excel)
- **Client Type**: Type of client (default: office-addin)
- **Channels**: Comma-separated list of channels to subscribe to (default: excel-addin-websocket)
- **Auto-connect**: Automatically connect on add-in load (disabled by default)

### Architecture

The WebSocket implementation consists of several components:

- **[WebSocketClient](../src/websocket/client/WebSocketClient.ts)**: Core WebSocket connection management
- **[ConnectionManager](../src/websocket/client/ConnectionManager.ts)**: High-level state management and settings
- **[MessageHandler](../src/websocket/client/MessageHandler.ts)**: Message parsing and routing
- **[KanataProtocol](../src/websocket/protocol/KanataProtocol.ts)**: Kanata-specific message formatting
- **[Logger](../src/websocket/logging/Logger.ts)**: Log-level based logging utility
- **[SettingsStorage](../src/websocket/storage/SettingsStorage.ts)**: Persistent settings storage

> Deep Dive: See [WebSocket Integration Details](features/websocket_integration_details.md)

## Excel Commands & Position Slots

The add-in provides extensive headless Excel operations dispatched via WebSocket:
- **Zero-Pollution Position Slots**: Fast workbook-level and sheet-local slot saving and switching stored directly inside `Office.context.document.settings` without modifying cell values or formatting.
- **Border Color Cycling (`set_border`)**: Cycles 6 distinct colors (`#E26D5C`, `#C7CC7A`, `#7560BA`, `#9DD9D2`, `#FFE1A8`, `#FFFFFF` clear) and displays an automatic 1.5-second non-intrusive floating shape badge to reveal border colors while cells remain selected.
- **Focus Auto-Filtering**: Mutating commands auto-filter to only execute when the corresponding Excel window is active.
- **Comprehensive Command Suite**: Full navigation, selection, zoom, clear, row/column operations, formulas, and data export.

> Deep Dive: See [Position Slots and Border Color Cycle Details](features/position_slots_and_border_cycle_details.md)
> Reference: See [WebSocket Commands Reference](features/websocket_commands_reference.md)

## WPS-Style Spotlight Engine & Anchor Column Matching

The add-in includes a zero-destruction spotlight highlighting system powered by native Excel conditional formatting:
- **Instant Crosshair & Direction Control**: Highlights active row, column, or crosshair using iconic WPS pastel palettes (Yellow, Green, Blue, Purple, Peach).
- **Zero Pollution & Safe Isolation**: Only touches own custom spotlight formulas; never alters cell values or manual background formatting, with 0 impact on the undo stack.
- **Anchor Key Column Highlighting**: When clicking any cell on a row (e.g., `B7`), the target column cell on that same row (`A7`) lights up with a designated accent color or 4-sided enclosing border box.
- **Flexible Column Resolution**: Supports column letters (`A`), 1-based numeric column indices (`1`-`16384`), and dynamic header titles (e.g. `#ID` or `姓名`) evaluated via `MATCH("...", $1:$1, 0)`.
- **Peer Value Matching**: Dynamically highlights all rows matching the active row's key column value.
- **Adaptive Selection Behavior**: Configurable via `active_cell` (follows focus), `banded_area` (expands to whole selection), or `auto_hide` (suspends highlight during multi-selection).

> Deep Dive: See [Excel Spotlight Anchor & Column Matching](features/excel_spotlight_anchor_matching.md)
> Deep Dive: See [Adaptive Selection Spotlight Engine](features/excel_spotlight_adaptive_selection.md)

## Split-File i18n & Whole-Workbook Crash Detox

- **Split-File i18n Architecture**: Strongly typed translation schema (`src/i18n/types.ts`) split into modular locale files (`locales/zh-CN.ts` and `locales/en-US.ts`). Supports multi-tier runtime auto-detection (`localStorage` -> `Office.context.displayLanguage` -> `navigator.language` -> `en-US` fallback) and reactive SolidJS locale switching.
- **Whole-Workbook Crash Detox & Lifecycle Purge**: High-performance 3-roundtrip batched scanner (`purgeAllOrphanSpotlightsAcrossWorkbook`) that cleans orphaned spotlight conditional formatting left across all sheets after crashes or abrupt closures. Unconditionally executed during startup auto-detox in `BackgroundService` and bound to browser `beforeunload` / `pagehide` hooks.
- **Zero False-Positive Formula Discrimination**: `isSpotlightFormula()` strictly isolates spotlight coordinate rules and peer-match patterns, completely preserving user custom conditional formatting formulas.

> Deep Dive: See [Split-File i18n & Whole-Workbook Crash Detox](features/i18n_and_crash_detox_details.md)

## Unified Diagnostic Logging Hub

- **Singleton Ring Buffer**: Fixed-size 1000-entry memory ring buffer capturing system, protocol, and command events without memory growth.
- **Global Error & Console Interception**: Hooks into `console.log/warn/error` and `window.onerror` / `unhandledrejection` to catch silent failures and runtime exceptions.
- **One-Click Diagnostic Log Bundle Export**: Exports formatted `.log` diagnostic bundles complete with Office host fingerprints, connection states, and timestamped trace lines for instant troubleshooting.

## Development

The project uses:
- TypeScript for type safety
- Webpack for bundling
- Office.js for Microsoft 365 integration
- Fluent UI for consistent styling

## Headless Execution & Shared Runtime

The add-in architecture is configured to support running in the background without requiring the task pane to be opened:
- **Shared Runtime**: Enables persistent JavaScript execution and continuous WebSocket communication.
- **Reference Sample**: See [references/excel-shared-runtime-scenario](../references/excel-shared-runtime-scenario) for Microsoft's official implementation.
- **Project Guidelines**: See [project_rules.md](project_rules.md) for architectural rules and integration standards with Kanata and Office Agents.


