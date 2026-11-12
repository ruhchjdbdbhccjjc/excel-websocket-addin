# Excel Shortcut Architecture

## Overview

The Excel Shortcut add-in is a Microsoft Office JavaScript API add-in that provides WebSocket integration with the Kanata keyboard layout manager. The add-in is built with TypeScript, uses Webpack for bundling, and follows the Microsoft 365 Agents Toolkit structure.

## Project Structure

```
ExcelShortcut/
├── appPackage/           # Office add-in manifest and assets
├── dist/                # Built output directory
├── docs/                # Documentation
│   ├── features/        # Feature documentation
│   ├── issues/          # Issue resolutions
│   └── project_rules.md # Local rules and reference repositories
├── env/                 # Environment configuration
├── infra/               # Infrastructure files
├── node_modules/        # npm dependencies
├── references/          # Reference repositories and samples
│   └── excel-shared-runtime-scenario/ # Microsoft shared runtime scenario
├── scripts/             # Build and deployment scripts
└── src/                 # Source code
    ├── commands/        # Office.js command functions
    ├── i18n/            # Split-file reactive internationalization
    │   ├── locales/     # zh-CN and en-US dictionaries
    │   ├── types.ts     # TranslationSchema definition
    │   └── index.ts     # useI18n reactive hook & detection
    ├── taskpane/        # Task pane UI and logic
    ├── utils/           # Utility functions
    └── websocket/       # WebSocket client implementation
        ├── client/      # Client connection & message handler
        ├── logging/     # Singleton ring-buffer logger & error hooks
        ├── protocol/    # Kanata message protocol definitions
        ├── service/     # Command execution, position slots & spotlight
        └── storage/     # Settings storage & schemas
```

## Components

### 1. Task Pane UI

**Location**: [`src/taskpane/`](../src/taskpane/)

**Purpose**: Provides the user interface for the add-in, including WebSocket connection controls and Excel command buttons.

**Files**:
- [`taskpane.html`](../src/taskpane/taskpane.html) - Task pane HTML structure
- [`taskpane.ts`](../src/taskpane/taskpane.ts) - Task pane logic and event handling
- [`taskpane.css`](../src/taskpane/taskpane.css) - Task pane styling
- [`excel.ts`](../src/taskpane/excel.ts) - Excel-specific task pane logic
- [`SpotlightTab.tsx`](../src/taskpane/components/spotlight/SpotlightTab.tsx) - Spotlight control panel and rule stack
- [`SpotlightRuleCard.tsx`](../src/taskpane/components/spotlight/SpotlightRuleCard.tsx) - Rule card supporting anchor cell and peer matching
- [`ColorPills.tsx`](../src/taskpane/components/spotlight/ColorPills.tsx) - Pastel color palette picker

**Key Functions**:
- [`initializeWebSocket()`](../src/taskpane/taskpane.ts#L36) - Initializes WebSocket connection manager
- [`setupEventListeners()`](../src/taskpane/taskpane.ts#L77) - Sets up UI event handlers
- [`updateConnectionUI()`](../src/taskpane/taskpane.ts#L220) - Updates connection status display
- [`updateLogViewer()`](../src/taskpane/taskpane.ts#L276) - Updates message log display
- [`runExcel()`](../src/taskpane/taskpane.ts#L301) - Executes Excel command

**Dependencies**:
- Office.js for Microsoft 365 integration
- WebSocket client components
- Logger for logging

### 2. Commands

**Location**: [`src/commands/`](../src/commands/)

**Purpose**: Contains Office.js command functions for different Office applications.

**Files**:
- [`commands.ts`](../src/commands/commands.ts) - Command registration
- [`excel.ts`](../src/commands/excel.ts) - Excel-specific commands
- [`outlook.ts`](../src/commands/outlook.ts) - Outlook-specific commands
- [`powerpoint.ts`](../src/commands/powerpoint.ts) - PowerPoint-specific commands
- [`word.ts`](../src/commands/word.ts) - Word-specific commands

**Key Functions**:
- [`runExcel()`](../src/commands/excel.ts) - Main Excel command handler

**Dependencies**:
- Office.js API

### 3. WebSocket Client

**Location**: [`src/websocket/`](../src/websocket/)

**Purpose**: Provides WebSocket client functionality for real-time communication with Kanata server.

**Subcomponents**:

#### 3.1 ConnectionManager

**Location**: [`src/websocket/client/ConnectionManager.ts`](../src/websocket/client/ConnectionManager.ts)

**Purpose**: High-level connection state management, settings persistence, and event coordination.

**Key Functions**:
- [`connect()`](../src/websocket/client/ConnectionManager.ts#L97) - Initiates connection
- [`disconnect()`](../src/websocket/client/ConnectionManager.ts#L107) - Terminates connection
- [`getSettings()`](../src/websocket/client/ConnectionManager.ts#L49) - Retrieves settings
- [`saveSettings()`](../src/websocket/client/ConnectionManager.ts#L78) - Persists settings
- [`getConnectionState()`](../src/websocket/client/ConnectionManager.ts#L113) - Gets current state

**Dependencies**:
- WebSocketClient
- MessageHandler
- SettingsStorage
- Logger

#### 3.2 WebSocketClient

**Location**: [`src/websocket/client/WebSocketClient.ts`](../src/websocket/client/WebSocketClient.ts)

**Purpose**: Low-level WebSocket connection management, reconnection logic, and heartbeat.

**Key Functions**:
- [`connect()`](../src/websocket/client/WebSocketClient.ts#L54) - Establishes WebSocket connection
- [`disconnect()`](../src/websocket/client/WebSocketClient.ts#L85) - Closes connection
- [`send()`](../src/websocket/client/WebSocketClient.ts#L108) - Sends JSON message
- [`sendRaw()`](../src/websocket/client/WebSocketClient.ts#L123) - Sends raw string
- [`isConnected()`](../src/websocket/client/WebSocketClient.ts#L140) - Checks connection status

**Dependencies**:
- KanataProtocol
- Logger

#### 3.3 MessageHandler

**Location**: [`src/websocket/client/MessageHandler.ts`](../src/websocket/client/MessageHandler.ts)

**Purpose**: Parses incoming messages and routes them to appropriate handlers.

**Key Functions**:
- [`handleMessage()`](../src/websocket/client/MessageHandler.ts#L20) - Main message entry point
- [`parseStringMessage()`](../src/websocket/client/MessageHandler.ts#L46) - Parses string messages
- [`parseArrayMessage()`](../src/websocket/client/MessageHandler.ts#L60) - Parses array messages
- [`parseObjectMessage()`](../src/websocket/client/MessageHandler.ts#L78) - Parses object messages

**Dependencies**:
- KanataProtocol types

#### 3.4 KanataProtocol

**Location**: [`src/websocket/protocol/KanataProtocol.ts`](../src/websocket/protocol/KanataProtocol.ts)

**Purpose**: Handles Kanata-specific message formatting and parsing.

**Key Functions**:
- [`parseMessage()`](../src/websocket/protocol/KanataProtocol.ts#L18) - Parses incoming messages
- [`formatRegisterMessage()`](../src/websocket/protocol/KanataProtocol.ts#L44) - Creates registration message
- [`formatHeartbeat()`](../src/websocket/protocol/KanataProtocol.ts#L66) - Creates heartbeat message

**Dependencies**:
- Protocol types
- UUID generator

#### 3.5 Logger

**Location**: [`src/websocket/logging/Logger.ts`](../src/websocket/logging/Logger.ts)

**Purpose**: High-performance singleton diagnostic logging system equipped with a 1000-entry fixed ring buffer, scoped subsystem tags, and global browser exception interception.

**Key Functions**:
- [`getInstance()`](../src/websocket/logging/Logger.ts#L36) - Returns shared Logger singleton
- [`installGlobalHooks()`](../src/websocket/logging/Logger.ts#L44) - Intercepts native `console.log/warn/error` and `window.onerror`/`unhandledrejection`
- [`log()`](../src/websocket/logging/Logger.ts#L80) - Enqueues entry into 1000-element ring buffer with scope tags
- [`debug()`, `info()`, `warn()`, `error()`](../src/websocket/logging/Logger.ts#L100-L115) - Level-specific log dispatch
- [`getLogs()`](../src/websocket/logging/Logger.ts#L118) - Retrieves in-memory ring buffer for live UI and diagnostic bundle export
- [`clearLogs()`](../src/websocket/logging/Logger.ts#L122) - Clears in-memory log buffer

**Dependencies**:
- None (self-contained)

#### 3.6 SettingsStorage

**Location**: [`src/websocket/storage/SettingsStorage.ts`](../src/websocket/storage/SettingsStorage.ts)

**Purpose**: Manages persistent storage of connection settings using browser storage APIs.

**Key Functions**:
- [`get()`](../src/websocket/storage/SettingsStorage.ts#L18) - Retrieves single setting
- [`set()`](../src/websocket/storage/SettingsStorage.ts#L33) - Saves single setting
- [`getMultiple()`](../src/websocket/storage/SettingsStorage.ts#L49) - Retrieves multiple settings
- [`setMultiple()`](../src/websocket/storage/SettingsStorage.ts#L66) - Saves multiple settings

**Dependencies**:
- Office.js storage API

#### 3.7 ConnectionSettings

**Location**: [`src/websocket/storage/ConnectionSettings.ts`](../src/websocket/storage/ConnectionSettings.ts)

**Purpose**: Defines the interface and default values for connection settings.

**Exports**:
- `ConnectionSettings` interface
- `DEFAULT_SETTINGS` constant

**Dependencies**:
- None (type definitions only)

#### 3.8 Services

**Location**: [`src/websocket/service/`](../src/websocket/service/)

**Purpose**: Implements headless Excel command execution, position slot persistence, and background shared runtime lifecycle.

**Subcomponents**:

##### 3.8.1 ExcelCommandExecutor

**Location**: [`src/websocket/service/ExcelCommandExecutor.ts`](../src/websocket/service/ExcelCommandExecutor.ts)

**Purpose**: Dispatches and executes Excel actions, border styling with temporary floating preview badges, and focus verification.

**Key Functions**:
- [`execute()`](../src/websocket/service/ExcelCommandExecutor.ts#L48) - Main command execution entrypoint with focus auto-filtering
- [`resolveCommandName()`](../src/websocket/service/ExcelCommandExecutor.ts#L118) - Resolves legacy and snake_case command aliases
- [`set_border`](../src/websocket/service/ExcelCommandExecutor.ts#L362) - 6-color cycle palette calculation, border styling, and floating badge generation

**Dependencies**:
- ExcelPositionSlotManager
- Office.js / Excel JavaScript API
- Logger

##### 3.8.2 ExcelPositionSlotManager

**Location**: [`src/websocket/service/ExcelPositionSlotManager.ts`](../src/websocket/service/ExcelPositionSlotManager.ts)

**Purpose**: Manages zero-pollution position slots stored in `Office.context.document.settings`.

**Key Functions**:
- [`saveSlot()`](../src/websocket/service/ExcelPositionSlotManager.ts#L107) - Saves current range coordinates to document settings
- [`switchSlot()`](../src/websocket/service/ExcelPositionSlotManager.ts#L151) - Activates sheet and selects saved range
- [`clearSlot()`](../src/websocket/service/ExcelPositionSlotManager.ts#L202) - Removes position slot
- [`isFocused()`](../src/websocket/service/ExcelPositionSlotManager.ts#L320) - Checks whether Excel instance is currently focused

**Dependencies**:
- Office.js Document Settings API
- Logger

##### 3.8.3 BackgroundService

**Location**: [`src/websocket/service/BackgroundService.ts`](../src/websocket/service/BackgroundService.ts)

**Purpose**: Manages shared runtime background lifecycle, keep-alive tracking, and headless execution readiness.

**Key Functions**:
- [`start()`](../src/websocket/service/BackgroundService.ts#L30) - Starts background service loop
- [`stop()`](../src/websocket/service/BackgroundService.ts#L45) - Stops background service loop

**Dependencies**:
- Logger

##### 3.8.4 ExcelSpotlightService

**Location**: [`src/websocket/service/ExcelSpotlightService.ts`](../src/websocket/service/ExcelSpotlightService.ts)

**Purpose**: High-performance WPS-style Spotlight engine utilizing Excel native conditional formatting to deliver zero-pollution cursor crosshairs, anchor key column highlights, adaptive selection modes, cross-sheet ghost prevention, and whole-workbook crash detox.

**Key Functions**:
- [`getInstance()`](../src/websocket/service/ExcelSpotlightService.ts#L91) - Singleton accessor
- [`onSelectionChanged()`](../src/websocket/service/ExcelSpotlightService.ts#L206) - Evaluates selection changes with trailing execution queue (`pendingSelection`) for zero-stuck fast arrow navigation
- [`applySpotlightForSelection()`](../src/websocket/service/ExcelSpotlightService.ts#L330) - Applies prioritized rule collection and purges crosshairs from previously active worksheets
- [`purgeAllOrphanSpotlightsAcrossWorkbook()`](../src/websocket/service/ExcelSpotlightService.ts#L535) - Batched 3-phase atomic purge across all sheets to eliminate crash pollution
- [`isSpotlightFormula()`](../src/websocket/service/ExcelSpotlightService.ts#L571) - Strictly isolates spotlight formulas from user custom conditional formatting
- [`clearActiveSpotlight()`](../src/websocket/service/ExcelSpotlightService.ts#L488) - Safely removes spotlight rules from the active sheet without disturbing user formats

**Dependencies**:
- Office.js / Excel Conditional Formatting API
- Logger

### 4. Utilities

**Location**: [`src/utils/`](../src/utils/)

**Purpose**: Provides utility functions used across the add-in.

**Files**:
- [`uuid.ts`](../src/utils/uuid.ts) - UUID generation utility

**Key Functions**:
- [`generateUUID()`](../src/utils/uuid.ts) - Generates unique identifiers

**Dependencies**:
- None (self-contained)

### 5. Internationalization (i18n)

**Location**: [`src/i18n/`](../src/i18n/)

**Purpose**: Modular split-file translation architecture separating strongly typed dictionary schemas from localized strings, supporting multi-tier host detection and reactive UI updates.

**Files**:
- [`types.ts`](../src/i18n/types.ts) - Strongly typed `TranslationSchema` interface
- [`locales/zh-CN.ts`](../src/i18n/locales/zh-CN.ts) - Simplified Chinese dictionary
- [`locales/en-US.ts`](../src/i18n/locales/en-US.ts) - American English dictionary
- [`index.ts`](../src/i18n/index.ts) - Reactive `useI18n()` hook, locale detection, and storage persistence

**Key Functions**:
- [`useI18n()`](../src/i18n/index.ts#L48) - Returns reactive `t()` dictionary, `locale()`, and `setLocale()`
- [`detectInitialLocale()`](../src/i18n/index.ts#L14) - Hierarchy: localStorage -> Office.context -> navigator.language -> en-US fallback

**Dependencies**:
- SolidJS signals (`createSignal`)

## Data Flow

### Initialization Flow

```
Office.onReady()
    │
    ├──▶ Show task pane UI
    ├──▶ Initialize Logger
    ├──▶ Create ConnectionManager
    ├──▶ Setup Event Listeners
    ├──▶ Load Saved Settings
    └──▶ Auto-connect (if enabled)
```

### Connection Flow

```
User clicks Connect
    │
    ├──▶ Validate Settings
    ├──▶ Save Settings
    ├──▶ ConnectionManager.connect()
    ├──▶ WebSocketClient.connect()
    ├──▶ Send Register Message
    └──▶ Update UI (Connected State)
```

### Message Flow

```
WebSocket Message Received
    │
    ├──▶ WebSocketClient.handleMessage()
    ├──▶ KanataProtocol.parseMessage()
    ├──▶ MessageHandler.handleMessage()
    ├──▶ Route by Type
    └──▶ Update Log Viewer
```

## Dependencies

### External Dependencies

- **Office.js**: Microsoft 365 JavaScript API for Office integration
- **Fluent UI**: Microsoft's UI component library for consistent styling
- **TypeScript**: Type-safe JavaScript superset
- **Webpack**: Module bundler for production builds

### Internal Dependencies

```
taskpane.ts
    ├── ConnectionManager
    │   ├── WebSocketClient
    │   │   ├── KanataProtocol
    │   │   └── Logger
    │   ├── MessageHandler
    │   ├── SettingsStorage
    │   └── Logger
    └── Logger

WebSocketClient
    ├── KanataProtocol
    └── Logger

ConnectionManager
    ├── WebSocketClient
    ├── MessageHandler
    ├── SettingsStorage
    └── Logger
```

## Configuration

### Build Configuration

**Location**: [`webpack.config.js`](../webpack.config.js)

Configures the build process for the add-in, including:
- Entry points
- Output configuration
- Loaders for TypeScript and CSS
- Development server settings

### Environment Configuration

**Location**: [`env/`](../env/)

Contains environment-specific configuration files for different deployment targets.

### Manifest Configuration

**Location**: [`appPackage/manifest.json`](../appPackage/manifest.json)

Defines the add-in's capabilities, permissions, and metadata for Office.

## Testing

Currently, testing is done manually through:
1. Sideload the add-in in Excel
2. Test WebSocket connection to Kanata
3. Verify UI functionality
4. Check browser console for errors

Future enhancements should include automated unit and integration tests.

## Deployment

The add-in can be deployed locally for development or to Azure for production:

**Local Deployment**:
1. Run `npm run build`
2. Sideload using `npx office-addin-dev-settings sideload ./dist/manifest.json`

**Azure Deployment**:
1. Use Microsoft 365 Agents Toolkit to provision Azure resources
2. Deploy using the toolkit or CLI
3. Update manifest with production URL
4. Build and sideload

## Security Considerations

- WebSocket connections use ws:// (unencrypted) for local development
- Production should use wss:// (encrypted) for secure communication
- Settings are stored in browser localStorage (not encrypted)
- No authentication is currently implemented
- Future: Add authentication and encryption for production use
