# WebSocket Integration Details

## Overview

The WebSocket integration provides real-time bidirectional communication between the Excel add-in and the Kanata keyboard layout manager. This feature enables the add-in to receive keyboard events and commands from Kanata, allowing for advanced automation and integration scenarios.

## Feature Logic Map

```
┌─────────────────────────────────────────────────────────────────┐
│                    Excel Add-in Task Pane                        │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐          │
│  │ Connection   │  │   Settings   │  │   Log Viewer │          │
│  │   Controls   │  │   Form       │  │   Display    │          │
│  └──────┬───────┘  └──────┬───────┘  └──────┬───────┘          │
│         │                 │                 │                    │
│         └─────────────────┴─────────────────┘                    │
│                           │                                      │
│                           ▼                                      │
│              ┌──────────────────────┐                           │
│              │  ConnectionManager    │                           │
│              │  (State Management)   │                           │
│              └──────────┬───────────┘                           │
│                         │                                       │
│         ┌───────────────┼───────────────┐                      │
│         │               │               │                      │
│         ▼               ▼               ▼                      │
│  ┌──────────────┐ ┌──────────────┐ ┌──────────────┐          │
│  │WebSocketClient│ │SettingsStorage│ │MessageHandler│          │
│  │(Connection)  │ │(Persistence)  │ │(Parsing)     │          │
│  └──────┬───────┘ └──────────────┘ └──────┬───────┘          │
│         │                                 │                    │
│         └────────────────┬────────────────┘                    │
│                          │                                     │
│                          ▼                                     │
│               ┌──────────────────┐                            │
│               │  Kanata Protocol  │                            │
│               │  (Message Format) │                            │
│               └─────────┬────────┘                            │
│                         │                                     │
│                         ▼                                     │
│              ┌──────────────────────┐                           │
│              │   Kanata Server      │                           │
│              │   (ws://localhost)   │                           │
│              └──────────────────────┘                           │
└─────────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────────┐
│                      Message Flow                                │
│                                                                 │
│  Registration Flow:                                             │
│  ┌─────────┐    ┌─────────┐    ┌─────────┐    ┌─────────┐     │
│  │ Connect │───▶│ Register│───▶│  Accept │───▶│  Ready  │     │
│  └─────────┘    └─────────┘    └─────────┘    └─────────┘     │
│                                                                 │
│  Message Handling Flow:                                         │
│  ┌─────────┐    ┌─────────┐    ┌─────────┐    ┌─────────┐     │
│  │ Receive │───▶│  Parse  │───▶│ Route   │───▶│ Execute │     │
│  └─────────┘    └─────────┘    └─────────┘    └─────────┘     │
│                                                                 │
│  Reconnection Flow:                                             │
│  ┌─────────┐    ┌─────────┐    ┌─────────┐                   │
│  │ Disconnect│───▶│  Wait   │───▶│ Reconnect│                   │
│  └─────────┘    └─────────┘    └─────────┘                   │
│     │              │              │                           │
│     │              │              └──▶ Max Attempts?           │
│     │              │                   │                       │
│     │              │                   ▼                       │
│     │              │              ┌─────────┐                │
│     │              │              │   Stop  │                │
│     │              │              └─────────┘                │
└─────────────────────────────────────────────────────────────────┘
```

## Key Components

### 1. ConnectionManager

The `ConnectionManager` is the high-level orchestrator that manages connection state, settings, and event handlers.

**Key Functions:**
- [`connect()`](../src/websocket/client/ConnectionManager.ts#L97) - Initiates WebSocket connection
- [`disconnect()`](../src/websocket/client/ConnectionManager.ts#L107) - Closes WebSocket connection
- [`getSettings()`](../src/websocket/client/ConnectionManager.ts#L49) - Retrieves connection settings
- [`saveSettings()`](../src/websocket/client/ConnectionManager.ts#L78) - Persists connection settings
- [`getConnectionState()`](../src/websocket/client/ConnectionManager.ts#L113) - Returns current connection state

**State Management:**
- Disconnected: Initial state, no connection
- Connecting: Connection attempt in progress
- Connected: Successfully connected and registered
- Error: Connection error occurred

### 2. WebSocketClient

The `WebSocketClient` handles the low-level WebSocket connection management, including reconnection logic and heartbeat.

**Key Functions:**
- [`connect()`](../src/websocket/client/WebSocketClient.ts#L54) - Establishes WebSocket connection
- [`disconnect()`](../src/websocket/client/WebSocketClient.ts#L85) - Terminates WebSocket connection
- [`send()`](../src/websocket/client/WebSocketClient.ts#L108) - Sends JSON message
- [`sendRaw()`](../src/websocket/client/WebSocketClient.ts#L123) - Sends raw string message
- [`isConnected()`](../src/websocket/client/WebSocketClient.ts#L140) - Checks connection status

**Reconnection Logic:**
- Exponential backoff: 1s, 2s, 4s, 8s, ... max 30s
- Maximum 10 reconnection attempts
- Manual disconnect stops reconnection
- Automatic heartbeat every 30 seconds

### 3. MessageHandler

The `MessageHandler` parses incoming messages and routes them to appropriate handlers.

**Key Functions:**
- [`handleMessage()`](../src/websocket/client/MessageHandler.ts#L20) - Main message entry point
- [`parseStringMessage()`](../src/websocket/client/MessageHandler.ts#L46) - Parses string messages
- [`parseArrayMessage()`](../src/websocket/client/MessageHandler.ts#L60) - Parses array messages
- [`parseObjectMessage()`](../src/websocket/client/MessageHandler.ts#L78) - Parses object messages

**Supported Message Types:**
- String: Simple text messages
- Array: JSON arrays with commands/events
- Object: JSON objects with type information

### 4. KanataProtocol

The `KanataProtocol` handles Kanata-specific message formatting and parsing.

**Key Functions:**
- [`parseMessage()`](../src/websocket/protocol/KanataProtocol.ts#L18) - Parses incoming messages
- [`formatRegisterMessage()`](../src/websocket/protocol/KanataProtocol.ts#L44) - Creates registration message
- [`formatHeartbeat()`](../src/websocket/protocol/KanataProtocol.ts#L66) - Creates heartbeat message

**Register Message Format:**
```json
{
  "type": "register",
  "instanceId": "uuid",
  "version": "1.0.0",
  "clientType": "office-addin",
  "clientName": "excel",
  "timestamp": 1234567890,
  "capabilities": ["receive-messages"],
  "channels": ["excel-addin-websocket"]
}
```

### 5. Logger

The `Logger` provides log-level based logging with filtering capabilities.

**Key Functions:**
- [`debug()`](../src/websocket/logging/Logger.ts#L62) - Debug level logging
- [`info()`](../src/websocket/logging/Logger.ts#L66) - Info level logging
- [`warn()`](../src/websocket/logging/Logger.ts#L70) - Warning level logging
- [`error()`](../src/websocket/logging/Logger.ts#L74) - Error level logging
- [`setLevel()`](../src/websocket/logging/Logger.ts#L23) - Sets minimum log level
- [`getLogs()`](../src/websocket/logging/Logger.ts#L78) - Retrieves log entries

**Log Levels:**
- Debug: Detailed diagnostic information
- Info: General informational messages
- Warn: Warning messages for potential issues
- Error: Error messages for failures

### 6. SettingsStorage

The `SettingsStorage` manages persistent storage of connection settings using browser storage APIs.

**Key Functions:**
- [`get()`](../src/websocket/storage/SettingsStorage.ts#L18) - Retrieves single setting
- [`set()`](../src/websocket/storage/SettingsStorage.ts#L33) - Saves single setting
- [`getMultiple()`](../src/websocket/storage/SettingsStorage.ts#L49) - Retrieves multiple settings
- [`setMultiple()`](../src/websocket/storage/SettingsStorage.ts#L66) - Saves multiple settings

**Storage Keys:**
- `ws_port`: WebSocket server port
- `ws_client_name`: Client identifier
- `ws_client_type`: Client type classification
- `ws_channels`: Subscribed channels
- `ws_auto_connect`: Auto-connect on load flag

## Connection Flow

### 1. Initialization

```
Office.onReady()
    │
    ▼
initializeWebSocket()
    │
    ├──▶ Create Logger
    ├──▶ Create ConnectionManager
    ├──▶ Setup Event Listeners
    ├──▶ Load Saved Settings
    ├──▶ Update UI with Settings
    └──▶ Auto-connect (if enabled)
```

### 2. Connection Process

```
User clicks "Connect"
    │
    ▼
Validate Settings
    │
    ├──▶ Port (1-65535)
    ├──▶ Client Name (non-empty)
    └──▶ Channels (comma-separated)
    │
    ▼
Save Settings
    │
    ▼
ConnectionManager.connect()
    │
    ▼
WebSocketClient.connect()
    │
    ├──▶ Create WebSocket
    ├──▶ Setup Event Handlers
    ├──▶ Wait for Open Event
    └──▶ Send Register Message
    │
    ▼
Update UI (Connected State)
```

### 3. Message Handling

```
WebSocket Message Received
    │
    ▼
WebSocketClient.handleMessage()
    │
    ▼
KanataProtocol.parseMessage()
    │
    ├──▶ Try JSON Parse
    │   ├──▶ Array? → Array Message
    │   ├──▶ Object? → Object Message
    │   └──▶ String? → String Message
    │
    ▼
MessageHandler.handleMessage()
    │
    ├──▶ Route by Type
    │   ├──▶ Command → Command Handler
    │   ├──▶ Event → Event Handler
    │   └──▶ Other → Log Only
    │
    ▼
Update Log Viewer
```

### 4. Reconnection Process

```
Connection Lost
    │
    ▼
Check Manual Disconnect?
    │
    ├──▶ Yes → Stop Reconnection
    └──▶ No → Continue
    │
    ▼
Check Max Attempts (10)
    │
    ├──▶ Reached → Stop Reconnection
    └──▶ Not Reached → Continue
    │
    ▼
Calculate Delay (Exponential Backoff)
    │
    ├──▶ Attempt 1: 1s
    ├──▶ Attempt 2: 2s
    ├──▶ Attempt 3: 4s
    └──▶ ... Max 30s
    │
    ▼
Schedule Reconnection
    │
    ▼
Attempt Reconnect
    │
    ├──▶ Success → Reset Attempts
    └──▶ Failed → Increment Attempts
```

## Configuration

### Default Settings

```typescript
{
  port: 8890,
  clientName: 'excel',
  clientType: 'office-addin',
  channels: ['excel-addin-websocket'],
  autoConnect: false,
  logLevel: 'info'
}
```

### UI Configuration

The task pane provides a configuration form with:
- Port input (number, 1-65535)
- Client name input (text)
- Client type input (text)
- Channels input (comma-separated text)
- Auto-connect checkbox

## Example Configs

No separate example config files are needed for this feature. Configuration is done through the task pane UI and stored in browser localStorage.

## Testing

No automated test files have been created yet. Testing is currently done manually through:
1. Sideload the add-in in Excel
2. Configure connection settings
3. Connect to Kanata server
4. Verify connection status
5. Test message reception

## Known Issues

See [KNOWN_ISSUES.md](../KNOWN_ISSUES.md) for known issues and their resolutions.

## Dependencies

- Office.js: Microsoft 365 JavaScript API
- WebSocket API: Browser WebSocket implementation
- LocalStorage API: Browser persistent storage

## Future Enhancements

Potential improvements:
- Add message encryption
- Implement message queue for offline scenarios
- Add support for binary messages
- Implement message acknowledgment
- Add performance metrics
- Create automated test suite
