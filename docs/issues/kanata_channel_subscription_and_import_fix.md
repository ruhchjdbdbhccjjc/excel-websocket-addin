# Kanata Channel Subscription & Missing Imports Retrospective

This document explains the root cause, analysis, and fix for the WebSocket channel subscription failure and the runtime `ReferenceError` during message processing.

---

## 1. Issue Summary

1. **Kanata Dispatched to 0 Clients**:
   - Kanata logged: `Sending message to 0 clients on channel: excel_websocket_Offic_Add_in`.
   - The add-in appeared connected on port `8890`, but broadcast commands never reached the Excel executor.
2. **Runtime `ReferenceError` in MessageHandler**:
   - Incoming WebSocket frames crashed with: `ReferenceError: KanataProtocol is not defined`.

---

## 2. Root Cause Analysis

1. **Channel Mismatch**:
   - In Kanata's Rust WebSocket server (`websocket_server.rs`), message routing to a channel checks:
     `c.channels.contains(&channel)`.
   - The add-in had hardcoded `channels: ['excel-addin-websocket']` or multi-channel arrays, while Kanata configuration sent messages exclusively to `excel_websocket_Offic_Add_in`.
   - In addition, the add-in did not send an explicit `{"type": "subscribe", "channels": [...]}` payload upon completing the connection handshake.
2. **Missing ES Module Import**:
   - `src/websocket/client/MessageHandler.ts` invoked `KanataProtocol.parseMessage(data)`, but the file lacked an `import { KanataProtocol } from '../protocol/KanataProtocol'` statement, causing an unhandled JavaScript exception in the browser runtime.

---

## 3. Issue Logic Map (ASCII Art)

```
===================================================================================
                       FAILURE PATH vs. RESOLUTION PATH
===================================================================================

[FAILURE PATH]
Kanata Macro triggers: (ws-send-channel "excel_websocket_Offic_Add_in" ...)
       |
       v
Kanata Server checks client channels
       |
       +---> Client registered with ["excel-addin-websocket"]
       +---> Channel mismatch!
       |
       v
Kanata logs: "Sending message to 0 clients on channel: excel_websocket_Offic_Add_in"
       |
       +---> [IF message did arrive]:
               MessageHandler.handleMessage() calls KanataProtocol.parseMessage()
               |
               v
               CRASH: ReferenceError: KanataProtocol is not defined!


-----------------------------------------------------------------------------------

[SUCCESS PATH (FIXED)]
Add-in connects to ws://localhost:8890
       |
       v
[WebSocketClient.handleOpen]
       |
       +---> Sends register: channels = ["excel_websocket_Offic_Add_in"]
       +---> Sends explicit subscribe: {"type": "subscribe", "channels": ["excel_websocket_Offic_Add_in"]}
       v
Kanata Server registers client in "excel_websocket_Offic_Add_in" channel map
       |
       v
Kanata Macro triggers: (ws-send-channel "excel_websocket_Offic_Add_in" ...)
       |
       v
Kanata Server sends message to 1 client on channel excel_websocket_Offic_Add_in
       |
       v
[MessageHandler.handleMessage]
       |
       +---> KanataProtocol.parseMessage() imported correctly
       +---> Parses envelope { type: "command", command: "...", params: { ... } }
       v
[ExcelCommandExecutor.execute]
       |
       +---> Focus verified -> Excel action executed successfully!
```

---

## 4. Code Changes & Modified Functions

### A. Fix Kanata Channel Subscription
- **File**: [`src/websocket/client/WebSocketClient.ts`](../src/websocket/client/WebSocketClient.ts#L202-L221)
- **File**: [`src/websocket/storage/ConnectionSettings.ts`](../src/websocket/storage/ConnectionSettings.ts#L26-L32)
- **Functions Modified**:
  - `[WebSocketClient.handleOpen](../src/websocket/client/WebSocketClient.ts#L202-L221)`: Dispatches standardized channel registration and sends an explicit `type: 'subscribe'` message.
  - `[DEFAULT_SETTINGS](../src/websocket/storage/ConnectionSettings.ts#L26-L32)`: Canonical default set to `['excel_websocket_Offic_Add_in']`.

### B. Fix Missing Imports in MessageHandler
- **File**: [`src/websocket/client/MessageHandler.ts`](../src/websocket/client/MessageHandler.ts#L6-L8)
- **Code Added**:
  ```typescript
  import { KanataProtocol } from '../protocol/KanataProtocol';
  import { ParsedMessage } from '../protocol/types';
  ```
- **Function Affected**:
  - `[MessageHandler.handleMessage](../src/websocket/client/MessageHandler.ts#L28-L35)` now parses without runtime exceptions.
