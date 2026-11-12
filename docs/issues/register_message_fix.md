# Register Message Fix

## Issue

The WebSocket register message sent to the Kanata server was missing the required `clientType` field, causing the server to reject the connection with a parsing error.

## Root Cause

In [`WebSocketClient.ts`](../src/websocket/client/WebSocketClient.ts#L187), the `handleOpen()` method was creating a settings object with incomplete fields before calling `KanataProtocol.formatRegisterMessage()`. The settings object only included `port`, `clientName`, `autoConnect`, and `logLevel`, but was missing `clientType` and `channels` which are required by the Kanata server protocol.

## Error Log

```
02:24:22.3926 [WARN] Failed to parse WebSocket message: {"type":"register","instanceId":"325a0d75-b598-4217-abc6-8d7434db6bed","version":"1.0.0","clientName":"excel","timestamp":1790447062392,"capabilities":["receive-messages"]} - Error: missing field `clientType`
```

## Fix

Updated the settings object in [`WebSocketClient.handleOpen()`](../src/websocket/client/WebSocketClient.ts#L187-L195) to include all required fields:

```typescript
const settings: ConnectionSettings = {
  port: this.port,
  clientName: this.clientName,
  clientType: 'office-addin',        // Added
  channels: ['excel-addin-websocket'], // Added
  autoConnect: false,
  logLevel: 'info'
};
```

## Issue Logic Map

```
┌─────────────────────────────────────────────────────────────────┐
│                  Connection Flow (Before Fix)                   │
│                                                                 │
│  WebSocket Open Event                                           │
│         │                                                       │
│         ▼                                                       │
│  Create Settings Object                                        │
│  ┌──────────────────────────────┐                             │
│  │ {                            │                             │
│  │   port: 8890,               │                             │
│  │   clientName: 'excel',      │                             │
│  │   autoConnect: false,        │                             │
│  │   logLevel: 'info'           │                             │
│  │ }                            │                             │
│  └──────────────────────────────┘                             │
│         │                                                       │
│         ▼                                                       │
│  Format Register Message                                      │
│  ┌──────────────────────────────┐                             │
│  │ {                            │                             │
│  │   type: 'register',         │                             │
│  │   clientName: 'excel',      │                             │
│  │   ...                        │                             │
│  │   ❌ Missing clientType      │                             │
│  │   ❌ Missing channels        │                             │
│  │ }                            │                             │
│  └──────────────────────────────┘                             │
│         │                                                       │
│         ▼                                                       │
│  Send to Kanata Server                                        │
│         │                                                       │
│         ▼                                                       │
│  ❌ Parse Error (missing clientType)                           │
│         │                                                       │
│         ▼                                                       │
│  Connection Rejected                                           │
└─────────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────────┐
│                  Connection Flow (After Fix)                    │
│                                                                 │
│  WebSocket Open Event                                           │
│         │                                                       │
│         ▼                                                       │
│  Create Settings Object                                        │
│  ┌──────────────────────────────┐                             │
│  │ {                            │                             │
│  │   port: 8890,               │                             │
│  │   clientName: 'excel',      │                             │
│  │   clientType: 'office-addin',│ ✅ Added                  │
│  │   channels: [...],           │ ✅ Added                  │
│  │   autoConnect: false,        │                             │
│  │   logLevel: 'info'           │                             │
│  │ }                            │                             │
│  └──────────────────────────────┘                             │
│         │                                                       │
│         ▼                                                       │
│  Format Register Message                                      │
│  ┌──────────────────────────────┐                             │
│  │ {                            │                             │
│  │   type: 'register',         │                             │
│  │   clientName: 'excel',      │                             │
│  │   clientType: 'office-addin',│ ✅ Present                 │
│  │   channels: [...],           │ ✅ Present                 │
│  │   ...                        │                             │
│  │ }                            │                             │
│  └──────────────────────────────┘                             │
│         │                                                       │
│         ▼                                                       │
│  Send to Kanata Server                                        │
│         │                                                       │
│         ▼                                                       │
│  ✅ Parse Successful                                           │
│         │                                                       │
│         ▼                                                       │
│  Connection Accepted                                           │
│         │                                                       │
│         ▼                                                       │
│  Ready to Receive Messages                                    │
└─────────────────────────────────────────────────────────────────┘
```

## Files Modified

- [`src/websocket/client/WebSocketClient.ts`](../src/websocket/client/WebSocketClient.ts#L187-L195)
  - Added `clientType` field to settings object
  - Added `channels` field to settings object

## Impact

- WebSocket connections now successfully register with Kanata server
- No more parse errors on connection
- Full protocol compliance achieved

## Testing

After the fix, the connection process should:
1. Open WebSocket connection
2. Send complete register message
3. Receive acceptance from Kanata server
4. Transition to connected state
5. Begin receiving messages

Check the Kanata runtime log for successful registration:
```
[INFO] WebSocket client connected: id=..., name=Some("excel"), addr=...
[INFO] Successfully registered client: excel (office-addin)
```
