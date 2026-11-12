/**
 * WebSocket module exports
 */

export { SettingsStorage } from './storage/SettingsStorage';
export { ConnectionSettings, DEFAULT_SETTINGS, LogLevel } from './storage/ConnectionSettings';
export { Logger, LogEntry } from './logging/Logger';
export { WebSocketClient, WebSocketClientEvents } from './client/WebSocketClient';
export { MessageHandler, MessageHandlerEvents } from './client/MessageHandler';
export { ConnectionManager, ConnectionState, ConnectionManagerEvents } from './client/ConnectionManager';
export { BackgroundService } from './service/BackgroundService';
export { ExcelCommandExecutor } from './service/ExcelCommandExecutor';
export { ExcelPositionSlotManager, PositionSlot, HistoryEntry } from './service/ExcelPositionSlotManager';
export { KanataProtocol } from './protocol/KanataProtocol';
export {
  ParsedMessage,
  WebSocketRegister,
  WebSocketCommand,
  WebSocketResponse,
  WebSocketMessage,
  ConnectionInfo
} from './protocol/types';
