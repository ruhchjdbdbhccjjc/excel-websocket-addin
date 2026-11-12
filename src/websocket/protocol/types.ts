/**
 * Protocol types for WebSocket communication
 */

export interface ParsedMessage {
  type: 'string' | 'array' | 'object';
  data: string | unknown[] | Record<string, unknown>;
  raw: string;
}

export interface WebSocketRegister {
  type: 'register';
  instanceId: string;
  version: string;
  clientType: string;
  clientName: string;
  timestamp: number;
  capabilities: string[];
  channels: string[];
}

export interface WebSocketCommand {
  type: 'command';
  id: string;
  command: string;
  params?: Record<string, unknown>;
}

export interface WebSocketResponse {
  type: 'response';
  id: string;
  status: 'success' | 'error';
  result?: unknown;
  error?: string;
}

export interface CommandResult {
  ok: boolean;
  data?: unknown;
  error?: string;
}

export interface WebSocketMessage {
  type?: 'command' | 'response' | 'register' | 'heartbeat';
  id?: string;
  command?: string;
  params?: Record<string, unknown>;
  status?: 'success' | 'error';
  result?: unknown;
  error?: string;
  data?: string | unknown[];
}

export interface ConnectionInfo {
  connected: boolean;
  port: number;
  clientName: string;
  instanceId: string;
  lastConnected?: Date;
  lastMessage?: Date;
}
