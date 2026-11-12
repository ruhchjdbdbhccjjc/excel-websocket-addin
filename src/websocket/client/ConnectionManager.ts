/**
 * ConnectionManager - High-level connection state management
 * Manages WebSocket client, settings, and provides unified interface to UI
 */

import { WebSocketClient } from './WebSocketClient';
import { MessageHandler } from './MessageHandler';
import { SettingsStorage } from '../storage/SettingsStorage';
import { ConnectionSettings, DEFAULT_SETTINGS } from '../storage/ConnectionSettings';
import { ConnectionInfo, ParsedMessage } from '../protocol/types';
import { Logger } from '../logging/Logger';

export enum ConnectionState {
  Disconnected = 'disconnected',
  Connecting = 'connecting',
  Connected = 'connected',
  Error = 'error'
}

export interface ConnectionManagerEvents {
  stateChange: (state: ConnectionState) => void;
  messageReceived: (parsed: ParsedMessage) => void;
  logEntry: (entry: any) => void;
}

export class ConnectionManager {
  private storage: SettingsStorage;
  private client: WebSocketClient;
  private messageHandler: MessageHandler;
  private currentState: ConnectionState = ConnectionState.Disconnected;
  private currentSettings: ConnectionSettings = DEFAULT_SETTINGS;
  private eventHandlers: { [K in keyof ConnectionManagerEvents]?: Array<ConnectionManagerEvents[K]> } = {};
  private logger: Logger;

  constructor(logger?: Logger) {
    this.logger = logger || new Logger();
    this.storage = new SettingsStorage();
    this.client = new WebSocketClient(this.logger);
    this.messageHandler = new MessageHandler();

    this.setupClientHandlers();
    this.setupMessageHandlers();
  }

  /**
   * Get current connection settings
   * @returns Connection settings
   */
  async getSettings(): Promise<ConnectionSettings> {
    try {
      const settings = await this.storage.getMultiple<Record<string, unknown>>(
        ['ws_port', 'ws_client_name', 'ws_client_type', 'ws_channels', 'ws_auto_connect'],
        {
          ws_port: DEFAULT_SETTINGS.port,
          ws_client_name: DEFAULT_SETTINGS.clientName,
          ws_client_type: DEFAULT_SETTINGS.clientType,
          ws_channels: DEFAULT_SETTINGS.channels,
          ws_auto_connect: DEFAULT_SETTINGS.autoConnect
        }
      );

      let channels = (settings.ws_channels as string[]) || DEFAULT_SETTINGS.channels;
      if (!Array.isArray(channels) || channels.length === 0) {
        channels = DEFAULT_SETTINGS.channels;
      }

      this.currentSettings = {
        port: (settings.ws_port as number) || DEFAULT_SETTINGS.port,
        clientName: (settings.ws_client_name as string) || DEFAULT_SETTINGS.clientName,
        clientType: (settings.ws_client_type as string) || DEFAULT_SETTINGS.clientType,
        channels: channels,
        autoConnect: settings.ws_auto_connect !== undefined && settings.ws_auto_connect !== null
          ? (settings.ws_auto_connect as boolean)
          : DEFAULT_SETTINGS.autoConnect,
        logLevel: 'info' as any // Default value, not used
      };

      return this.currentSettings;
    } catch (error) {
      this.logger.error('Failed to load settings, using defaults', error);
      return DEFAULT_SETTINGS;
    }
  }

  /**
   * Save connection settings
   * @param settings - Settings to save
   */
  async saveSettings(settings: ConnectionSettings): Promise<void> {
    try {
      await this.storage.setMultiple({
        ws_port: settings.port,
        ws_client_name: settings.clientName,
        ws_client_type: settings.clientType,
        ws_channels: settings.channels,
        ws_auto_connect: settings.autoConnect
      });

      this.currentSettings = settings;
      this.logger.info('Settings saved', settings);
    } catch (error) {
      this.logger.error('Failed to save settings', error);
      throw error;
    }
  }

  /**
   * Connect to the WebSocket server
   */
  async connect(): Promise<void> {
    const settings = await this.getSettings();
    this.setState(ConnectionState.Connecting);

    try {
      await this.client.connect(settings.port, settings.clientName, settings.channels, settings.clientType);
    } catch (error) {
      this.setState(ConnectionState.Disconnected);
      this.logger.warn('Initial connection failed, auto-reconnect is retrying in background...');
      this.client.scheduleReconnect();
    }
  }

  /**
   * Disconnect from the WebSocket server
   */
  disconnect(): void {
    this.client.disconnect();
    this.setState(ConnectionState.Disconnected);
  }

  /**
   * Send a message to the WebSocket server
   * @param message Message payload to send
   */
  send(message: WebSocketResponse | WebSocketMessage): void {
    this.client.send(message as any);
  }

  /**
   * Get current connection state
   * @returns Current connection state
   */
  getConnectionState(): ConnectionState {
    return this.currentState;
  }

  /**
   * Get connection info
   * @returns Connection info
   */
  getConnectionInfo(): ConnectionInfo {
    return this.client.getConnectionInfo();
  }

  /**
   * Register an event handler
   * @param event - The event name
   * @param handler - The handler function
   */
  on<K extends keyof ConnectionManagerEvents>(
    event: K,
    handler: ConnectionManagerEvents[K]
  ): void {
    if (!this.eventHandlers[event]) {
      this.eventHandlers[event] = [];
    }
    (this.eventHandlers[event] as Array<ConnectionManagerEvents[K]>).push(handler);
  }

  /**
   * Get logs (for compatibility)
   * @returns Empty array (no longer using log storage)
   */
  getLogs(): any[] {
    return [];
  }

  /**
   * Clear logs (for compatibility)
   */
  clearLogs(): void {
    // No-op - using console.log instead
  }

  /**
   * Cleanup resources
   */
  dispose(): void {
    this.disconnect();
  }

  /**
   * Setup WebSocket client event handlers
   */
  private setupClientHandlers(): void {
    this.client.on('connected', (info: ConnectionInfo) => {
      this.logger.info('Client connected', info);
      this.setState(ConnectionState.Connected);
    });

    this.client.on('disconnected', (reason?: string) => {
      this.logger.info('Client disconnected', reason);
      this.setState(ConnectionState.Disconnected);
    });

    this.client.on('error', (error: Error) => {
      this.logger.error('Client error', error);
      this.setState(ConnectionState.Error);
    });

    this.client.on('message', (message: any, parsed: ParsedMessage) => {
      this.messageHandler.handleMessage(message);
    });
  }

  /**
   * Setup message handler event handlers
   */
  private setupMessageHandlers(): void {
    this.messageHandler.on('messageReceived', (parsed: ParsedMessage) => {
      const handlers = this.eventHandlers.messageReceived;
      if (handlers) {
        handlers.forEach((h) => {
          try {
            h(parsed);
          } catch (err) {
            this.logger.error('Error in messageReceived handler:', err);
          }
        });
      }
    });

    this.messageHandler.on('commandReceived', (command: string, params?: Record<string, unknown>) => {
      this.logger.info('Command received', { command, params });
    });

    this.messageHandler.on('eventReceived', (event: string, data?: unknown) => {
      this.logger.info('Event received', { event, data });
    });
  }

  /**
   * Set the current state and notify listeners
   * @param state - The new state
   */
  private setState(state: ConnectionState): void {
    if (this.currentState !== state) {
      this.currentState = state;
      this.logger.info('Connection state changed', state);
      const handlers = this.eventHandlers.stateChange;
      if (handlers) {
        handlers.forEach((h) => {
          try {
            h(state);
          } catch (err) {
            this.logger.error('Error in stateChange handler:', err);
          }
        });
      }
    }
  }
}
