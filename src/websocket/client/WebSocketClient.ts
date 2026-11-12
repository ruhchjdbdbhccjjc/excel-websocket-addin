/**
 * WebSocketClient - Core WebSocket connection management
 * Handles connection lifecycle, message send/receive, and reconnection logic
 */

import { ConnectionInfo, WebSocketMessage } from '../protocol/types';
import { KanataProtocol } from '../protocol/KanataProtocol';
import { ConnectionSettings } from '../storage/ConnectionSettings';
import { generateUUID } from '../../utils/uuid';
import { Logger } from '../logging/Logger';

export interface WebSocketClientEvents {
  connected: (info: ConnectionInfo) => void;
  disconnected: (reason?: string) => void;
  message: (message: WebSocketMessage, parsed: any) => void;
  error: (error: Error) => void;
}

export class WebSocketClient {
  private ws: WebSocket | null = null;
  private port: number = 8890;
  private clientName: string = 'excel';
  private instanceId: string = generateUUID();
  private eventHandlers: Partial<WebSocketClientEvents> = {};
  private reconnectTimer: NodeJS.Timeout | null = null;
  private reconnectAttempts: number = 0;
  private maxReconnectAttempts: number = Infinity;
  private reconnectDelay: number = 2000; // Start with 2 seconds
  private maxReconnectDelay: number = 5000; // Cap at 5 seconds
  private heartbeatTimer: NodeJS.Timeout | null = null;
  private heartbeatInterval: number = 30000; // 30 seconds
  private isManualDisconnect: boolean = false;
  private channels: string[] = ['excel_websocket_Offic_Add_in'];
  private clientType: string = 'office-addin';
  private logger: Logger;

  constructor(logger?: Logger) {
    this.logger = logger || new Logger();
  }

  /**
   * Connect to the WebSocket server
   * @param port - The port to connect to
   * @param clientName - The client name for identification
   * @param channels - The channels to subscribe to
   * @param clientType - The client type identifier
   * @returns Promise when connection is established
   */
  async connect(port: number, clientName: string, channels?: string[], clientType?: string): Promise<void> {
    this.port = port;
    this.clientName = clientName;
    if (channels && channels.length > 0) {
      this.channels = [...channels];
    }
    if (clientType) {
      this.clientType = clientType;
    }
    this.isManualDisconnect = false;
    this.reconnectAttempts = 0;

    const isSecurePort = port === 8891;
    const protocol = isSecurePort ? 'wss://' : 'ws://';
    const url = `${protocol}localhost:${port}`;
    this.logger.info(`Connecting to WebSocket server at ${url}`);

    try {
      this.ws = new WebSocket(url);
      this.setupEventHandlers();
      
      await new Promise<void>((resolve, reject) => {
        const onOpen = () => {
          this.ws?.removeEventListener('open', onOpen);
          this.ws?.removeEventListener('error', onError);
          this.logger.info('WebSocket connection opened successfully');
          resolve();
        };
        const onError = (error: Event) => {
          this.ws?.removeEventListener('open', onOpen);
          this.ws?.removeEventListener('error', onError);
          this.logger.error('WebSocket connection failed');
          reject(new Error('WebSocket connection failed'));
        };
        this.ws?.addEventListener('open', onOpen);
        this.ws?.addEventListener('error', onError);
      });

      this.logger.info('WebSocket connection established');
    } catch (error) {
      this.logger.error('Failed to connect to WebSocket server', error);
      throw error;
    }
  }

  /**
   * Disconnect from the WebSocket server
   */
  disconnect(): void {
    this.isManualDisconnect = true;
    this.clearReconnectTimer();
    this.clearHeartbeatTimer();

    if (this.ws) {
      this.logger.info('Disconnecting from WebSocket server');
      this.ws.close(1000, 'Manual disconnect');
      this.ws = null;
    }
  }

  /**
   * Send a message to the server
   * @param message - The message to send
   */
  send(message: WebSocketMessage): void {
    if (!this.ws || this.ws.readyState !== WebSocket.OPEN) {
      this.logger.warn('Cannot send message: WebSocket not connected');
      return;
    }

    try {
      const messageStr = JSON.stringify(message);
      this.ws.send(messageStr);
      this.logger.debug('Message sent', message);
    } catch (error) {
      this.logger.error('Failed to send message', error);
    }
  }

  /**
   * Send a raw string message
   * @param data - The string data to send
   */
  sendRaw(data: string): void {
    if (!this.ws || this.ws.readyState !== WebSocket.OPEN) {
      this.logger.warn('Cannot send message: WebSocket not connected');
      return;
    }

    try {
      this.ws.send(data);
      this.logger.debug('Raw message sent', data);
    } catch (error) {
      this.logger.error('Failed to send raw message', error);
    }
  }

  /**
   * Register an event handler
   * @param event - The event name
   * @param handler - The handler function
   */
  on<K extends keyof WebSocketClientEvents>(
    event: K,
    handler: WebSocketClientEvents[K]
  ): void {
    this.eventHandlers[event] = handler;
  }

  /**
   * Check if the client is connected
   * @returns Whether the client is connected
   */
  isConnected(): boolean {
    return this.ws !== null && this.ws.readyState === WebSocket.OPEN;
  }

  /**
   * Get connection information
   * @returns Connection info object
   */
  getConnectionInfo(): ConnectionInfo {
    return {
      connected: this.isConnected(),
      port: this.port,
      clientName: this.clientName,
      instanceId: this.instanceId
    };
  }

  /**
   * Setup WebSocket event handlers
   */
  private setupEventHandlers(): void {
    if (!this.ws) return;

    this.ws.onopen = () => this.handleOpen();
    this.ws.onclose = (event) => this.handleClose(event);
    this.ws.onerror = (error) => this.handleError(error);
    this.ws.onmessage = (event) => this.handleMessage(event);
  }

  /**
   * Handle WebSocket open event
   */
  private handleOpen(): void {
    this.logger.info('WebSocket connection opened');
    this.reconnectAttempts = 0;
    this.reconnectDelay = 1000;

    // Fallback to default canonical channel if empty
    if (!this.channels || this.channels.length === 0) {
      this.channels = ['excel_websocket_Offic_Add_in'];
    }

    // Send register message with configured channels
    const settings: ConnectionSettings = {
      port: this.port,
      clientName: this.clientName,
      clientType: this.clientType,
      channels: this.channels,
      autoConnect: false,
      logLevel: 'info'
    };
    const registerMessage = KanataProtocol.formatRegisterMessage(settings);
    this.sendRaw(registerMessage);

    // Also send explicit subscribe message so Kanata immediately maps this client
    const subscribeMessage = JSON.stringify({
      type: 'subscribe',
      channels: this.channels
    });
    this.sendRaw(subscribeMessage);
    this.logger.info('Registered and subscribed to Kanata channels:', this.channels);

    // Start heartbeat
    this.startHeartbeat();

    // Emit connected event
    const info = this.getConnectionInfo();
    info.lastConnected = new Date();
    this.eventHandlers.connected?.(info);
  }

  /**
   * Handle WebSocket close event
   * @param event - The close event
   */
  private handleClose(event: CloseEvent): void {
    this.logger.info(`WebSocket connection closed: ${event.code} - ${event.reason}`);
    this.clearHeartbeatTimer();

    const info = this.getConnectionInfo();
    this.eventHandlers.disconnected?.(event.reason);

    // Attempt reconnection if not manual disconnect
    if (!this.isManualDisconnect && this.reconnectAttempts < this.maxReconnectAttempts) {
      this.scheduleReconnect();
    }
  }

  /**
   * Handle WebSocket error event
   * @param error - The error event
   */
  private handleError(error: Event): void {
    this.logger.error('WebSocket error occurred', error);
    this.eventHandlers.error?.(new Error('WebSocket error'));
  }

  /**
   * Handle WebSocket message event
   * @param event - The message event
   */
  private handleMessage(event: MessageEvent): void {
    try {
      const data = event.data as string;
      this.logger.debug('Message received:', data);

      const parsed = KanataProtocol.parseMessage(data);
      this.eventHandlers.message?.(data as any, parsed);

      // Update last message time
      const info = this.getConnectionInfo();
      info.lastMessage = new Date();
    } catch (error) {
      this.logger.error('Failed to handle message', error);
    }
  }

  /**
   * Schedule a reconnection attempt
   */
  public scheduleReconnect(): void {
    if (this.isManualDisconnect) return;
    this.clearReconnectTimer();

    this.reconnectAttempts++;
    const delay = Math.min(
      this.reconnectDelay * Math.min(this.reconnectAttempts, 2),
      this.maxReconnectDelay
    );

    this.logger.info(`Reconnecting to Kanata WebSocket in ${delay}ms (attempt #${this.reconnectAttempts})...`);

    this.reconnectTimer = setTimeout(async () => {
      try {
        await this.connect(this.port, this.clientName);
      } catch (error) {
        this.logger.debug('Reconnection attempt failed, scheduling next retry');
        this.scheduleReconnect();
      }
    }, delay);
  }

  /**
   * Clear the reconnection timer
   */
  private clearReconnectTimer(): void {
    if (this.reconnectTimer) {
      clearTimeout(this.reconnectTimer);
      this.reconnectTimer = null;
    }
  }

  /**
   * Start the heartbeat timer
   */
  private startHeartbeat(): void {
    this.clearHeartbeatTimer();
    this.heartbeatTimer = setInterval(() => {
      if (this.isConnected()) {
        const heartbeat = KanataProtocol.formatHeartbeat();
        this.sendRaw(heartbeat);
      }
    }, this.heartbeatInterval);
  }

  /**
   * Clear the heartbeat timer
   */
  private clearHeartbeatTimer(): void {
    if (this.heartbeatTimer) {
      clearInterval(this.heartbeatTimer);
      this.heartbeatTimer = null;
    }
  }
}