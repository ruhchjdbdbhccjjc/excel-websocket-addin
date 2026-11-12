/**
 * MessageHandler - Protocol message handling
 * Parses incoming messages and routes them to appropriate handlers
 */

import { KanataProtocol } from '../protocol/KanataProtocol';
import { ParsedMessage } from '../protocol/types';
import { Logger } from '../logging/Logger';

export interface MessageHandlerEvents {
  messageReceived: (parsed: ParsedMessage) => void;
  commandReceived: (command: string, params?: Record<string, unknown>) => void;
  eventReceived: (event: string, data?: unknown) => void;
}

export class MessageHandler {
  private eventHandlers: Partial<MessageHandlerEvents> = {};
  private logger: Logger;

  constructor(logger?: Logger) {
    this.logger = logger || Logger.getInstance();
  }

  /**
   * Handle an incoming message
   * @param data - The raw message data
   */
  handleMessage(data: string): void {
    try {
      const parsed = KanataProtocol.parseMessage(data);
      this.logger.debug('Message parsed successfully', parsed, 'WS-IN');

      // Emit generic message event
      this.eventHandlers.messageReceived?.(parsed);

      // Route based on message type
      switch (parsed.type) {
        case 'string':
          this.handleStringMessage(parsed.data as string);
          break;
        case 'array':
          this.handleArrayMessage(parsed.data as unknown[]);
          break;
        case 'object':
          this.handleObjectMessage(parsed.data as Record<string, unknown>);
          break;
      }
    } catch (error) {
      this.logger.error('Failed to handle incoming message', error, 'WS-IN');
    }
  }

  /**
   * Handle a string message
   * @param data - The string data
   */
  private handleStringMessage(data: string): void {
    this.logger.info('String message received', data, 'WS-IN');

    // Check if it's a command in the format "command:params"
    if (data.includes(':')) {
      const [command, ...params] = data.split(':');
      const paramValue = params.join(':');
      this.eventHandlers.commandReceived?.(command, { value: paramValue });
    } else {
      this.eventHandlers.commandReceived?.(data);
    }
  }

  /**
   * Handle an array message
   * @param data - The array data
   */
  private handleArrayMessage(data: unknown[]): void {
    this.logger.info('Array message received', data, 'WS-IN');

    // Kanata array format: ["event", "event-name", "key", "value"]
    if (data.length >= 2 && data[0] === 'event') {
      const eventName = data[1] as string;
      const eventData = data.slice(2);
      this.eventHandlers.eventReceived?.(eventName, eventData);
    }
  }

  /**
   * Handle an object message
   * @param data - The object data
   */
  private handleObjectMessage(data: Record<string, unknown>): void {
    this.logger.info('Object message received', data, 'WS-IN');

    // Check for event field
    if (data.event) {
      this.eventHandlers.eventReceived?.(data.event as string, data);
    }

    // Check for command field
    if (data.command) {
      this.eventHandlers.commandReceived?.(data.command as string, data.params as Record<string, unknown>);
    }
  }

  /**
   * Register an event handler
   * @param event - The event name
   * @param handler - The handler function
   */
  on<K extends keyof MessageHandlerEvents>(
    event: K,
    handler: MessageHandlerEvents[K]
  ): void {
    this.eventHandlers[event] = handler;
  }

  /**
   * Format a message for display
   * @param parsed - The parsed message
   * @returns Formatted string for display
   */
  formatForDisplay(parsed: ParsedMessage): string {
    return KanataProtocol.formatForDisplay(parsed);
  }
}
