/**
 * KanataProtocol - Handle Kanata-specific message formats
 * Supports simple strings, JSON arrays, and JSON objects
 */

import { ParsedMessage, WebSocketRegister } from './types';
import { ConnectionSettings } from '../storage/ConnectionSettings';
import { generateUUID } from '../../utils/uuid';

export class KanataProtocol {
  /**
   * Parse a message from Kanata
   * Kanata sends simple strings, JSON arrays, or JSON objects
   * @param data - The raw message data
   * @returns Parsed message with type information
   */
  static parseMessage(data: string): ParsedMessage {
    try {
      // Try to parse as JSON
      const parsed = JSON.parse(data);

      if (Array.isArray(parsed)) {
        return {
          type: 'array',
          data: parsed,
          raw: data
        };
      } else if (typeof parsed === 'object' && parsed !== null) {
        return {
          type: 'object',
          data: parsed,
          raw: data
        };
      }
    } catch (e) {
      // Not JSON, treat as string
    }

    return {
      type: 'string',
      data: data,
      raw: data
    };
  }

  /**
   * Format a register message to send to the server
   * This is sent immediately after connection to identify the client
   * @param settings - Connection settings
   * @returns JSON string of register message
   */
  static formatRegisterMessage(settings: ConnectionSettings): string {
    const register: WebSocketRegister = {
      type: 'register',
      instanceId: generateUUID(),
      version: '1.0.0',
      clientType: settings.clientType,
      clientName: settings.clientName,
      timestamp: Date.now(),
      capabilities: ['receive-messages'],
      channels: settings.channels
    };
    return JSON.stringify(register);
  }

  /**
   * Format a heartbeat message
   * @returns JSON string of heartbeat message
   */
  static formatHeartbeat(): string {
    return JSON.stringify({
      type: 'heartbeat',
      timestamp: Date.now()
    });
  }

  /**
   * Check if a message is a valid Kanata message
   * @param data - The message data
   * @returns Whether the message is valid
   */
  static isValidMessage(data: string): boolean {
    // All strings are considered valid (Kanata sends simple strings)
    if (typeof data === 'string' && data.length > 0) {
      return true;
    }
    return false;
  }

  /**
   * Format a message for display in the UI
   * @param parsed - The parsed message
   * @returns Formatted string for display
   */
  static formatForDisplay(parsed: ParsedMessage): string {
    switch (parsed.type) {
      case 'string':
        return parsed.data as string;
      case 'array':
        return JSON.stringify(parsed.data, null, 2);
      case 'object':
        return JSON.stringify(parsed.data, null, 2);
      default:
        return parsed.raw;
    }
  }
}
