/**
 * ConnectionSettings - Interface for WebSocket connection settings
 */

export interface ConnectionSettings {
  port: number;
  clientName: string;
  clientType: string;
  channels: string[];
  autoConnect: boolean;
  automationEnabled?: boolean;
  onlyWhenFocused?: boolean;
  logLevel?: string; // Optional for backward compatibility
}

export const DEFAULT_SETTINGS: ConnectionSettings = {
  port: 8890,
  clientName: 'excel',
  clientType: 'office-addin',
  channels: ['excel_websocket_Offic_Add_in'],
  autoConnect: true,
  automationEnabled: true,
  onlyWhenFocused: true,
  logLevel: 'info'
};
