/**
 * BackgroundService - Headless Background Service for Excel Shared Runtime
 * Manages WebSocket connection, handles dynamic Ribbon updates, and executes remote Excel commands.
 */

import { ConnectionManager, ConnectionState } from '../client/ConnectionManager';
import { Logger } from '../logging/Logger';
import { CommandResult, ParsedMessage, WebSocketResponse } from '../protocol/types';
import { ExcelCommandExecutor } from './ExcelCommandExecutor';
import { ExcelSpotlightService } from './ExcelSpotlightService';

export class BackgroundService {
  private static instance: BackgroundService | null = null;
  private connectionManager: ConnectionManager;
  private commandExecutor: ExcelCommandExecutor;
  private logger: Logger;
  private isInitialized = false;

  private constructor() {
    this.logger = Logger.getInstance();
    this.connectionManager = new ConnectionManager(this.logger);
    this.commandExecutor = new ExcelCommandExecutor(this.logger, this.connectionManager);
  }

  /**
   * Get singleton instance
   */
  public static getInstance(): BackgroundService {
    if (!BackgroundService.instance) {
      BackgroundService.instance = new BackgroundService();
    }
    return BackgroundService.instance;
  }

  /**
   * Get the underlying ConnectionManager
   */
  public getConnectionManager(): ConnectionManager {
    return this.connectionManager;
  }

  /**
   * Get the ExcelCommandExecutor
   */
  public getCommandExecutor(): ExcelCommandExecutor {
    return this.commandExecutor;
  }

  /**
   * Get the Logger
   */
  public getLogger(): Logger {
    return this.logger;
  }

  /**
   * Initialize the headless background service
   */
  public async initialize(): Promise<void> {
    if (this.isInitialized) {
      return;
    }
    this.isInitialized = true;

    this.logger.info('Initializing Headless BackgroundService in Shared Runtime');

    // 1. Configure Excel to auto-load this add-in on document open
    this.configureStartupBehavior();

    // 2. Setup event handlers for WebSocket lifecycle
    this.connectionManager.on('stateChange', (state: ConnectionState) => {
      this.logger.info(`WebSocket state changed: ${state}`);
      this.updateRibbonState(state);
      if (state === ConnectionState.Connected) {
        this.pushStatusUpdate().catch(() => {});
      }
    });

    this.connectionManager.on('messageReceived', (message: ParsedMessage) => {
      this.handleRemoteMessage(message);
    });

    // 3. Sync initial ribbon status
    await this.updateRibbonState(this.connectionManager.getConnectionState());

    // 4. Auto-connect if configured or by default
    try {
      const settings = await this.connectionManager.getSettings();
      this.logger.info('BackgroundService settings loaded:', settings);
      
      if (settings.automationEnabled !== false && settings.autoConnect !== false) {
        await this.connectionManager.connect();
      } else {
        this.logger.info('BackgroundService automation is disabled; staying offline');
        await this.updateRibbonState(ConnectionState.Disconnected);
      }
    } catch (error) {
      this.logger.error('Failed to auto-connect in BackgroundService:', error);
    }

    // 5. Initialize position slot manager and register passive selection history tracking
    try {
      this.commandExecutor.getSlotManager().loadPersistedSettings();
      await this.commandExecutor.getSlotManager().registerSelectionTracker();
    } catch (err) {
      this.logger.warn('Selection tracking setup warning:', err);
    }

    // 6. Startup Auto-Detox: Clean up any orphan spotlight conditional formatting left by prior crashes
    try {
      const spotlight = ExcelSpotlightService.getInstance(this.logger);
      await spotlight.purgeAllOrphanSpotlightsAcrossWorkbook();
      if (spotlight.getSettings().enabled) {
        await spotlight.refreshSpotlight();
      }
    } catch (err) {
      this.logger.warn('Spotlight startup auto-detox warning:', err);
    }
  }

  /**
   * Ensure add-in starts automatically in background on document open
   */
  private configureStartupBehavior(): void {
    try {
      if (typeof Office !== 'undefined' && Office.addin && Office.addin.setStartupBehavior) {
        Office.addin.setStartupBehavior(Office.StartupBehavior.load)
          .then(() => {
            this.logger.info('Set Office.addin.setStartupBehavior to load on document open');
          })
          .catch((error) => {
            this.logger.warn('Could not set startup behavior (might not be supported on this platform):', error);
          });
      }
    } catch (error) {
      this.logger.warn('Could not set startup behavior (might not be supported on this platform):', error);
    }
  }

  /**
   * Update the Excel ribbon button icons and visibility dynamically
   * @param state Current WebSocket connection state
   */
  public async updateRibbonState(state: ConnectionState): Promise<void> {
    if (typeof Office === 'undefined' || !Office.ribbon || !Office.ribbon.requestUpdate) {
      this.logger.debug('Office.ribbon.requestUpdate not available in this host environment');
      return;
    }

    const isConnected = state === ConnectionState.Connected;

    try {
      await Office.ribbon.requestUpdate({
        tabs: [
          {
            id: 'TabHome',
            groups: [
              {
                id: 'ExcelWebSocketGroup',
                controls: [
                  {
                    id: 'btnExcelWebSocket',
                    enabled: isConnected
                  }
                ]
              }
            ]
          }
        ]
      });
      this.logger.info(`Ribbon updated: btnExcelWebSocket.enabled=${isConnected}`);
    } catch (error) {
      this.logger.warn('Office.ribbon.requestUpdate skipped:', error);
    }
  }

  /**
   * Process incoming WebSocket message from Kanata or external agent
   * @param message Parsed incoming message
   */
  private async handleRemoteMessage(message: ParsedMessage): Promise<void> {
    if (!this.connectionManager.isAutomationEnabled()) {
      this.logger.warn('Ignoring remote message: Kanata automation is disabled by user');
      return;
    }

    this.logger.info('BackgroundService received message:', message);

    let cmdPayload: any = null;

    if (message.type === 'object' && typeof message.data === 'object' && message.data !== null) {
      cmdPayload = message.data;
    } else if (typeof message.data === 'string') {
      try {
        cmdPayload = JSON.parse(message.data);
      } catch {
        cmdPayload = { command: message.data };
      }
    }

    if (!cmdPayload) {
      return;
    }

    const commandName = cmdPayload.command || cmdPayload.action;
    const commandId = cmdPayload.id || `excel-${Date.now()}`;
    const params = (cmdPayload.params as Record<string, unknown>) || cmdPayload;

    if (!commandName) {
      this.logger.debug('Ignoring non-command payload:', cmdPayload);
      return;
    }

    // Execute via ExcelCommandExecutor
    try {
      const result = await this.commandExecutor.execute(commandName, params);

      // Send standard Kanata protocol response
      const response: WebSocketResponse = {
        type: 'response',
        id: commandId,
        status: result.ok ? 'success' : 'error',
        result: result.data,
        error: result.ok ? undefined : result.error
      };

      this.connectionManager.send(response);
      this.logger.info(`Sent response for [${commandId}]:`, response);
    } catch (err: any) {
      const errorMsg = err instanceof Error ? err.message : String(err);
      this.logger.error(`Unhandled error executing remote command [${commandName}]:`, err);

      const errorResponse: WebSocketResponse = {
        type: 'response',
        id: commandId,
        status: 'error',
        error: errorMsg
      };
      this.connectionManager.send(errorResponse);
    }
  }

  /**
   * Push current workbook status, active cell, and focus details to the WebSocket server
   */
  public async pushStatusUpdate(): Promise<void> {
    try {
      const status = await this.commandExecutor.getSlotManager().getWorkbookStatus();
      this.connectionManager.send({
        type: 'status' as any,
        event: 'workbook_status',
        data: status
      } as any);
      this.logger.debug('Pushed workbook status update:', status);
    } catch (err) {
      this.logger.debug('Failed to push status update:', err);
    }
  }

  /**
   * Execute Excel operations headlessly in background (alias method for backwards compatibility)
   * @param command Command payload
   */
  public async executeExcelCommand(command: any): Promise<void> {
    const action = command.action || command.command;
    const params = command.params || command;
    if (action) {
      await this.commandExecutor.execute(action, params);
    }
  }
}
