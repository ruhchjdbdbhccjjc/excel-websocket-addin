import "./excel";
//import "./outlook";
//import "./powerpoint";
//import "./word";

// Add WebSocket initialization here since this is the actual loaded file
import { ConnectionManager } from "../websocket/client/ConnectionManager";
import { ConnectionState } from "../websocket/client/ConnectionManager";
import { ParsedMessage } from "../websocket/protocol/types";
import { Logger, LogLevel } from "../websocket/logging/Logger";
import { BackgroundService } from "../websocket/service/BackgroundService";

let connectionManager: ConnectionManager | null = null;
let logger: Logger | null = new Logger();

Office.onReady(async (info) => {
  logger?.info("Office.onReady called, host:", info.host);
  
  if (info.host === Office.HostType.Excel) {

    // 1. Initialize Headless Background Service (boots WebSocket & syncs ribbon)
    const bgService = BackgroundService.getInstance();
    await bgService.initialize();

    const sideloadMsg = document.getElementById("sideload-msg");
    if (sideloadMsg) {
      sideloadMsg.style.display = "none";
    }
    const appBody = document.getElementById("app-body");
    if (appBody) {
      appBody.style.display = "flex";
    }
    const run = document.getElementById("run");
    if (run) {
      run.onclick = runExcel;
    }
    
    // 2. Bind TaskPane UI to shared background connection manager
    await initializeWebSocket(bgService);
  }
});

async function initializeWebSocket(bgService: BackgroundService) {
  logger?.info("initializeWebSocket called");
  
  logger = bgService.getLogger();
  connectionManager = bgService.getConnectionManager();
  
  logger?.info('Connection manager retrieved from BackgroundService');
  
  // Setup event listeners now that connectionManager is initialized
  setupEventListeners();
  
  // Load saved settings
  const settings = await connectionManager.getSettings();
  
  logger?.info('Settings loaded:', settings);
  
  // Update UI with saved settings
  const portInput = document.getElementById('ws-port') as HTMLInputElement;
  const clientNameInput = document.getElementById('ws-client-name') as HTMLInputElement;
  const clientTypeInput = document.getElementById('ws-client-type') as HTMLInputElement;
  const channelsInput = document.getElementById('ws-channels') as HTMLInputElement;
  const autoConnectCheckbox = document.getElementById('ws-auto-connect') as HTMLInputElement;
  
  logger?.debug('UI elements retrieved: port=', portInput ? 'found' : 'not found');
  
  if (portInput) portInput.value = settings.port.toString();
  if (clientNameInput) clientNameInput.value = settings.clientName;
  if (clientTypeInput) clientTypeInput.value = settings.clientType;
  if (channelsInput) channelsInput.value = settings.channels.join(',');
  if (autoConnectCheckbox) autoConnectCheckbox.checked = settings.autoConnect;
  
  logger?.debug('UI updated');
  
  // If not already connected or connecting, connect if autoConnect is enabled
  if (connectionManager.getConnectionState() === ConnectionState.Disconnected && settings.autoConnect) {
    try {
      await connectionManager.connect();
    } catch (error) {
      logger?.error('Auto-connect failed:', error);
    }
  }
  
  // Update UI with current state
  await updateConnectionUI();
  
  logger?.info('initializeWebSocket complete');
}

function setupEventListeners() {
  logger?.debug('setupEventListeners called');
  
  // Connect button
  const connectBtn = document.getElementById("ws-connect");
  logger?.debug("Connect button element:", connectBtn ? "found" : "not found");
  
  if (connectBtn) {
    connectBtn.addEventListener("click", async () => {
      logger?.info('TEST: Connect button clicked - logging works!');
      
      // Wait for connectionManager to be initialized if needed
      if (!connectionManager) {
        logger?.error('Connection manager not initialized yet');
        return;
      }
      
      const portInput = document.getElementById('ws-port') as HTMLInputElement;
      const clientNameInput = document.getElementById('ws-client-name') as HTMLInputElement;
      const clientTypeInput = document.getElementById('ws-client-type') as HTMLInputElement;
      const channelsInput = document.getElementById('ws-channels') as HTMLInputElement;
      const autoConnectCheckbox = document.getElementById('ws-auto-connect') as HTMLInputElement;
      
      const port = parseInt(portInput.value);
      const clientName = clientNameInput.value;
      const clientType = clientTypeInput.value;
      const channels = channelsInput.value.split(',').map(c => c.trim()).filter(c => c.length > 0);
      
      logger?.info('Port:', port, 'Client:', clientName, 'Type:', clientType, 'Channels:', channels);
      
      // Validate port
      if (isNaN(port) || port < 1 || port > 65535) {
        logger?.error('Invalid port number');
        return;
      }
      
      // Validate client name
      if (!clientName.trim()) {
        logger?.error('Client name is empty');
        return;
      }
      
      logger?.info('Saving settings...');
      
      await connectionManager.saveSettings({
        port,
        clientName,
        clientType,
        channels: channels.length > 0 ? channels : ['excel-addin-websocket'],
        autoConnect: autoConnectCheckbox.checked
      });
      
      logger?.info('Settings saved. Attempting to connect...');
      
      try {
        await connectionManager.connect();
        logger?.info('Connection successful!');
      } catch (error) {
        logger?.error('Connection failed:', error);
      }
      await updateConnectionUI();
    });
  } else {
    logger?.error('Connect button not found in DOM!');
  }
  
  // Disconnect button
  const disconnectBtn = document.getElementById('ws-disconnect');
  if (disconnectBtn) {
    disconnectBtn.addEventListener('click', async () => {
      if (connectionManager) {
        connectionManager.disconnect();
        await updateConnectionUI();
      }
    });
  }
  
  // Clear logs button
  const clearLogsBtn = document.getElementById('clear-logs');
  if (clearLogsBtn) {
    clearLogsBtn.addEventListener('click', async () => {
      if (connectionManager) {
        connectionManager.clearLogs();
        await updateLogViewer();
      }
    });
  }
  
  // State change handler
  if (connectionManager) {
    connectionManager.on('stateChange', async (state: ConnectionState) => {
      await updateConnectionUI();
    });
    
    // Message received handler
    connectionManager.on('messageReceived', async (parsed: ParsedMessage) => {
      await updateLogViewer();
    });
    
    // Log entry handler
    connectionManager.on('logEntry', async () => {
      await updateLogViewer();
    });
  }
}

async function updateConnectionUI() {
  if (!connectionManager) return;
  
  const state = connectionManager.getConnectionState();
  const statusIndicator = document.getElementById('status-indicator');
  const statusText = document.getElementById('status-text');
  const connectionInfo = document.getElementById('connection-info');
  const connectBtn = document.getElementById('ws-connect');
  const disconnectBtn = document.getElementById('ws-disconnect');
  const logViewer = document.getElementById('log-viewer');
  
  // Update status indicator
  if (statusIndicator) {
    statusIndicator.className = 'status-dot ' + state;
  }
  
  // Update status text
  const statusMessages: Record<ConnectionState, string> = {
    disconnected: 'Disconnected',
    connecting: 'Connecting...',
    connected: 'Connected',
    error: 'Connection Error'
  };
  if (statusText) {
    statusText.textContent = statusMessages[state];
  }
  
  // Show/hide connection info
  if (state === ConnectionState.Connected) {
    const settings = await connectionManager.getSettings();
    if (connectionInfo) {
      connectionInfo.style.display = 'block';
      const channelDisplay = document.getElementById('channel-display');
      if (channelDisplay) {
        const channelsStr = settings.channels && Array.isArray(settings.channels) ? settings.channels.join(', ') : '';
        channelDisplay.textContent = `${settings.clientName}:${settings.port} (${settings.clientType}) [${channelsStr}]`;
      }
    }
    if (logViewer) {
      logViewer.style.display = 'block';
    }
  } else {
    if (connectionInfo) {
      connectionInfo.style.display = 'none';
    }
    if (logViewer) {
      logViewer.style.display = 'none';
    }
  }
  
  // Enable/disable buttons
  if (connectBtn) {
    connectBtn.disabled = state === ConnectionState.Connected || state === ConnectionState.Connecting;
  }
  if (disconnectBtn) {
    disconnectBtn.disabled = state !== ConnectionState.Connected;
  }
}

async function updateLogViewer() {
  if (!connectionManager) return;
  
  const logContainer = document.getElementById('log-container');
  if (!logContainer) return;
  
  const logs = connectionManager.getLogs();
  const messageLogs = logs.filter(log => 
    log.message.includes('Message received') || 
    log.message.includes('String message received') ||
    log.message.includes('Array message received') ||
    log.message.includes('Object message received')
  );
  
  // Show last 20 message logs
  const recentLogs = messageLogs.slice(-20);
  
  logContainer.innerHTML = recentLogs.map(log => {
    const timestamp = log.timestamp.toLocaleTimeString();
    const dataStr = log.data ? JSON.stringify(log.data) : '';
    return `<div class="log-entry ${log.level}">[${timestamp}] ${log.message} ${dataStr}</div>`;
  }).join('');
  
  // Scroll to bottom
  logContainer.scrollTop = logContainer.scrollHeight;
}

export async function runExcel() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "yellow";

      await context.sync();
      logger?.info(`The range address was ${range.address}.`);
    });
  } catch (error) {
    logger?.error(error);
  }
}
