import { Component, createSignal, onMount, For } from 'solid-js';
import { BackgroundService } from '../../../websocket/service/BackgroundService';
import { ConnectionState } from '../../../websocket/client/ConnectionManager';
import { LogEntry } from '../../../websocket/logging/Logger';
import { useI18n } from '../../../i18n';

export const SettingsTab: Component = () => {
  const { t, locale, setLocale } = useI18n();
  const bgService = BackgroundService.getInstance();
  const connMgr = bgService.getConnectionManager();
  const logger = bgService.getLogger();

  const [connState, setConnState] = createSignal<ConnectionState>(ConnectionState.Disconnected);
  const [port, setPort] = createSignal(8890);
  const [clientName, setClientName] = createSignal('excel');
  const [channels, setChannels] = createSignal('excel_websocket_Offic_Add_in');
  const [logs, setLogs] = createSignal<LogEntry[]>([]);
  const [toastMsg, setToastMsg] = createSignal<string>('');

  const refreshState = async () => {
    setConnState(connMgr.getConnectionState());
    const s = await connMgr.getSettings();
    setPort(s.port);
    setClientName(s.clientName);
    setChannels(s.channels.join(','));
    setLogs(logger.getLogs().slice(-20));
  };

  onMount(async () => {
    try {
      await refreshState();
      connMgr.on('stateChange', (state) => {
        setConnState(state);
      });
      setInterval(() => {
        setLogs(logger.getLogs().slice(-20));
      }, 1500);
    } catch (err) {
      console.warn('SettingsTab onMount error:', err);
    }
  });

  const getFormattedLogs = (): string => {
    const allLogs = logger.getLogs();
    const diag = typeof Office !== 'undefined' && Office.context ? Office.context.diagnostics : null;
    const hostInfo = diag ? `${diag.host} (${diag.platform}) v${diag.version}` : 'Unknown Host';

    const header = [
      '================================================================================',
      '                     Excel Flow Add-in Diagnostic Log Bundle',
      ` Exported At: ${new Date().toLocaleString()} (${new Date().toISOString()})`,
      ` Environment: ${hostInfo}`,
      ` Language: ${locale()}`,
      ` Connection State: ${connState()}`,
      ` WebSocket Server: ws://localhost:${port()}`,
      ` Channels: ${channels()}`,
      ` Total Log Count: ${allLogs.length} (Ring Buffer Max: 1000)`,
      '================================================================================\n'
    ].join('\n');

    const lines = allLogs.map((entry) => {
      const time = entry.timestamp ? new Date(entry.timestamp).toISOString() : new Date().toISOString();
      const level = (entry.level || 'INFO').toUpperCase().padEnd(5, ' ');
      const scope = (entry.scope || 'SYSTEM').padEnd(10, ' ');
      const dataStr = entry.data !== undefined ? ` | Data: ${JSON.stringify(entry.data)}` : '';
      const errStr = entry.error ? ` | Error: ${JSON.stringify(entry.error)}` : '';
      return `[${time}] [${level}] [${scope}] ${entry.message}${dataStr}${errStr}`;
    }).join('\n');

    return `${header}\n${lines}`;
  };

  const handleExportFile = () => {
    try {
      const content = getFormattedLogs();
      const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      const now = new Date();
      const timeStr = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}-${String(now.getHours()).padStart(2, '0')}${String(now.getMinutes()).padStart(2, '0')}${String(now.getSeconds()).padStart(2, '0')}`;
      a.href = url;
      a.download = `excel-flow-debug-${timeStr}.log`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      setToastMsg(t().common.downloadSuccess);
      setTimeout(() => setToastMsg(''), 3000);
    } catch {
      handleCopyLogs();
    }
  };

  const handleCopyLogs = async () => {
    try {
      const content = getFormattedLogs();
      if (navigator.clipboard && navigator.clipboard.writeText) {
        await navigator.clipboard.writeText(content);
      } else {
        const textarea = document.createElement('textarea');
        textarea.value = content;
        document.body.appendChild(textarea);
        textarea.select();
        document.execCommand('copy');
        document.body.removeChild(textarea);
      }
      setToastMsg(t().common.copiedSuccess);
      setTimeout(() => setToastMsg(''), 3000);
    } catch {
      setToastMsg('✕ Failed to copy logs');
      setTimeout(() => setToastMsg(''), 3000);
    }
  };

  const handleConnect = async () => {
    await connMgr.saveSettings({
      port: Number(port()),
      clientName: clientName(),
      channels: channels().split(',').map((c) => c.trim()).filter(Boolean)
    });
    await connMgr.connect();
    await refreshState();
  };

  const handleDisconnect = async () => {
    await connMgr.disconnect();
    await refreshState();
  };

  const isConnected = () => connState() === ConnectionState.Connected;

  return (
    <div class="tab-panel settings-tab">
      {/* 1. Language Selector Card */}
      <div class="field-card" style={{ "margin-bottom": "14px" }}>
        <div class="control-group" style={{ "margin-bottom": "0" }}>
          <label class="control-label" style={{ "font-weight": "600" }}>{t().settings.languageLabel}</label>
          <div class="chip-group">
            <button
              type="button"
              class="chip-btn"
              classList={{ active: locale() === 'zh-CN' }}
              onClick={() => setLocale('zh-CN')}
            >
              🇨🇳 中文 (简体)
            </button>
            <button
              type="button"
              class="chip-btn"
              classList={{ active: locale() === 'en-US' }}
              onClick={() => setLocale('en-US')}
            >
              🇺🇸 English
            </button>
          </div>
        </div>
      </div>

      <div class="card-intro">
        <h3>{t().settings.bridgeTitle}</h3>
        <p>{t().settings.bridgeDesc}</p>
      </div>

      <div class="field-card">
        <div class="field-row">
          <label>{t().settings.port}</label>
          <input
            type="number"
            class="text-input mini"
            value={port()}
            onInput={(e) => setPort(Number(e.currentTarget.value))}
          />
        </div>

        <div class="field-row">
          <label>{t().settings.clientName}</label>
          <input
            type="text"
            class="text-input"
            value={clientName()}
            onInput={(e) => setClientName(e.currentTarget.value)}
          />
        </div>

        <div class="field-row">
          <label>{t().settings.channels}</label>
          <input
            type="text"
            class="text-input"
            value={channels()}
            onInput={(e) => setChannels(e.currentTarget.value)}
          />
        </div>
      </div>

      <div class="action-stack">
        <button
          type="button"
          class="action-btn"
          classList={{ primary: !isConnected(), danger: isConnected() }}
          onClick={isConnected() ? handleDisconnect : handleConnect}
        >
          {isConnected() ? `🔌 ${t().settings.disconnect}` : `⚡ ${t().settings.connect}`}
        </button>
      </div>

      {/* Log Monitor with Export and Copy */}
      <div class="log-section">
        <div class="log-header">
          <div class="log-header-left">
            <span>{t().settings.consoleLogs}:</span>
            {toastMsg() && <span class="log-toast">{toastMsg()}</span>}
          </div>
          <div class="log-actions">
            <button
              type="button"
              class="log-action-btn export"
              onClick={handleExportFile}
              title={t().settings.exportLogs}
            >
              {t().settings.exportLogs}
            </button>
            <button
              type="button"
              class="log-action-btn copy"
              onClick={handleCopyLogs}
              title={t().settings.copyLogs}
            >
              {t().settings.copyLogs}
            </button>
            <button
              type="button"
              class="log-action-btn clear"
              onClick={() => {
                logger.clearLogs();
                setLogs([]);
              }}
              title="Clear"
            >
              🗑️
            </button>
          </div>
        </div>
        <div class="log-box">
          <For each={logs()}>
            {(entry) => (
              <div class="log-line" classList={{ err: entry.level === 'error', warn: entry.level === 'warn' }}>
                <span class="log-time">{entry.timestamp.toLocaleTimeString()}</span>
                <span class="log-level">[{entry.level.toUpperCase()}]</span>
                <span class="log-msg">{entry.message}</span>
              </div>
            )}
          </For>
        </div>
      </div>
    </div>
  );
};
