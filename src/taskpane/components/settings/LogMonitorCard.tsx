import { Component, createSignal, For, onMount, onCleanup } from 'solid-js';
import { Logger, LogEntry } from '../../../websocket/logging/Logger';
import { ConnectionState } from '../../../websocket/client/ConnectionManager';
import { useI18n } from '../../../i18n';

interface LogMonitorCardProps {
  logger: Logger;
  port: () => number;
  channels: () => string;
  connState: () => ConnectionState;
}

export const LogMonitorCard: Component<LogMonitorCardProps> = (props) => {
  const { t, locale } = useI18n();
  const [logs, setLogs] = createSignal<LogEntry[]>([]);
  const [toastMsg, setToastMsg] = createSignal<string>('');

  onMount(() => {
    setLogs(props.logger.getLogs().slice(-20));
    const timer = setInterval(() => {
      setLogs(props.logger.getLogs().slice(-20));
    }, 1500);
    onCleanup(() => clearInterval(timer));
  });

  const getFormattedLogs = (): string => {
    const allLogs = props.logger.getLogs();
    const diag = typeof Office !== 'undefined' && Office.context ? Office.context.diagnostics : null;
    const hostInfo = diag ? `${diag.host} (${diag.platform}) v${diag.version}` : 'Unknown Host';

    const header = [
      '================================================================================',
      '                     Excel Flow Add-in Diagnostic Log Bundle',
      ` Exported At: ${new Date().toLocaleString()} (${new Date().toISOString()})`,
      ` Environment: ${hostInfo}`,
      ` Language: ${locale()}`,
      ` Connection State: ${props.connState()}`,
      ` WebSocket Server: ws://localhost:${props.port()}`,
      ` Channels: ${props.channels()}`,
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

  return (
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
              props.logger.clearLogs();
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
  );
};
