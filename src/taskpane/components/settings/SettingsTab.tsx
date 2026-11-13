import { Component, createSignal, onMount } from 'solid-js';
import { BackgroundService } from '../../../websocket/service/BackgroundService';
import { ConnectionState } from '../../../websocket/client/ConnectionManager';
import { useI18n } from '../../../i18n';
import { LogMonitorCard } from './LogMonitorCard';

export const SettingsTab: Component = () => {
  const { t, locale, setLocale } = useI18n();
  const bgService = BackgroundService.getInstance();
  const connMgr = bgService.getConnectionManager();
  const logger = bgService.getLogger();

  const [connState, setConnState] = createSignal<ConnectionState>(ConnectionState.Disconnected);
  const [port, setPort] = createSignal(8890);
  const [clientName, setClientName] = createSignal('excel');
  const [channels, setChannels] = createSignal('excel_websocket_Offic_Add_in');
  const [automationEnabled, setAutomationEnabled] = createSignal(true);

  const refreshState = async () => {
    setConnState(connMgr.getConnectionState());
    const s = await connMgr.getSettings();
    setPort(s.port);
    setClientName(s.clientName);
    setChannels(s.channels.join(','));
    setAutomationEnabled(s.automationEnabled !== false);
  };

  onMount(async () => {
    try {
      await refreshState();
      connMgr.on('stateChange', (state) => {
        setConnState(state);
      });
    } catch (err) {
      console.warn('SettingsTab onMount error:', err);
    }
  });

  const handleToggleAutomation = async () => {
    const next = !automationEnabled();
    setAutomationEnabled(next);
    await connMgr.setAutomationEnabled(next);
    await refreshState();
  };

  const handleConnect = async () => {
    await connMgr.saveSettings({
      port: Number(port()),
      clientName: clientName(),
      channels: channels().split(',').map((c) => c.trim()).filter(Boolean),
      autoConnect: true,
      clientType: 'office-addin',
      automationEnabled: automationEnabled()
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

      {/* 2. Automation Master Toggle Card */}
      <div class="field-card" style={{ "margin-bottom": "14px" }}>
        <div class="field-row" style={{ "justify-content": "space-between", "align-items": "center" }}>
          <div style={{ "display": "flex", "flex-direction": "column", "gap": "2px", "padding-right": "8px" }}>
            <span style={{ "font-weight": "600", "font-size": "13px" }}>{t().settings.automationToggle}</span>
            <span style={{ "font-size": "11px", "color": "var(--text-muted)" }}>
              {t().settings.automationDesc}
            </span>
          </div>
          <button
            type="button"
            class="toggle-switch-btn"
            classList={{ active: automationEnabled() }}
            onClick={handleToggleAutomation}
            title={automationEnabled() ? t().common.enabled : t().common.disabled}
          >
            <span class="toggle-knob" />
            <span class="toggle-label">{automationEnabled() ? t().common.enabled : t().common.disabled}</span>
          </button>
        </div>
      </div>

      {/* 3. Kanata WebSocket Bridge Settings */}
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

      {/* 4. Log Monitor */}
      <LogMonitorCard
        logger={logger}
        port={port}
        channels={channels}
        connState={connState}
      />
    </div>
  );
};
