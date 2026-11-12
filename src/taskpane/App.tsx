import { Component, createSignal, onMount } from 'solid-js';
import { BackgroundService } from '../websocket/service/BackgroundService';
import { ConnectionState } from '../websocket/client/ConnectionManager';
import { useI18n } from '../i18n';
import { SpotlightTab } from './components/spotlight/SpotlightTab';
import { PositionSlotsTab } from './components/slots/PositionSlotsTab';
import { NavigationTab } from './components/navigation/NavigationTab';
import { SettingsTab } from './components/settings/SettingsTab';
import { CommandPalette } from './components/CommandPalette';

import { useTheme, ThemeMode } from './theme/ThemeManager';

export type TabId = 'spotlight' | 'slots' | 'nav' | 'palette' | 'settings';

export const App: Component = () => {
  const { locale, setLocale, t } = useI18n();
  const { themeMode, setThemeMode, resolvedTheme } = useTheme();
  const isDialog = () => typeof window !== 'undefined' && Boolean(window.location?.search?.includes('dialog=true'));

  const closeDialogWindow = () => {
    try {
      if (typeof Office !== 'undefined' && Office.context && Office.context.ui && Office.context.ui.messageParent) {
        Office.context.ui.messageParent(JSON.stringify({ action: 'close' }));
      }
    } catch {
      // ignore
    }
  };

  const [activeTab, setActiveTab] = createSignal<TabId>(isDialog() ? 'palette' : 'spotlight');
  const [connState, setConnState] = createSignal<ConnectionState>(ConnectionState.Disconnected);

  onMount(() => {
    try {
      if (!isDialog()) {
        const bgService = BackgroundService.getInstance();
        const connMgr = bgService.getConnectionManager();

        setConnState(connMgr.getConnectionState());
        connMgr.on('stateChange', (state) => {
          setConnState(state);
        });

        // 1. WebSocket trigger from Kanata (bypasses Excel native conflict dialogs)
        connMgr.on('messageReceived', (parsed: any) => {
          const raw = parsed?.raw || parsed;
          const cmd = parsed?.command || raw?.command || (Array.isArray(raw) ? raw[0] : null) || raw?.type;
          if (cmd === 'open_command_palette') {
            setActiveTab('palette');
            const isExplicitTaskpane = raw?.mode === 'taskpane' || raw?.params?.mode === 'taskpane' || parsed?.mode === 'taskpane';
            if (isExplicitTaskpane && typeof Office !== 'undefined' && Office.addin && Office.addin.showAsTaskpane) {
              Office.addin.showAsTaskpane().catch(() => {});
            }
          }
        });
      }

      // 2. Global DOM keyboard listener within add-in window
      // Supports both Ctrl+Shift+P (primary standard) and Ctrl+Shift+K (alias)
      const handleGlobalKeyDown = (e: KeyboardEvent) => {
        if (e.key === 'Escape' && isDialog()) {
          const target = e.target as HTMLElement | null;
          if (target instanceof HTMLInputElement && target.value.trim().length > 0) {
            return;
          }
          e.preventDefault();
          closeDialogWindow();
          return;
        }
        if ((e.ctrlKey || e.metaKey) && e.shiftKey) {
          const key = e.key.toLowerCase();
          if (key === 'p' || key === 'k') {
            e.preventDefault();
            e.stopPropagation();
            setActiveTab('palette');
          }
        }
      };
      window.addEventListener('keydown', handleGlobalKeyDown, false);
    } catch (err) {
      console.warn('App onMount connection manager warning:', err);
    }
  });

  const isConnected = () => connState() === ConnectionState.Connected;

  const cycleTheme = () => {
    const current = themeMode();
    if (current === 'auto') setThemeMode('light');
    else if (current === 'light') setThemeMode('dark');
    else setThemeMode('auto');
  };

  const toggleLanguage = () => {
    setLocale(locale() === 'zh-CN' ? 'en-US' : 'zh-CN');
  };

  return (
    <div class="app-shell">
      <header class="app-header">
        <div class="header-brand">
          <span class="brand-icon">⚡</span>
          <span class="brand-title">Excel Flow</span>
        </div>
        <div style={{ display: 'flex', 'align-items': 'center', gap: '6px' }}>
          {/* Language Toggle Pill */}
          <button
            type="button"
            onClick={toggleLanguage}
            title="Switch Language / 切换语言"
            style={{
              padding: '2px 7px',
              'font-size': '11px',
              'font-weight': '600',
              border: '1px solid var(--border-subtle)',
              'border-radius': '12px',
              background: 'rgba(128,128,128,0.12)',
              color: 'var(--text-primary)',
              cursor: 'pointer'
            }}
          >
            {locale() === 'zh-CN' ? '🇨🇳 中文' : '🇺🇸 EN'}
          </button>

          {/* Theme Mode Toggle Pill */}
          <button
            type="button"
            onClick={cycleTheme}
            title={`Theme Mode: ${themeMode()} (Active: ${resolvedTheme()})`}
            style={{
              padding: '2px 7px',
              'font-size': '11px',
              'font-weight': '600',
              border: '1px solid var(--border-subtle)',
              'border-radius': '12px',
              background: 'rgba(128,128,128,0.12)',
              color: 'var(--text-primary)',
              cursor: 'pointer'
            }}
          >
            {themeMode() === 'auto' ? '🌓 自动' : themeMode() === 'light' ? '☀️ 浅色' : '🌙 深色'}
          </button>

          {/* Connection Status Pill */}
          <div class="status-pill" classList={{ online: isConnected(), offline: !isConnected() }}>
            <span class="status-dot" />
            <span class="status-text">{isConnected() ? t().brand.statusConnected : t().brand.statusDisconnected}</span>
          </div>

          {/* Dialog Close Button */}
          {isDialog() && (
            <button
              type="button"
              onClick={closeDialogWindow}
              title="Close Dialog / 关闭窗口 (Esc)"
              style={{
                padding: '2px 8px',
                'font-size': '11px',
                'font-weight': '600',
                border: '1px solid rgba(239, 68, 68, 0.4)',
                'border-radius': '12px',
                background: 'rgba(239, 68, 68, 0.15)',
                color: 'var(--text-primary)',
                cursor: 'pointer'
              }}
            >
              ✕ {locale() === 'zh-CN' ? '关闭' : 'Close'}
            </button>
          )}
        </div>
      </header>

      {!isDialog() && (
        <nav class="tab-nav">
          <button
            type="button"
            class="nav-tab-btn"
            classList={{ active: activeTab() === 'spotlight' }}
            onClick={() => setActiveTab('spotlight')}
          >
            <span class="tab-icon">🔦</span>
            <span class="tab-label">{t().tabs?.spotlight || 'Spotlight'}</span>
          </button>

          <button
            type="button"
            class="nav-tab-btn"
            classList={{ active: activeTab() === 'slots' }}
            onClick={() => setActiveTab('slots')}
          >
            <span class="tab-icon">📍</span>
            <span class="tab-label">{t().tabs?.slots || 'Slots'}</span>
          </button>

          <button
            type="button"
            class="nav-tab-btn"
            classList={{ active: activeTab() === 'nav' }}
            onClick={() => setActiveTab('nav')}
          >
            <span class="tab-icon">🧭</span>
            <span class="tab-label">{t().tabs?.nav || 'Nav'}</span>
          </button>

          <button
            type="button"
            class="nav-tab-btn"
            classList={{ active: activeTab() === 'palette' }}
            onClick={() => setActiveTab('palette')}
          >
            <span class="tab-icon">🚀</span>
            <span class="tab-label">CMD</span>
          </button>

          <button
            type="button"
            class="nav-tab-btn"
            classList={{ active: activeTab() === 'settings' }}
            onClick={() => setActiveTab('settings')}
          >
            <span class="tab-icon">⚙️</span>
            <span class="tab-label">{t().tabs?.settings || 'Settings'}</span>
          </button>
        </nav>
      )}

      <main class="app-main">
        {activeTab() === 'spotlight' && <SpotlightTab/>}
        {activeTab() === 'slots' && <PositionSlotsTab/>}
        {activeTab() === 'nav' && <NavigationTab/>}
        {activeTab() === 'palette' && <CommandPalette/>}
        {activeTab() === 'settings' && <SettingsTab/>}
      </main>

      <footer class="app-footer">
        <span class="footer-hint">Kanata Office Add-in • Shared Runtime</span>
      </footer>
    </div>
  );
};
