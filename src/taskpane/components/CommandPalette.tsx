import { Component, createSignal, createMemo, onMount, For, createEffect, Show } from 'solid-js';
import { BackgroundService } from '../../websocket/service/BackgroundService';
import { ConnectionState } from '../../websocket/client/ConnectionManager';
import { useI18n } from '../../i18n';
import {
  searchAndRankCommands,
  getTotalCommandsCount,
  MsoBilingualCommand
} from '../../data/CommandSearchEngine';

const TOTAL_COUNT = getTotalCommandsCount();
const RECENT_STORAGE_KEY = 'excel_flow_recent_mso_commands';

function loadRecentCommands(): string[] {
  try {
    const raw = localStorage.getItem(RECENT_STORAGE_KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed)) return parsed;
    }
  } catch {
    // ignore
  }
  return ['FreezePanes', 'AutoSum', 'PivotTableInsert', 'RemoveDuplicates', 'FormatPainter', 'Filter', 'Undo', 'Redo'];
}

function saveRecentCommand(idMso: string, current: string[]): string[] {
  const next = [idMso, ...current.filter((x) => x !== idMso)].slice(0, 25);
  try {
    localStorage.setItem(RECENT_STORAGE_KEY, JSON.stringify(next));
  } catch {
    // ignore
  }
  return next;
}

export const CommandPalette: Component = () => {
  const { locale } = useI18n();
  const bgService = BackgroundService.getInstance();
  const connMgr = bgService.getConnectionManager();

  const isZh = () => locale() === 'zh-CN';

  const [searchQuery, setSearchQuery] = createSignal('');
  const [selectedIndex, setSelectedIndex] = createSignal(0);
  const [recentIds, setRecentIds] = createSignal<string[]>(loadRecentCommands());
  const [statusMsg, setStatusMsg] = createSignal(`Ready: ${TOTAL_COUNT} native Excel commands loaded.`);
  const [isConnected, setIsConnected] = createSignal(connMgr.getConnectionState() === ConnectionState.Connected);
  const [searchDuration, setSearchDuration] = createSignal<number>(0);

  let searchInputRef: HTMLInputElement | undefined;
  let listContainerRef: HTMLDivElement | undefined;

  const filteredCommands = createMemo(() => {
    const query = searchQuery();
    const recents = recentIds();
    const t0 = performance.now();
    const results = searchAndRankCommands(query, recents, 60);
    const dur = performance.now() - t0;
    setSearchDuration(Math.round(dur * 100) / 100);
    return results;
  });

  const selectedCommand = createMemo(() => {
    const list = filteredCommands();
    const idx = selectedIndex();
    return list[idx] || null;
  });

  // Ensure scroll follows arrow key navigation
  createEffect(() => {
    const idx = selectedIndex();
    if (!listContainerRef) return;
    if (filteredCommands().length > 0) {
      const items = listContainerRef.children;
      if (items && items[idx]) {
        const el = items[idx] as HTMLElement;
        el.scrollIntoView({ block: 'nearest', behavior: 'auto' });
      }
    }
  });

  onMount(() => {
    setTimeout(() => searchInputRef?.focus(), 60);

    connMgr.on('stateChange', (state) => {
      setIsConnected(state === ConnectionState.Connected);
    });

    connMgr.on('messageReceived', (parsed: any) => {
      const raw = parsed?.raw || parsed;
      const cmd = parsed?.command || raw?.command || (Array.isArray(raw) ? raw[0] : null) || raw?.type;
      if (cmd === 'open_command_palette') {
        setTimeout(() => {
          searchInputRef?.focus();
          searchInputRef?.select();
        }, 50);
      }
      if (raw?.type === 'response' && raw?.result?.idMso) {
        if (raw.status === 'success') {
          setStatusMsg(`✓ Executed [${raw.result.idMso}] via Native COM (Undo supported)`);
        } else {
          setStatusMsg(`✕ COM Error: ${raw.error || 'Failed to execute'}`);
        }
      }
    });
  });

  const isDialog = () => typeof window !== 'undefined' && window.location && window.location.search && window.location.search.includes('dialog=true');

  const closeDialogWindow = (payload?: any) => {
    try {
      if (typeof Office !== 'undefined' && Office.context && Office.context.ui && Office.context.ui.messageParent) {
        Office.context.ui.messageParent(JSON.stringify(payload || { action: 'close' }));
      }
    } catch {
      // ignore
    }
  };

  const executeCommand = (cmd: MsoBilingualCommand) => {
    const msoId = cmd.id;
    // Only persist verified official/registered commands to recent list (prevent typo pollution)
    if (cmd.type !== 'custom') {
      setRecentIds((prev) => saveRecentCommand(msoId, prev));
    }

    // In floating dialog mode, delegate execution to parent shared runtime and close modal immediately
    if (isDialog()) {
      setStatusMsg(`⏳ Executing [${msoId}] via Native COM...`);
      closeDialogWindow({ action: 'execute_mso', idMso: msoId });
      return;
    }

    if (isConnected()) {
      setStatusMsg(`⏳ Executing [${msoId}] via Native COM...`);
      connMgr.send({
        type: 'command',
        command: 'execute_mso',
        id: `mso_${Date.now()}`,
        params: {
          idMso: msoId
        }
      } as any);
    } else {
      setStatusMsg(`⚠ Kanata (:8890) offline. Connect Kanata to execute [${msoId}].`);
    }
  };

  const handleKeyDown = (e: KeyboardEvent) => {
    const list = filteredCommands();
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      setSelectedIndex((prev) => (list.length > 0 ? (prev + 1) % list.length : 0));
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      setSelectedIndex((prev) => (list.length > 0 ? (prev - 1 + list.length) % list.length : 0));
    } else if (e.key === 'Escape') {
      if (searchQuery()) {
        e.preventDefault();
        setSearchQuery('');
      } else if (isDialog()) {
        e.preventDefault();
        closeDialogWindow();
      }
    } else if (e.key === 'Enter') {
      e.preventDefault();
      const chosen = list[selectedIndex()];
      if (chosen) {
        executeCommand(chosen);
      } else if (searchQuery().trim()) {
        const raw = searchQuery().trim();
        executeCommand({
          id: raw,
          type: 'custom',
          tab: 'Custom',
          group: '',
          detail: 'Custom idMso',
          en: { label: raw, supertip: '', acronym: '' },
          zh: { label: raw, supertip: '', pinyin: '', initials: '' },
          search: ''
        });
      }
    }
  };

  return (
    <div class="tab-panel command-palette-tab">
      <div class="card-intro">
        <div style={{ display: 'flex', 'justify-content': 'space-between', 'align-items': 'center' }}>
          <h3 style={{ margin: '0' }}>🚀 Excel Command Palette</h3>
          <span style={{ 'font-size': '11px', opacity: 0.7, 'font-family': 'monospace' }}>
            {searchDuration()}ms ({filteredCommands().length}/{TOTAL_COUNT})
          </span>
        </div>
        <p style={{ 'margin-top': '4px', 'font-size': '12px' }}>
          <strong>Ctrl+Shift+P</strong> (or <strong>Ctrl+Shift+K</strong>) | 拼音 (<code>djcg</code>), 汉字 (<code>冻结</code>), English (<code>fp</code>).
        </p>
      </div>

      <div class="field-card" style={{ 'margin-top': '8px' }}>
        <input
          ref={searchInputRef}
          type="text"
          class="text-input"
          value={searchQuery()}
          onInput={(e) => {
            setSearchQuery(e.currentTarget.value);
            setSelectedIndex(0);
          }}
          onKeyDown={handleKeyDown}
          placeholder={isZh() ? "搜索 2,243 条命令 (例如: djcg, 冻结, fp, pivot)..." : "Search 2,243 commands (e.g. freeze, fp, djcg, pivot)..."}
          style={{ 'font-size': '13px', padding: '8px 10px' }}
        />
      </div>

      <div
        ref={listContainerRef}
        class="command-list-container"
        style={{
          'max-height': '260px',
          'overflow-y': 'auto',
          border: '1px solid rgba(128,128,128,0.25)',
          'border-radius': '6px',
          padding: '4px',
          'margin-bottom': '8px'
        }}
      >
        {filteredCommands().length === 0 ? (
          <div style={{ padding: '14px', opacity: 0.75, 'text-align': 'center' }}>
            {isZh()
              ? <>未找到匹配命令。按 <strong>Enter</strong> 执行原始 idMso: <code>{searchQuery().trim()}</code></>
              : <>No matching commands. Press <strong>Enter</strong> to run idMso: <code>{searchQuery().trim()}</code></>}
          </div>
        ) : (
          <For each={filteredCommands()}>
            {(cmd, idx) => {
              const isSelected = () => idx() === selectedIndex();
              const isRecent = () => recentIds().includes(cmd.id);

              const primary = () => isZh() ? cmd.zh.label : cmd.en.label;
              const secondary = () => isZh() ? cmd.en.label : cmd.zh.label;

              return (
                <div
                  onClick={() => {
                    setSelectedIndex(idx());
                    executeCommand(cmd);
                  }}
                  style={{
                    padding: '6px 10px',
                    'border-radius': '5px',
                    cursor: 'pointer',
                    'margin-bottom': '2px',
                    background: isSelected() ? 'rgba(0, 120, 212, 0.18)' : 'transparent',
                    'border-left': isSelected() ? '3px solid #0078d4' : '3px solid transparent',
                    display: 'flex',
                    'flex-direction': 'column',
                    gap: '2px',
                    transition: 'background 0.05s ease'
                  }}
                >
                  <div style={{ display: 'flex', 'justify-content': 'space-between', 'align-items': 'center' }}>
                    <span style={{ 'font-weight': '600', 'font-size': '13px' }}>
                      {primary()}{' '}
                      <span style={{ opacity: 0.65, 'font-weight': '400', 'font-size': '12px' }}>
                        {primary() !== secondary() ? `(${secondary()})` : ''}
                      </span>
                    </span>
                    <div style={{ display: 'flex', gap: '4px', 'align-items': 'center' }}>
                      {isRecent() && (
                        <span style={{ 'font-size': '10px', background: 'rgba(0,120,212,0.22)', color: '#0078d4', padding: '1px 5px', 'border-radius': '4px', 'font-weight': '600' }}>
                          ★ {isZh() ? '常用' : 'Recent'}
                        </span>
                      )}
                      <span style={{ 'font-size': '10px', opacity: 0.6, background: 'rgba(128,128,128,0.15)', padding: '1px 5px', 'border-radius': '4px', 'font-family': 'monospace' }}>
                        {cmd.id}
                      </span>
                    </div>
                  </div>
                  <div style={{ 'font-size': '11px', opacity: 0.68 }}>
                    {cmd.detail}
                  </div>
                </div>
              );
            }}
          </For>
        )}
      </div>

      {/* Selected Command Supertip / Explanation Preview */}
      <Show when={selectedCommand()}>
        {(cmd) => {
          const tip = () => isZh()
            ? (cmd().zh.supertip || cmd().en.supertip)
            : (cmd().en.supertip || cmd().zh.supertip);

          return (
            <Show when={tip()}>
              <div
                style={{
                  padding: '7px 10px',
                  background: 'rgba(128, 128, 128, 0.08)',
                  border: '1px solid rgba(128, 128, 128, 0.2)',
                  'border-radius': '5px',
                  'font-size': '11px',
                  'line-height': '1.4',
                  'margin-bottom': '8px',
                  'max-height': '68px',
                  'overflow-y': 'auto',
                  opacity: 0.88
                }}
              >
                <strong>💡 {isZh() ? '官方说明: ' : 'Official Tip: '}</strong>
                {tip()}
              </div>
            </Show>
          );
        }}
      </Show>

      <div class="slot-feedback">
        <span class="status-indicator">{isConnected() ? '🟢' : '🟠'}</span>
        <span class="status-msg">{statusMsg()}</span>
      </div>
    </div>
  );
};
