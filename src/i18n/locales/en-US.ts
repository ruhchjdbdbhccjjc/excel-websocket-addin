import { TranslationSchema } from '../types';

export const enUS: TranslationSchema = {
  common: {
    enabled: 'Enabled',
    disabled: 'Disabled',
    save: 'Save Settings',
    clear: 'Turn Off & Clear All',
    refresh: 'Resync Cursor',
    copy: 'Copy All Logs',
    export: 'Export Log File',
    copiedSuccess: '✓ All logs copied to clipboard!',
    downloadSuccess: '✓ Diagnostic log file downloaded!',
    purgedSuccess: '✓ All residual spotlight rules purged from workbook!'
  },
  brand: {
    statusConnected: 'Kanata Online',
    statusDisconnected: 'Disconnected'
  },
  tabs: {
    spotlight: 'Spotlight',
    slots: 'Slots',
    nav: 'Navigate',
    settings: 'Settings'
  },
  spotlight: {
    title: 'WPS Spotlight',
    subtitle: 'Real-time eye tracking via zero-destruction conditional formats',
    styleLabel: 'Highlight Style:',
    styleFill: '🎨 Soft Fill',
    styleBorder: '🛤️ Guide Rails',
    styleBorderDesc: 'Dual guide-rails outline: draws perimeter lines without covering cell fills',
    modeLabel: 'Direction Mode:',
    modeBoth: '✚ Crosshair (Row+Col)',
    modeRow: '━ Row Only',
    modeCol: '┃ Column Only',
    multiSelectLabel: 'Multi-Cell Selection Behavior:',
    multiActiveCell: '🎯 Active Cell',
    multiActiveCellTip: '💡 Crosshair tracks the white active cell within selection for clean alignment.',
    multiBandedArea: '🔲 Banded Area',
    multiBandedAreaTip: '💡 Crosshair expands into a full band covering all selected rows and columns.',
    multiAutoHide: '⏸️ Auto-Hide',
    multiAutoHideTip: '💡 Automatically hides highlight on multi-selection, seamlessly recovers on single click.',
    colorsLabel: 'Preset Pastel Colors:',
    rulesTitle: 'Dynamic Peer Value Rules',
    addRule: '+ Add Rule',
    emptyRules: 'No custom rules yet. Click "+ Add Rule" to link rows sharing identical column values.',
    purgeResiduals: '🧹 Purge Workbook Residuals',
    purgeResidualsDesc: 'Scans all worksheets across the entire workbook and removes any orphaned spotlight formats left from prior crashes or forced exits',
    ruleCard: {
      namePlaceholder: 'Rule Name',
      enabled: 'Enabled',
      paused: 'Paused',
      typeLabel: 'Rule Type:',
      typeKeyColumn: 'Key Column Anchor (e.g. click B7 to highlight A7)',
      typePeerMatch: 'Peer Value Linking (highlight rows with matching value)',
      colLabel: 'Key Column (letter / # / header):',
      colTip: 'Supports column letter (e.g. A), 1-based index (e.g. 1), header name (e.g. #ID), or @ (track active column)',
      colPlaceholder: 'A or 1 or #ID',
      scopeLabel: 'Target Sheet (* for all):',
      colorLabel: 'Rule Highlight Color:'
    }
  },
  settings: {
    bridgeTitle: '⚙️ Kanata WebSocket Bridge',
    bridgeDesc: 'Bidirectional IPC bridge connected to local Kanata daemon (:8890).',
    port: 'Listen Port:',
    clientName: 'Client Identity:',
    channels: 'Subscribed Channels:',
    connect: 'Connect Bridge',
    disconnect: 'Disconnect',
    stateConnected: 'Online (Connected)',
    stateDisconnected: 'Offline (Disconnected)',
    stateConnecting: 'Connecting...',
    languageLabel: 'Display Language:',
    exportLogs: '📥 Export Log File',
    copyLogs: '📋 Copy All Logs',
    consoleLogs: 'Live Diagnostic Console',
    automationToggle: 'Kanata Automation Bridge',
    automationDesc: 'Enable background WebSocket connection and command automation from Kanata.'
  }
};
