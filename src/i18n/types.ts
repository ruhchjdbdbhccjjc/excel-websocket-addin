export type SupportedLocale = 'zh-CN' | 'en-US';

export interface TranslationSchema {
  common: {
    enabled: string;
    disabled: string;
    save: string;
    clear: string;
    refresh: string;
    copy: string;
    export: string;
    copiedSuccess: string;
    downloadSuccess: string;
    purgedSuccess: string;
  };
  brand: {
    statusConnected: string;
    statusDisconnected: string;
  };
  tabs: {
    spotlight: string;
    slots: string;
    nav: string;
    settings: string;
  };
  spotlight: {
    title: string;
    subtitle: string;
    styleLabel: string;
    styleFill: string;
    styleBorder: string;
    styleBorderDesc: string;
    modeLabel: string;
    modeBoth: string;
    modeRow: string;
    modeCol: string;
    multiSelectLabel: string;
    multiActiveCell: string;
    multiActiveCellTip: string;
    multiBandedArea: string;
    multiBandedAreaTip: string;
    multiAutoHide: string;
    multiAutoHideTip: string;
    colorsLabel: string;
    rulesTitle: string;
    addRule: string;
    emptyRules: string;
    purgeResiduals: string;
    purgeResidualsDesc: string;
    ruleCard: {
      namePlaceholder: string;
      enabled: string;
      paused: string;
      typeLabel: string;
      typeKeyColumn: string;
      typePeerMatch: string;
      colLabel: string;
      colTip: string;
      colPlaceholder: string;
      scopeLabel: string;
      colorLabel: string;
    };
  };
  settings: {
    bridgeTitle: string;
    bridgeDesc: string;
    port: string;
    clientName: string;
    channels: string;
    connect: string;
    disconnect: string;
    stateConnected: string;
    stateDisconnected: string;
    stateConnecting: string;
    languageLabel: string;
    exportLogs: string;
    copyLogs: string;
    consoleLogs: string;
    automationToggle: string;
    automationDesc: string;
  };
}
