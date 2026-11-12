import { Component, createSignal, onMount, For, Show } from 'solid-js';
import {
  ExcelSpotlightService,
  SpotlightMode,
  SpotlightStyle,
  SpotlightMultiSelectBehavior,
  SpotlightSettings,
  SpotlightRule,
  WPS_PALETTE
} from '../../../websocket/service/ExcelSpotlightService';
import { useI18n } from '../../../i18n';
import { ColorPills } from './ColorPills';
import { SpotlightRuleCard } from './SpotlightRuleCard';

export const SpotlightTab: Component = () => {
  const { t } = useI18n();
  const spotlight = ExcelSpotlightService.getInstance();
  const [settings, setSettings] = createSignal<SpotlightSettings>(spotlight.getSettings());
  const [feedbackMsg, setFeedbackMsg] = createSignal<string>('');

  onMount(() => {
    setSettings(spotlight.getSettings());
  });

  const showToast = (msg: string) => {
    setFeedbackMsg(msg);
    setTimeout(() => setFeedbackMsg(''), 3000);
  };

  const handleToggleMaster = async () => {
    const updated = await spotlight.toggle();
    setSettings(updated);
  };

  const handleModeChange = async (mode: SpotlightMode) => {
    const updated = await spotlight.configure({ mode });
    setSettings(updated);
  };

  const handleStyleChange = async (style: SpotlightStyle) => {
    const updated = await spotlight.configure({ style });
    setSettings(updated);
  };

  const handleMultiSelectBehaviorChange = async (multiSelectBehavior: SpotlightMultiSelectBehavior) => {
    const updated = await spotlight.configure({ multiSelectBehavior });
    setSettings(updated);
  };

  const handleColorChange = async (color: string) => {
    const updated = await spotlight.configure({ color });
    setSettings(updated);
  };

  const handleAddRule = async () => {
    const isFirst = settings().rules.length === 0;
    const defaultCol = isFirst ? 'A' : String.fromCharCode(65 + (settings().rules.length % 5));
    const colors = [WPS_PALETTE.peach, WPS_PALETTE.green, WPS_PALETTE.blue, WPS_PALETTE.purple];
    const nextColor = colors[settings().rules.length % colors.length];

    const newRule: SpotlightRule = {
      id: `rule_${Date.now()}`,
      name: `重点列 (${defaultCol})`,
      enabled: true,
      type: 'key_column',
      keyColumn: defaultCol,
      color: nextColor,
      sheetScope: '*'
    };

    const rules = [...settings().rules, newRule];
    const updated = await spotlight.configure({ rules });
    setSettings(updated);
  };

  const handleUpdateRule = async (updatedRule: SpotlightRule) => {
    const rules = settings().rules.map((r) => (r.id === updatedRule.id ? updatedRule : r));
    const updated = await spotlight.configure({ rules });
    setSettings(updated);
  };

  const handleDeleteRule = async (ruleId: string) => {
    const rules = settings().rules.filter((r) => r.id !== ruleId);
    const updated = await spotlight.configure({ rules });
    setSettings(updated);
  };

  const handleClear = async () => {
    await spotlight.clearActiveSpotlight();
    await spotlight.toggle(false);
    setSettings(spotlight.getSettings());
  };

  const handleRefresh = async () => {
    await spotlight.refreshSpotlight();
  };

  const handlePurgeResiduals = async () => {
    const count = await spotlight.purgeAllOrphanSpotlightsAcrossWorkbook();
    showToast(`${t().common.purgedSuccess} (${count})`);
  };

  return (
    <div class="tab-panel spotlight-tab">
      {/* 1. Master Hero Switch */}
      <div class="hero-card">
        <div class="hero-header">
          <div class="hero-title">
            <span class="icon">🔦</span>
            <div class="title-text">
              <h3>{t().spotlight.title}</h3>
              <p>{t().spotlight.subtitle}</p>
            </div>
          </div>
          <button
            type="button"
            class="toggle-switch-btn"
            classList={{ active: settings().enabled }}
            onClick={handleToggleMaster}
            title={settings().enabled ? t().common.disabled : t().common.enabled}
          >
            <span class="toggle-knob" />
            <span class="toggle-label">{settings().enabled ? t().common.enabled : t().common.disabled}</span>
          </button>
        </div>
      </div>

      {feedbackMsg() && <div class="feedback-toast">{feedbackMsg()}</div>}

      {/* 2. Highlighting Style (Fill vs Guide Rails) */}
      <div class="control-group">
        <label class="control-label">{t().spotlight.styleLabel}</label>
        <div class="chip-group">
          <button
            type="button"
            class="chip-btn"
            classList={{ active: settings().style === 'fill' }}
            onClick={() => handleStyleChange('fill')}
          >
            {t().spotlight.styleFill}
          </button>
          <button
            type="button"
            class="chip-btn"
            classList={{ active: settings().style === 'border' }}
            onClick={() => handleStyleChange('border')}
            title={t().spotlight.styleBorderDesc}
          >
            {t().spotlight.styleBorder}
          </button>
        </div>
      </div>

      {/* 3. Targeting Mode Segmented Chips */}
      <div class="control-group">
        <label class="control-label">{t().spotlight.modeLabel}</label>
        <div class="chip-group">
          <button
            type="button"
            class="chip-btn"
            classList={{ active: settings().mode === 'both' }}
            onClick={() => handleModeChange('both')}
          >
            {t().spotlight.modeBoth}
          </button>
          <button
            type="button"
            class="chip-btn"
            classList={{ active: settings().mode === 'row_only' }}
            onClick={() => handleModeChange('row_only')}
          >
            {t().spotlight.modeRow}
          </button>
          <button
            type="button"
            class="chip-btn"
            classList={{ active: settings().mode === 'col_only' }}
            onClick={() => handleModeChange('col_only')}
          >
            {t().spotlight.modeCol}
          </button>
        </div>
      </div>

      {/* 4. Multi-Cell Selection Behavior */}
      <div class="control-group">
        <label class="control-label">{t().spotlight.multiSelectLabel}</label>
        <div class="chip-group">
          <button
            type="button"
            class="chip-btn"
            classList={{ active: settings().multiSelectBehavior === 'active_cell' }}
            onClick={() => handleMultiSelectBehaviorChange('active_cell')}
            title={t().spotlight.multiActiveCellTip}
          >
            {t().spotlight.multiActiveCell}
          </button>
          <button
            type="button"
            class="chip-btn"
            classList={{ active: settings().multiSelectBehavior === 'banded_area' }}
            onClick={() => handleMultiSelectBehaviorChange('banded_area')}
            title={t().spotlight.multiBandedAreaTip}
          >
            {t().spotlight.multiBandedArea}
          </button>
          <button
            type="button"
            class="chip-btn"
            classList={{ active: settings().multiSelectBehavior === 'auto_hide' }}
            onClick={() => handleMultiSelectBehaviorChange('auto_hide')}
            title={t().spotlight.multiAutoHideTip}
          >
            {t().spotlight.multiAutoHide}
          </button>
        </div>
        <div class="control-hint">
          {settings().multiSelectBehavior === 'active_cell' && t().spotlight.multiActiveCellTip}
          {settings().multiSelectBehavior === 'banded_area' && t().spotlight.multiBandedAreaTip}
          {settings().multiSelectBehavior === 'auto_hide' && t().spotlight.multiAutoHideTip}
        </div>
      </div>

      {/* 5. WPS Preset Color Palette */}
      <ColorPills
        selectedColor={settings().color}
        onSelectColor={handleColorChange}
      />

      {/* 6. Multi-Rule Builder Section */}
      <div class="rules-section">
        <div class="section-header">
          <div class="header-left">
            <span class="section-icon">✨</span>
            <span class="section-title">{t().spotlight.rulesTitle} ({settings().rules.length})</span>
          </div>
          <button type="button" class="add-rule-btn" onClick={handleAddRule}>
            {t().spotlight.addRule}
          </button>
        </div>

        <Show when={settings().rules.length === 0}>
          <div class="empty-rules-prompt">
            {t().spotlight.emptyRules}
          </div>
        </Show>

        <div class="rules-stack">
          <For each={settings().rules}>
            {(rule) => (
              <SpotlightRuleCard
                rule={rule}
                onUpdate={handleUpdateRule}
                onDelete={handleDeleteRule}
              />
            )}
          </For>
        </div>
      </div>

      {/* 7. Quick Actions */}
      <div class="button-footer">
        <button type="button" class="action-btn secondary" onClick={handleRefresh}>
          🔄 {t().common.refresh}
        </button>
        <button type="button" class="action-btn danger" onClick={handleClear}>
          ✕ {t().common.clear}
        </button>
      </div>

      {/* 8. Crash Recovery & Detox Section */}
      <div class="detox-section" style={{ "margin-top": "12px", "padding-top": "10px", "border-top": "1px dashed #e2e8f0" }}>
        <button
          type="button"
          class="action-btn subtle"
          onClick={handlePurgeResiduals}
          title={t().spotlight.purgeResidualsDesc}
          style={{ width: "100%", "font-size": "11px", color: "#64748b", background: "#f8fafc", border: "1px solid #cbd5e1" }}
        >
          {t().spotlight.purgeResiduals}
        </button>
      </div>
    </div>
  );
};
