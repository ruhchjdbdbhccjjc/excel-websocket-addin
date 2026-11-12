import { Component } from 'solid-js';
import { SpotlightRule, SpotlightRuleType } from '../../../websocket/service/ExcelSpotlightService';
import { useI18n } from '../../../i18n';
import { ColorPills } from './ColorPills';

interface Props {
  rule: SpotlightRule;
  onUpdate: (updated: SpotlightRule) => void;
  onDelete: (id: string) => void;
}

export const SpotlightRuleCard: Component<Props> = (props) => {
  const { t } = useI18n();

  const handleToggle = () => {
    props.onUpdate({ ...props.rule, enabled: !props.rule.enabled });
  };

  const handleNameChange = (name: string) => {
    props.onUpdate({ ...props.rule, name });
  };

  const isDefaultName = (name: string) => {
    return !name || name.startsWith('同行重点列') || name.startsWith('同类值') || name.startsWith('Key Column') || name.startsWith('Peer Value');
  };

  const handleTypeChange = (type: SpotlightRuleType) => {
    const col = props.rule.keyColumn || 'A';
    const newName = isDefaultName(props.rule.name)
      ? type === 'key_column'
        ? `${t().spotlight.ruleCard.typeKeyColumn.split(' ')[0]} (${col})`
        : `${t().spotlight.ruleCard.typePeerMatch.split(' ')[0]} (${col})`
      : props.rule.name;
    props.onUpdate({ ...props.rule, type, name: newName });
  };

  const handleKeyColChange = (keyColumn: string) => {
    const trimmed = keyColumn.trim();
    const col = trimmed || 'A';
    const newName = isDefaultName(props.rule.name)
      ? props.rule.type === 'peer_match'
        ? `${t().spotlight.ruleCard.typePeerMatch.split(' ')[0]} (${col})`
        : `${t().spotlight.ruleCard.typeKeyColumn.split(' ')[0]} (${col})`
      : props.rule.name;
    props.onUpdate({ ...props.rule, keyColumn: trimmed, name: newName });
  };

  const handleScopeChange = (sheetScope: string) => {
    props.onUpdate({ ...props.rule, sheetScope });
  };

  const handleColorChange = (color: string) => {
    props.onUpdate({ ...props.rule, color });
  };

  return (
    <div class="rule-card" classList={{ active: props.rule.enabled }}>
      <div class="rule-card-header">
        <input
          type="text"
          class="rule-name-input"
          value={props.rule.name}
          onInput={(e) => handleNameChange(e.currentTarget.value)}
          placeholder={t().spotlight.ruleCard.namePlaceholder}
        />
        <div class="rule-header-actions">
          <button
            type="button"
            class="rule-toggle-btn"
            classList={{ enabled: props.rule.enabled }}
            onClick={handleToggle}
            title={props.rule.enabled ? t().spotlight.ruleCard.paused : t().spotlight.ruleCard.enabled}
          >
            {props.rule.enabled ? t().spotlight.ruleCard.enabled : t().spotlight.ruleCard.paused}
          </button>
          <button
            type="button"
            class="rule-delete-btn"
            onClick={() => props.onDelete(props.rule.id)}
            title="Delete rule"
          >
            ✕
          </button>
        </div>
      </div>

      <div class="rule-card-body">
        <div class="rule-field-row">
          <div class="rule-field" style={{ flex: 1 }}>
            <label>{t().spotlight.ruleCard.typeLabel}</label>
            <select
              class="rule-type-select"
              value={props.rule.type || 'key_column'}
              onChange={(e) => handleTypeChange(e.currentTarget.value as SpotlightRuleType)}
            >
              <option value="key_column">{t().spotlight.ruleCard.typeKeyColumn}</option>
              <option value="peer_match">{t().spotlight.ruleCard.typePeerMatch}</option>
            </select>
          </div>
        </div>

        <div class="rule-field-row">
          <div class="rule-field col-field">
            <label title={t().spotlight.ruleCard.colTip}>{t().spotlight.ruleCard.colLabel}</label>
            <input
              type="text"
              class="rule-col-input"
              value={props.rule.keyColumn || 'A'}
              placeholder={t().spotlight.ruleCard.colPlaceholder}
              title={t().spotlight.ruleCard.colTip}
              onInput={(e) => handleKeyColChange(e.currentTarget.value)}
            />
          </div>
          <div class="rule-field scope-field">
            <label>{t().spotlight.ruleCard.scopeLabel}</label>
            <input
              type="text"
              class="rule-scope-input"
              value={String(props.rule.sheetScope || '*')}
              onInput={(e) => handleScopeChange(e.currentTarget.value)}
            />
          </div>
        </div>

        <div class="rule-field color-field">
          <label>{t().spotlight.ruleCard.colorLabel}</label>
          <ColorPills selectedColor={props.rule.color} onSelectColor={handleColorChange} hideLabel={true} />
        </div>
      </div>
    </div>
  );
};
