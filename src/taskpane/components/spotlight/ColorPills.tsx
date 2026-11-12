import { Component, For } from 'solid-js';
import { WPS_PALETTE } from '../../../websocket/service/ExcelSpotlightService';

interface ColorPillsProps {
  selectedColor?: string;
  activeColor?: string;
  onSelectColor?: (hex: string) => void;
  onSelect?: (hex: string) => void;
  hideLabel?: boolean;
}

export const ColorPills: Component<ColorPillsProps> = (props) => {
  const palette = [
    { name: '经典黄 (Default)', hex: WPS_PALETTE.yellow, border: '#EAB308' },
    { name: '薄荷绿', hex: WPS_PALETTE.green, border: '#22C55E' },
    { name: '晴空蓝', hex: WPS_PALETTE.blue, border: '#38BDF8' },
    { name: '薰衣紫', hex: WPS_PALETTE.purple, border: '#A855F7' },
    { name: '蜜桃橙', hex: WPS_PALETTE.peach, border: '#F43F5E' },
  ];

  const currentColor = () => (props.selectedColor || props.activeColor || '').toUpperCase();
  const handleSelect = (hex: string) => {
    if (props.onSelectColor) props.onSelectColor(hex);
    if (props.onSelect) props.onSelect(hex);
  };

  return (
    <div class="color-pills-container">
      {!props.hideLabel && (
        <div class="color-pills-label">高亮预设色彩 (WPS Pastel Colors):</div>
      )}
      <div class="color-pills-row">
        <For each={palette}>
          {(item) => {
            const isSelected = () => currentColor() === item.hex.toUpperCase();
            return (
              <button
                type="button"
                class="color-pill-btn"
                classList={{ active: isSelected() }}
                style={{
                  'background-color': item.hex,
                  'border-color': isSelected() ? item.border : 'rgba(255,255,255,0.2)'
                }}
                title={item.name}
                onClick={() => handleSelect(item.hex)}
              >
                {isSelected() && <span class="check-mark" style={{ color: item.border }}>✓</span>}
              </button>
            );
          }}
        </For>
      </div>
    </div>
  );
};
