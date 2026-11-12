import { Component, createSignal, For, onMount } from 'solid-js';
import { BackgroundService } from '../../../websocket/service/BackgroundService';

export const PositionSlotsTab: Component = () => {
  const slotManager = BackgroundService.getInstance().getCommandExecutor().getSlotManager();
  const [activeSlot, setActiveSlot] = createSignal<string | null>(null);
  const [feedback, setFeedback] = createSignal<string>('就绪：支持与 Kanata 键盘槽位即时联动');

  const slotGrid = [
    ['u', 'i', 'o'],
    ['j', 'k', 'l'],
    ['m', ',', '.']
  ];

  const handleJump = async (slot: string) => {
    setActiveSlot(slot);
    setFeedback(`正在跳转至槽位 [${slot}]...`);
    const res = await slotManager.switchWorkbookSlot(slot);
    if (res.ok) {
      setFeedback(`✓ 已跳转至槽位 [${slot}]: ${res.data?.fullAddress || ''}`);
    } else {
      setFeedback(`✕ 槽位 [${slot}] 为空或未保存`);
    }
  };

  const handleSave = async (slot: string) => {
    setActiveSlot(slot);
    setFeedback(`正在保存当前选区至槽位 [${slot}]...`);
    const res = await slotManager.saveWorkbookSlot(slot);
    if (res.ok) {
      setFeedback(`✓ 已保存至槽位 [${slot}]: ${res.data?.fullAddress || ''}`);
    } else {
      setFeedback(`✕ 保存失败: ${res.error || ''}`);
    }
  };

  return (
    <div class="tab-panel slots-tab">
      <div class="card-intro">
        <h3>📍 3×3 快速位置槽 (Position Slots)</h3>
        <p>对应 Kanata <code>excel-fast-range-switch</code> 键位，快速锚定单元格。</p>
      </div>

      <div class="slots-grid-wrapper">
        <div class="slots-3x3">
          <For each={slotGrid}>
            {(row) => (
              <div class="slots-row">
                <For each={row}>
                  {(slotKey) => (
                    <div class="slot-tile" classList={{ active: activeSlot() === slotKey }}>
                      <span class="slot-letter">{slotKey}</span>
                      <div class="slot-actions">
                        <button
                          type="button"
                          class="slot-btn jump"
                          title={`跳转至槽位 ${slotKey}`}
                          onClick={() => handleJump(slotKey)}
                        >
                          跳
                        </button>
                        <button
                          type="button"
                          class="slot-btn save"
                          title={`保存当前选区至槽位 ${slotKey}`}
                          onClick={() => handleSave(slotKey)}
                        >
                          存
                        </button>
                      </div>
                    </div>
                  )}
                </For>
              </div>
            )}
          </For>
        </div>
      </div>

      <div class="slot-feedback">
        <span class="status-indicator">ℹ️</span>
        <span class="status-msg">{feedback()}</span>
      </div>
    </div>
  );
};
