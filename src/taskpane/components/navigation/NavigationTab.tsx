import { Component, createSignal } from 'solid-js';
import { ExcelCommandExecutor } from '../../../websocket/service/ExcelCommandExecutor';

export const NavigationTab: Component = () => {
  const executor = new ExcelCommandExecutor();
  const [navSheetName, setNavSheetName] = createSignal('导航');
  const [statusMsg, setStatusMsg] = createSignal('点击下方按钮一键生成/更新全表导航超链接目录');
  const [isProcessing, setIsProcessing] = createSignal(false);

  const handleGenerateNavigation = async () => {
    setIsProcessing(true);
    setStatusMsg('正在生成工作表导航目录...');
    try {
      const res = await executor.execute('set_navigation', { name: navSheetName() });
      if (res.ok) {
        setStatusMsg(`✓ 成功生成/更新导航目录！包含 ${res.data?.totalWorksheets || 0} 个工作表。`);
      } else {
        setStatusMsg(`✕ 生成失败: ${res.error || '未知错误'}`);
      }
    } catch (err: any) {
      setStatusMsg(`✕ 异常: ${err.message || String(err)}`);
    } finally {
      setIsProcessing(false);
    }
  };

  const handleNavigateHome = async () => {
    setStatusMsg('正在返回导航主页...');
    const res = await executor.execute('navigate_home', { name: navSheetName() });
    if (res.ok) {
      setStatusMsg('✓ 已返回导航页！');
    } else {
      setStatusMsg(`✕ 未找到导航表 '${navSheetName()}'`);
    }
  };

  return (
    <div class="tab-panel nav-tab">
      <div class="card-intro">
        <h3>🧭 工作表导航中心 (Table of Contents)</h3>
        <p>自动提取工作簿所有工作表，生成带直达超链接的精美目录索引。</p>
      </div>

      <div class="field-card">
        <label>导航表名称 (Navigation Sheet Name):</label>
        <input
          type="text"
          class="text-input"
          value={navSheetName()}
          onInput={(e) => setNavSheetName(e.currentTarget.value)}
        />
      </div>

      <div class="action-stack">
        <button
          type="button"
          class="action-btn primary"
          disabled={isProcessing()}
          onClick={handleGenerateNavigation}
        >
          {isProcessing() ? '⏳ 处理中...' : '📑 一键生成/更新导航目录'}
        </button>

        <button
          type="button"
          class="action-btn secondary"
          onClick={handleNavigateHome}
        >
          🏠 快速回到导航首页
        </button>
      </div>

      <div class="slot-feedback">
        <span class="status-indicator">📌</span>
        <span class="status-msg">{statusMsg()}</span>
      </div>
    </div>
  );
};
