import { TranslationSchema } from '../types';

export const zhCN: TranslationSchema = {
  common: {
    enabled: '已开启',
    disabled: '已关闭',
    save: '保存设置',
    clear: '全部关闭并完全清除',
    refresh: '重新同步光标',
    copy: '复制全部日志',
    export: '导出日志文件',
    copiedSuccess: '✓ 已复制全部日志到剪贴板！',
    downloadSuccess: '✓ 已成功下载日志文件！',
    purgedSuccess: '✓ 已彻底净化全工作簿残留聚光灯！'
  },
  brand: {
    statusConnected: 'Kanata 已连接',
    statusDisconnected: '未连接'
  },
  tabs: {
    spotlight: '聚光灯',
    slots: '位置槽',
    nav: '导航',
    settings: '设置'
  },
  spotlight: {
    title: 'WPS 聚光灯 (Spotlight)',
    subtitle: '动态视线追踪，原生条件格式零损渲染',
    styleLabel: '高亮呈现样式 (Highlight Style):',
    styleFill: '🎨 背景柔光 (Soft Fill)',
    styleBorder: '🛤️ 导轨边框 (Guide Rails)',
    styleBorderDesc: '双导轨细线模式：仅画边缘线，绝不遮盖单元格原本的底色',
    modeLabel: '高亮范围模式 (Direction Mode):',
    modeBoth: '✚ 十字全开 (行+列)',
    modeRow: '━ 仅高亮行 (Row)',
    modeCol: '┃ 仅高亮列 (Col)',
    multiSelectLabel: '多选选区行为 (Multi-Cell Selection):',
    multiActiveCell: '🎯 活动焦点',
    multiActiveCellTip: '💡 框选多格时十字准星以活动焦点格为中心，保持清爽视线。',
    multiBandedArea: '🔲 选区包围',
    multiBandedAreaTip: '💡 十字准星展宽为选区所有行与列的整块带状包围。',
    multiAutoHide: '⏸️ 自动避让',
    multiAutoHideTip: '💡 框选多格时自动隐藏高亮免打扰，重选单格时瞬间恢复。',
    colorsLabel: '预设柔光色彩:',
    rulesTitle: '同类值跨行联动规则',
    addRule: '+ 添加规则',
    emptyRules: '暂无自定义联动规则。点击「+ 添加规则」依据某列同类值跨行联动高亮。',
    purgeResiduals: '🧹 净化全表聚光灯残留',
    purgeResidualsDesc: '扫描全工作簿所有工作表，彻底抹除因之前 Excel 崩溃或异常退出遗留的全部聚光灯规则',
    ruleCard: {
      namePlaceholder: '规则名称',
      enabled: '已启用',
      paused: '已暂停',
      typeLabel: '规则类型:',
      typeKeyColumn: '同行重点列 (点击B7高亮A7)',
      typePeerMatch: '同类值跨行联动 (同行A7+同值行)',
      colLabel: '依据列 (列/序号/名):',
      colTip: '支持输入列标 (如 A)、列序号 (如 1)、表头名 (如 #ID 或 姓名)、或 @ (跟随当前列)',
      colPlaceholder: 'A 或 1 或 #ID',
      scopeLabel: '作用表 (* 为全表):',
      colorLabel: '专属高亮颜色:'
    }
  },
  settings: {
    bridgeTitle: '⚙️ Kanata WebSocket 桥接设置',
    bridgeDesc: '控制与本地 Kanata 服务进程 (Port 8890) 的双向实时通信。',
    port: '监听端口 (Port):',
    clientName: '客户端标识 (Client Name):',
    channels: '订阅频道 (Channels):',
    connect: '连接服务',
    disconnect: '断开连接',
    stateConnected: '已连接 (Online)',
    stateDisconnected: '已断开 (Offline)',
    stateConnecting: '连接中...',
    languageLabel: '界面显示语言 (Language):',
    exportLogs: '📥 导出日志文件',
    copyLogs: '📋 复制全部日志',
    consoleLogs: '实时诊断控制台'
  }
};
