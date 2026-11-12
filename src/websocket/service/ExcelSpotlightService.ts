/**
 * ExcelSpotlightService - High-performance WPS-style 聚光灯 (Spotlight) Engine
 * 
 * Provides:
 * 1. Instantaneous crosshair row & column highlight via native Excel Conditional Formatting.
 * 2. Zero-destruction: Never touches user background colors or cell values; 0 Undo stack impact.
 * 3. 3 Targeting Modes: 'both' (Crosshair), 'row_only', and 'col_only'.
 * 4. WPS Iconic Pastel Palette: Soft Yellow, Soft Green, Soft Blue, Soft Lavender, Soft Peach.
 * 5. Multi-Rule Support: Dynamic peer-value matching and sheet-specific scoping.
 * 6. Headless Ready: Runs in Shared Runtime even when Taskpane is closed.
 */

import { Logger } from '../logging/Logger';

export type SpotlightMode = 'both' | 'row_only' | 'col_only';
export type SpotlightStyle = 'fill' | 'border';
export type SpotlightMultiSelectBehavior = 'active_cell' | 'banded_area' | 'auto_hide';

export type SpotlightRuleType = 'key_column' | 'peer_match' | 'crosshair' | 'name_match' | 'index_match' | 'anchor_cell';

export interface SpotlightRule {
  id: string;
  name: string;
  enabled: boolean;
  type: SpotlightRuleType;
  sheetScope?: string | string[]; // '*' for all sheets, or specific sheet name(s)
  keyColumn?: string;            // e.g. 'A', '1', '#ID', or '@' (active column)
  targetNames?: string[];        // e.g. ['Total', 'Subtotal'] for header matches
  color: string;
  secondaryColor?: string;       // Column color if different from row
}

export interface SpotlightSettings {
  enabled: boolean;
  mode: SpotlightMode;
  style: SpotlightStyle;         // 'fill' (背景柔光) vs 'border' (导轨双线边框)
  color: string;
  secondaryColor?: string;
  multiSelectBehavior: SpotlightMultiSelectBehavior; // 'active_cell' (聚焦活动格) vs 'banded_area' (选区包围) vs 'auto_hide' (自动避让)
  rules: SpotlightRule[];
}

export interface SpotlightSelectionInfo {
  activeRow: number;
  activeCol: number;
  sheetName: string;
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
  rowCount: number;
  columnCount: number;
  isEntireRow: boolean;
  isEntireColumn: boolean;
  isEntireSheet: boolean;
  isMultiCell: boolean;
}

export const WPS_PALETTE = {
  yellow: '#FEF08A', // 经典淡黄 (WPS iconic default)
  green: '#DCFCE7',  // 薄荷淡绿
  blue: '#E0F2FE',   // 晴空浅蓝
  purple: '#F3E8FF', // 优雅淡紫
  peach: '#FFE4E6'   // 柔和蜜桃
};

export class ExcelSpotlightService {
  private static instance: ExcelSpotlightService | null = null;
  private logger: Logger;
  private settings: SpotlightSettings = {
    enabled: false,
    mode: 'both',
    style: 'fill',
    color: WPS_PALETTE.yellow,
    multiSelectBehavior: 'active_cell',
    rules: []
  };

  private isListenerActive = false;
  private lastHighlightedSheet = '';
  private lastRow = -1;
  private lastCol = -1;
  private lastStartRow = -1;
  private lastEndRow = -1;
  private lastStartCol = -1;
  private lastEndCol = -1;
  private isUpdating = false;
  private pendingSelection: SpotlightSelectionInfo | null = null;

  private readonly SETTINGS_STORAGE_KEY = 'excel_spotlight_settings_v1';

  public static getInstance(logger?: Logger): ExcelSpotlightService {
    if (!ExcelSpotlightService.instance) {
      ExcelSpotlightService.instance = new ExcelSpotlightService(logger);
    }
    return ExcelSpotlightService.instance;
  }

  private constructor(logger?: Logger) {
    this.logger = logger || new Logger();
    this.loadPersistedSettings();
    this.registerUnloadHooks();
  }

  private registerUnloadHooks(): void {
    if (typeof window === 'undefined') return;
    const handleUnload = () => {
      if (this.settings.enabled) {
        this.clearActiveSpotlight().catch(() => {});
      }
    };
    window.addEventListener('beforeunload', handleUnload);
    window.addEventListener('pagehide', handleUnload);
  }


  private loadPersistedSettings(): void {
    try {
      if (typeof window !== 'undefined' && window.localStorage) {
        const raw = window.localStorage.getItem(this.SETTINGS_STORAGE_KEY);
        if (raw) {
          const parsed = JSON.parse(raw);
          if (parsed && typeof parsed === 'object') {
            this.settings = { ...this.settings, ...parsed };
            this.logger.info('WPS 聚光灯: 已恢复本地持久化配置', this.settings);
          }
        }
      }
    } catch (e) {
      this.logger.warn('Could not load persisted spotlight settings:', e);
    }
  }

  private persistSettings(): void {
    try {
      if (typeof window !== 'undefined' && window.localStorage) {
        window.localStorage.setItem(this.SETTINGS_STORAGE_KEY, JSON.stringify(this.settings));
      }
    } catch (e) {
      this.logger.warn('Could not persist spotlight settings:', e);
    }
  }

  public getSettings(): SpotlightSettings {
    return { ...this.settings };
  }

  public isEnabled(): boolean {
    return this.settings.enabled;
  }

  /**
   * Toggle spotlight state (on -> off, off -> on)
   */
  public async toggle(overrideState?: boolean): Promise<SpotlightSettings> {
    const newState = overrideState !== undefined ? overrideState : !this.settings.enabled;
    this.settings.enabled = newState;
    this.persistSettings();

    if (newState) {
      this.logger.info('WPS 聚光灯: Enabled');
      await this.refreshSpotlight();
    } else {
      this.logger.info('WPS 聚光灯: Disabled');
      await this.clearActiveSpotlight();
    }

    return this.getSettings();
  }

  /**
   * Update spotlight configuration settings
   */
  public async configure(newConfig: Partial<SpotlightSettings>): Promise<SpotlightSettings> {
    if (newConfig.enabled !== undefined) this.settings.enabled = newConfig.enabled;
    if (newConfig.mode !== undefined) this.settings.mode = newConfig.mode;
    if (newConfig.style !== undefined) this.settings.style = newConfig.style;
    if (newConfig.color !== undefined) this.settings.color = newConfig.color;
    if (newConfig.secondaryColor !== undefined) this.settings.secondaryColor = newConfig.secondaryColor;
    if (newConfig.multiSelectBehavior !== undefined) this.settings.multiSelectBehavior = newConfig.multiSelectBehavior;
    if (newConfig.rules !== undefined) this.settings.rules = newConfig.rules;

    this.persistSettings();

    // Reset cache so mode or color change re-applies immediately
    this.lastRow = -1;
    this.lastCol = -1;
    this.lastStartRow = -1;
    this.lastEndRow = -1;
    this.lastStartCol = -1;
    this.lastEndCol = -1;

    if (this.settings.enabled) {
      await this.refreshSpotlight();
    } else {
      await this.clearActiveSpotlight();
    }

    return this.getSettings();
  }

  /**
   * Called on Excel workbook onSelectionChanged event
   * Supports either full SpotlightSelectionInfo or legacy row/col/sheetName parameters
   */
  public async onSelectionChanged(
    target: SpotlightSelectionInfo | number,
    col?: number,
    sheetName?: string
  ): Promise<void> {
    if (!this.settings.enabled) return;

    let sel: SpotlightSelectionInfo;
    if (typeof target === 'number') {
      const row = target;
      const column = col ?? 0;
      const sheet = sheetName ?? '';
      sel = {
        activeRow: row,
        activeCol: column,
        sheetName: sheet,
        startRow: row,
        endRow: row,
        startCol: column,
        endCol: column,
        rowCount: 1,
        columnCount: 1,
        isEntireRow: false,
        isEntireColumn: false,
        isEntireSheet: false,
        isMultiCell: false
      };
    } else {
      sel = target;
    }

    if (this.isUpdating) {
      this.pendingSelection = sel;
      return;
    }

    if (
      this.lastRow === sel.activeRow &&
      this.lastCol === sel.activeCol &&
      this.lastStartRow === sel.startRow &&
      this.lastEndRow === sel.endRow &&
      this.lastStartCol === sel.startCol &&
      this.lastEndCol === sel.endCol &&
      this.lastHighlightedSheet === sel.sheetName
    ) {
      return;
    }

    const prevSheetName = this.lastHighlightedSheet;
    this.lastRow = sel.activeRow;
    this.lastCol = sel.activeCol;
    this.lastStartRow = sel.startRow;
    this.lastEndRow = sel.endRow;
    this.lastStartCol = sel.startCol;
    this.lastEndCol = sel.endCol;
    this.lastHighlightedSheet = sel.sheetName;

    await this.applySpotlightForSelection(sel, prevSheetName);
  }

  /**
   * Refresh spotlight on the current active cell / selection
   */
  public async refreshSpotlight(): Promise<void> {
    if (!this.isExcelAvailable()) return;

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.load('name');
        const activeCell = context.workbook.getActiveCell();
        activeCell.load('rowIndex,columnIndex');

        let selInfo: SpotlightSelectionInfo;
        try {
          const selectedRange = context.workbook.getSelectedRange();
          selectedRange.load('address,rowIndex,columnIndex,rowCount,columnCount');
          await context.sync();

          const rowCount = selectedRange.rowCount || 1;
          const columnCount = selectedRange.columnCount || 1;
          selInfo = {
            activeRow: activeCell.rowIndex,
            activeCol: activeCell.columnIndex,
            sheetName: sheet.name,
            startRow: selectedRange.rowIndex,
            endRow: selectedRange.rowIndex + rowCount - 1,
            startCol: selectedRange.columnIndex,
            endCol: selectedRange.columnIndex + columnCount - 1,
            rowCount,
            columnCount,
            isEntireRow: columnCount >= 16384 && rowCount < 1048576,
            isEntireColumn: rowCount >= 1048576 && columnCount < 16384,
            isEntireSheet: rowCount >= 1048576 && columnCount >= 16384,
            isMultiCell: rowCount > 1 || columnCount > 1
          };
        } catch {
          await context.sync();
          selInfo = {
            activeRow: activeCell.rowIndex,
            activeCol: activeCell.columnIndex,
            sheetName: sheet.name,
            startRow: activeCell.rowIndex,
            endRow: activeCell.rowIndex,
            startCol: activeCell.columnIndex,
            endCol: activeCell.columnIndex,
            rowCount: 1,
            columnCount: 1,
            isEntireRow: false,
            isEntireColumn: false,
            isEntireSheet: false,
            isMultiCell: false
          };
        }

        this.lastRow = -1;
        this.lastCol = -1;
        this.lastStartRow = -1;
        this.lastEndRow = -1;
        this.lastStartCol = -1;
        this.lastEndCol = -1;
        await this.applySpotlightForSelection(selInfo);
      });
    } catch (err: any) {
      if (this.isCellEditError(err)) {
        this.logger.debug('Spotlight refresh deferred: Excel is in cell-editing mode.');
      } else {
        this.logger.error('Could not refresh spotlight:', err);
      }
    }
  }

  /**
   * Core execution: Applies native conditional formatting formula for selection
   */
  private async applySpotlightForSelection(sel: SpotlightSelectionInfo, prevSheetName?: string): Promise<void> {
    if (!this.isExcelAvailable()) return;

    // 1. Edge Case: Entire Sheet selected (Ctrl+A) -> Quietly clear and suspend spotlight to prevent full yellow screen
    if (sel.isEntireSheet) {
      await this.clearActiveSpotlight();
      return;
    }

    // 2. Edge Case: User chose 'auto_hide' behavior and multiple cells are selected -> Temporarily clear spotlight
    if (this.settings.multiSelectBehavior === 'auto_hide' && sel.isMultiCell) {
      await this.clearActiveSpotlight();
      return;
    }

    this.isUpdating = true;

    try {
      await Excel.run(async (context) => {
        // Cross-sheet cleanup: If user switched worksheets, purge residual crosshairs from the previous worksheet
        if (prevSheetName && prevSheetName !== sel.sheetName) {
          try {
            const prevSheet = context.workbook.worksheets.getItemOrNullObject(prevSheetName);
            prevSheet.load('isNullObject');
            await context.sync();
            if (!prevSheet.isNullObject) {
              const prevRange = prevSheet.getRange();
              const prevCfs = prevRange.conditionalFormats;
              const prevItems = await this.getSpotlightItems(context, prevCfs);
              if (prevItems.length > 0) {
                if (prevCfs.items.length === prevItems.length) {
                  prevCfs.clearAll();
                } else {
                  for (let i = prevItems.length - 1; i >= 0; i--) {
                    prevItems[i].delete();
                  }
                }
              }
            }
          } catch (e) {
            this.logger.debug('Previous sheet cleanup skipped or sheet renamed:', e);
          }
        }

        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const targetRange = sheet.getRange();

        // Effective direction mode with smart row/col selection auto-downgrade
        let effectiveMode: SpotlightMode = this.settings.mode;
        if (sel.isEntireRow) {
          effectiveMode = 'row_only';
        } else if (sel.isEntireColumn) {
          effectiveMode = 'col_only';
        }

        const excelRow = sel.activeRow + 1;
        const excelCol = sel.activeCol + 1;

        // Build master formula depending on multiSelectBehavior
        let rowPart = `ROW()=${excelRow}`;
        let colPart = `COLUMN()=${excelCol}`;

        if (this.settings.multiSelectBehavior === 'banded_area' && sel.isMultiCell) {
          const rStart = sel.startRow + 1;
          const rEnd = sel.endRow + 1;
          const cStart = sel.startCol + 1;
          const cEnd = sel.endCol + 1;

          rowPart = rStart === rEnd ? `ROW()=${rStart}` : `AND(ROW()>=${rStart},ROW()<=${rEnd})`;
          colPart = cStart === cEnd ? `COLUMN()=${cStart}` : `AND(COLUMN()>=${cStart},COLUMN()<=${cEnd})`;
        }

        let masterFormula = `=OR(${rowPart},${colPart})`;
        if (effectiveMode === 'row_only') {
          masterFormula = `=${rowPart}`;
        } else if (effectiveMode === 'col_only') {
          masterFormula = `=${colPart}`;
        }

        const mainColor = this.settings.color || WPS_PALETTE.yellow;
        const desiredRules: Array<{
          formula: string;
          style: SpotlightStyle;
          mode: SpotlightMode;
          color: string;
        }> = [];

        // Extra multi-rules (High-priority specific rules placed first so they take precedence over master spotlight)
        // Suppress row-specific key_column & peer_match anchors when entire column is selected to avoid leaking color into cell A1
        if (!sel.isEntireColumn) {
          for (const rule of this.settings.rules) {
            if (!rule.enabled || !this.isSheetInScope(rule.sheetScope, sel.sheetName)) continue;
            const colInfo = this.resolveColumnInfo(rule.keyColumn, sel.activeCol);

            // Type 1: Key Column / Anchor Cell Highlighting (e.g. click B7 -> A7 is highlighted with rule.color)
            if (rule.type === 'key_column' || rule.type === 'anchor_cell' || rule.type === 'index_match') {
              const anchorFormula = `=AND(${rowPart},${colInfo.columnFormulaPart})`;
              desiredRules.push({
                formula: anchorFormula,
                style: this.settings.style,
                mode: 'both', // Full 4-sided box if border style is selected
                color: rule.color
              });
            }

            // Type 2: Peer match (Same value across other rows + highlight key cell in active row)
            if (rule.type === 'peer_match') {
              let peerFormula = '';
              let rowExclusion = `ROW()<>${excelRow}`;
              if (this.settings.multiSelectBehavior === 'banded_area' && sel.isMultiCell && sel.startRow !== sel.endRow) {
                const rStart = sel.startRow + 1;
                const rEnd = sel.endRow + 1;
                rowExclusion = `OR(ROW()<${rStart},ROW()>${rEnd})`;
              }

              if (colInfo.isDynamicHeader && colInfo.headerName) {
                const h = colInfo.headerName;
                peerFormula = `=AND(${rowExclusion},INDEX(1:1048576,ROW(),MATCH("${h}",$1:$1,0))<>"",INDEX(1:1048576,ROW(),MATCH("${h}",$1:$1,0))=INDEX(1:1048576,${excelRow},MATCH("${h}",$1:$1,0)))`;
              } else {
                const keyCol = colInfo.colLetter;
                peerFormula = `=AND(${rowExclusion},$${keyCol}1<>"",INDEX($${keyCol}:$${keyCol},${excelRow})<>"",$${keyCol}1=INDEX($${keyCol}:$${keyCol},${excelRow}))`;
              }

              desiredRules.push({
                formula: peerFormula,
                style: this.settings.style,
                mode: 'row_only',
                color: rule.color
              });

              // 2) Also highlight key cell (e.g. A7) in the active row with rule color
              const selfAnchorFormula = `=AND(${rowPart},${colInfo.columnFormulaPart})`;
              desiredRules.push({
                formula: selfAnchorFormula,
                style: this.settings.style,
                mode: 'both', // Full 4-sided box if border style is selected
                color: rule.color
              });
            }
          }
        }

        // Master spotlight rule (Crosshair or row/column highlight)
        desiredRules.push({
          formula: masterFormula,
          style: this.settings.style,
          mode: effectiveMode,
          color: mainColor
        });

        // 2. Safe Rule Isolation: Load rules and isolate our own spotlight rules
        const cfs = targetRange.conditionalFormats;
        const existingSpotlightItems = await this.getSpotlightItems(context, cfs);
        const K = existingSpotlightItems.length;
        const M = desiredRules.length;

        // 3. Delete surplus spotlight rules in REVERSE order to avoid collection index shift
        if (K > M) {
          const surplus = existingSpotlightItems.slice(M);
          for (let i = surplus.length - 1; i >= 0; i--) {
            surplus[i].delete();
          }
        }

        // 4. High-Performance In-Place Reuse: Update existing rules without collection mutations
        const reuseCount = Math.min(K, M);
        for (let i = 0; i < reuseCount; i++) {
          const cf = existingSpotlightItems[i];
          const d = desiredRules[i];
          cf.custom.rule.formula = d.formula;
          this.applyFormatStyle(cf, d);
        }

        // 5. Add any new rules if count increased
        for (let i = reuseCount; i < M; i++) {
          const d = desiredRules[i];
          const cf = cfs.add(Excel.ConditionalFormatType.custom);
          cf.custom.rule.formula = d.formula;
          this.applyFormatStyle(cf, d);
        }

        await context.sync();
        this.logger.info(`WPS 聚光灯生效: row=${excelRow}, col=${excelCol}, mode=${effectiveMode}, style=${this.settings.style}, color=${mainColor}`);
      });
    } catch (err: any) {
      if (this.isCellEditError(err)) {
        this.logger.debug('Spotlight apply deferred: Excel is in cell-editing mode.');
      } else {
        this.logger.error('Error applying spotlight:', err);
      }
    } finally {
      this.isUpdating = false;
      if (this.pendingSelection) {
        const next = this.pendingSelection;
        this.pendingSelection = null;
        this.onSelectionChanged(next).catch(() => {});
      }
    }
  }

  /**
   * Completely clear spotlight from active worksheet (Safe Isolation)
   */
  public async clearActiveSpotlight(): Promise<void> {
    if (!this.isExcelAvailable()) return;

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const targetRange = sheet.getRange();
        const cfs = targetRange.conditionalFormats;
        const itemsToDelete = await this.getSpotlightItems(context, cfs);

        if (itemsToDelete.length > 0) {
          if (cfs.items.length === itemsToDelete.length) {
            cfs.clearAll();
          } else {
            for (let i = itemsToDelete.length - 1; i >= 0; i--) {
              itemsToDelete[i].delete();
            }
          }
          await context.sync();
        }
        this.logger.info('WPS 聚光灯已清除 (仅删除专属规则，用户规则已完整保留)');
      });
      this.lastRow = -1;
      this.lastCol = -1;
      this.lastStartRow = -1;
      this.lastEndRow = -1;
      this.lastStartCol = -1;
      this.lastEndCol = -1;
      this.lastHighlightedSheet = '';
    } catch (err: any) {
      this.lastRow = -1;
      this.lastCol = -1;
      this.lastStartRow = -1;
      this.lastEndRow = -1;
      this.lastStartCol = -1;
      this.lastEndCol = -1;
      this.lastHighlightedSheet = '';
      if (this.isCellEditError(err)) {
        this.logger.debug('Spotlight clear deferred: Excel is in cell-editing mode.');
      } else {
        this.logger.error('Error clearing spotlight:', err);
      }
    }
  }

  /**
   * Scan all worksheets in the current workbook and clean up any leftover spotlight rules.
   * Solves workbook pollution caused by ungraceful exit, crash, or reload.
   * Batched into minimal IPC round-trips for high performance.
   * Returns the count of purged rules.
   */
  public async purgeAllOrphanSpotlightsAcrossWorkbook(): Promise<number> {
    if (!this.isExcelAvailable()) return 0;

    let totalPurged = 0;
    try {
      await Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        sheets.load('items/name');
        await context.sync();

        const sheetData: Array<{ sheet: Excel.Worksheet; cfs: Excel.ConditionalFormatCollection }> = [];
        for (const sheet of sheets.items) {
          const cfs = sheet.getRange().conditionalFormats;
          cfs.load('items/type');
          sheetData.push({ sheet, cfs });
        }
        await context.sync();

        const customRulesToLoad: Array<{ cf: Excel.ConditionalFormat; cfs: Excel.ConditionalFormatCollection }> = [];
        for (const entry of sheetData) {
          for (const item of entry.cfs.items) {
            if (item.type === Excel.ConditionalFormatType.custom) {
              item.custom.rule.load('formula');
              customRulesToLoad.push({ cf: item, cfs: entry.cfs });
            }
          }
        }

        if (customRulesToLoad.length > 0) {
          await context.sync();
          for (const item of customRulesToLoad) {
            if (this.isSpotlightFormula(item.cf.custom.rule.formula)) {
              item.cf.delete();
              totalPurged++;
            }
          }
          await context.sync();
        }

        this.logger.info(`工作簿聚光灯残留全量清理完成: 共清理 ${totalPurged} 条规则`);
      });
      this.lastRow = -1;
      this.lastCol = -1;
      this.lastStartRow = -1;
      this.lastEndRow = -1;
      this.lastStartCol = -1;
      this.lastEndCol = -1;
      this.lastHighlightedSheet = '';
    } catch (err: any) {
      if (this.isCellEditError(err)) {
        this.logger.debug('Spotlight purge deferred: Excel is in cell-editing mode.');
      } else {
        this.logger.error('Error purging workbook spotlights:', err);
      }
    }
    return totalPurged;
  }

  /**
   * Helper to detect if an Excel error is due to active cell editing mode
   */
  private isCellEditError(err: any): boolean {
    if (!err) return false;
    const code = err.code || err.debugInfo?.code || '';
    const message = err.message || err.debugInfo?.message || '';
    return (
      code === 'InvalidOperationInCellEditMode' ||
      (typeof message === 'string' && (
        message.includes('cell-editing mode') ||
        message.includes('cell edit mode')
      ))
    );
  }

  /**
   * Identifies whether a formula is a spotlight-generated conditional formatting rule.
   * Strictly isolates spotlight rules without false positives on user formulas.
   */
  public isSpotlightFormula(rawFormula?: string): boolean {
    if (!rawFormula) return false;
    const f = rawFormula.toUpperCase().replace(/\s+/g, '');

    const hasSpotlightAnchor =
      f.startsWith('=ROW()=') ||
      f.startsWith('=COLUMN()=') ||
      f.startsWith('=OR(ROW()=') ||
      f.startsWith('=OR(AND(ROW()') ||
      (f.startsWith('=OR(') && (f.includes('ROW()') || f.includes('COLUMN()'))) ||
      f.startsWith('=AND(ROW()') ||
      f.startsWith('=AND(COLUMN()');

    const isSpotlightPeerMatch =
      f.includes('INDEX(1:1048576,ROW(),MATCH(') ||
      (f.includes('INDEX(') && f.includes('ROW()') && (f.includes('ROW()<>') || f.includes('ROW()<') || f.includes('ROW()>')));

    return hasSpotlightAnchor || isSpotlightPeerMatch;
  }

  /**
   * Safe helper to find all spotlight conditional formatting rules on range
   */
  private async getSpotlightItems(
    context: Excel.RequestContext,
    cfs: Excel.ConditionalFormatCollection
  ): Promise<Excel.ConditionalFormat[]> {
    cfs.load('items/type');
    await context.sync();

    const customItems: Excel.ConditionalFormat[] = [];
    for (const item of cfs.items) {
      if (item.type === Excel.ConditionalFormatType.custom) {
        item.custom.rule.load('formula');
        customItems.push(item);
      }
    }
    if (customItems.length === 0) return [];

    await context.sync();
    return customItems.filter((item) => this.isSpotlightFormula(item.custom.rule.formula));
  }

  /**
   * Helper to apply styling (Fill vs Rail Borders) to a conditional format rule
   */
  private applyFormatStyle(
    cf: Excel.ConditionalFormat,
    desc: { style: SpotlightStyle; mode: SpotlightMode; color: string }
  ): void {
    if (desc.style === 'border') {
      const format = cf.custom.format;
      if (desc.mode === 'row_only' || desc.mode === 'both') {
        format.borders.top.style = Excel.BorderLineStyle.continuous;
        format.borders.top.color = desc.color;
        format.borders.bottom.style = Excel.BorderLineStyle.continuous;
        format.borders.bottom.color = desc.color;
      }
      if (desc.mode === 'col_only' || desc.mode === 'both') {
        format.borders.left.style = Excel.BorderLineStyle.continuous;
        format.borders.left.color = desc.color;
        format.borders.right.style = Excel.BorderLineStyle.continuous;
        format.borders.right.color = desc.color;
      }
    } else {
      cf.custom.format.fill.color = desc.color;
    }
  }

  /**
   * Validate whether a sheet name matches the rule's scope
   */
  private isSheetInScope(scope: string | string[] | undefined, sheetName: string): boolean {
    if (!scope || scope === '*' || scope === 'all') return true;
    if (Array.isArray(scope)) {
      return scope.some((s) => s.toLowerCase() === sheetName.toLowerCase());
    }
    if (scope.includes('*')) {
      const regex = new RegExp('^' + scope.replace(/\*/g, '.*') + '$', 'i');
      return regex.test(sheetName);
    }
    return scope.toLowerCase() === sheetName.toLowerCase();
  }

  /**
   * Parse user column input (e.g. 'A', '1', '#ID', or '姓名') into:
   * - colLetter (e.g. 'A')
   * - columnFormulaPart (e.g. 'COLUMN()=1' or 'COLUMN()=COLUMN($A$1)')
   * - isDynamicHeader / headerName
   */
  private resolveColumnInfo(
    colInput: string | undefined,
    currentColIndex: number
  ): { colLetter: string; columnFormulaPart: string; isDynamicHeader: boolean; headerName?: string } {
    const raw = (colInput || '').trim();
    if (!raw || raw === '@' || raw === '*' || raw.toUpperCase() === 'AUTO') {
      const letter = this.getColumnLetter(currentColIndex);
      return {
        colLetter: letter,
        columnFormulaPart: `COLUMN()=${currentColIndex + 1}`,
        isDynamicHeader: false
      };
    }

    // 1. Explicit header prefix '#' (e.g. '#ID') or Chinese/special characters
    if (raw.startsWith('#') || /[^\w]/.test(raw) || /[\u4e00-\u9fa5]/.test(raw)) {
      const header = raw.replace(/^#/, '').trim();
      const escaped = header.replace(/"/g, '""');
      return {
        colLetter: 'A',
        columnFormulaPart: `COLUMN()=IFERROR(MATCH("${escaped}",$1:$1,0),0)`,
        isDynamicHeader: true,
        headerName: escaped
      };
    }

    // 2. Pure number: e.g. "1", "2" -> clamp safely between 1 and 16384 (XFD)
    if (/^\d+$/.test(raw)) {
      let colNum = parseInt(raw, 10);
      if (isNaN(colNum) || colNum < 1) colNum = 1;
      if (colNum > 16384) colNum = 16384;
      const colLetter = this.getColumnLetter(colNum - 1);
      return {
        colLetter,
        columnFormulaPart: `COLUMN()=${colNum}`,
        isDynamicHeader: false
      };
    }

    // 3. Single column letter (A-Z) -> 100% column letter
    if (/^[A-Za-z]$/.test(raw)) {
      const colLetter = raw.toUpperCase();
      return {
        colLetter,
        columnFormulaPart: `COLUMN()=COLUMN($${colLetter}$1)`,
        isDynamicHeader: false
      };
    }

    // 4. 2 to 3 letters: Check if valid Excel column (<= 16384 XFD)
    if (/^[A-Za-z]{2,3}$/.test(raw)) {
      const upper = raw.toUpperCase();
      const colIndex = this.colLetterToIndex(upper);
      if (colIndex < 16384) {
        return {
          colLetter: upper,
          columnFormulaPart: `COLUMN()=COLUMN($${upper}$1)`,
          isDynamicHeader: false
        };
      }
    }

    // 5. Default fallback: Header name match
    const escaped = raw.replace(/"/g, '""');
    return {
      colLetter: 'A',
      columnFormulaPart: `COLUMN()=IFERROR(MATCH("${escaped}",$1:$1,0),0)`,
      isDynamicHeader: true,
      headerName: escaped
    };
  }

  /**
   * Convert Excel column letters to 0-based column index (A->0, Z->25, AA->26)
   */
  private colLetterToIndex(letter: string): number {
    let sum = 0;
    const upper = letter.toUpperCase();
    for (let i = 0; i < upper.length; i++) {
      sum = sum * 26 + (upper.charCodeAt(i) - 64);
    }
    return Math.max(0, sum - 1);
  }

  /**
   * Convert 0-indexed column number to Excel column letters (0->A, 25->Z, 26->AA)
   */
  private getColumnLetter(colIndex: number): string {
    let temp = colIndex + 1;
    let letter = '';
    while (temp > 0) {
      const mod = (temp - 1) % 26;
      letter = String.fromCharCode(65 + mod) + letter;
      temp = Math.floor((temp - mod) / 26);
    }
    return letter || 'A';
  }

  private isExcelAvailable(): boolean {
    return typeof Excel !== 'undefined' && typeof Excel.run === 'function';
  }
}
