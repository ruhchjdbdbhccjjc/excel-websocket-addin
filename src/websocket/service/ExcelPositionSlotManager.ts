/**
 * ExcelPositionSlotManager - Fast Position Slot & History Navigation System
 * 
 * Provides:
 * 1. Tier 1: In-memory Map for instantaneous (<1ms) slot lookups and jumps.
 * 2. Tier 2: Office Document Settings persistence via Office.context.document.settings
 *    (zero document pollution, no "recode" sheet created, survives workbook save/reload).
 * 3. Workbook Scope Slots: Jump between arbitrary cells across different worksheets.
 * 4. Sheet Scope Slots: Jump between cell ranges local to the active worksheet.
 * 5. Selection Breadcrumb History (Jump List): Backwards/forwards position stack navigation.
 */

import { CommandResult } from '../protocol/types';
import { Logger } from '../logging/Logger';
import { ExcelSpotlightService } from './ExcelSpotlightService';

export interface PositionSlot {
  slot: string;
  sheetName: string;
  rangeAddress: string;
  fullAddress: string;
  timestamp: number;
}

export interface HistoryEntry {
  sheetName: string;
  rangeAddress: string;
  fullAddress: string;
  timestamp: number;
}

export class ExcelPositionSlotManager {
  private logger: Logger;
  private workbookSlots: Map<string, PositionSlot> = new Map();
  // Key format: `${sheetName}:${slot}`
  private sheetSlots: Map<string, PositionSlot> = new Map();
  // History stack for jump-list navigation
  private historyStack: HistoryEntry[] = [];
  private historyIndex = -1;
  private isTrackerRegistered = false;
  private readonly MAX_HISTORY = 50;
  private isDomFocused = false;
  private lastActiveTimestamp = Date.now();
  private workbookFileName = '';
  private workbookUrl = '';

  constructor(logger?: Logger) {
    this.logger = logger || new Logger();
    this.setupFocusListeners();
    this.loadPersistedSettings();
    const isDialogMode = typeof window !== 'undefined' && Boolean(window.location?.search?.includes('dialog=true'));
    if (!isDialogMode) {
      this.registerSelectionTracker().catch(() => {});
    }
  }

  /**
   * Set up DOM focus listeners for the add-in runtime
   */
  private setupFocusListeners(): void {
    if (typeof window !== 'undefined') {
      window.addEventListener('focus', () => {
        this.isDomFocused = true;
        this.lastActiveTimestamp = Date.now();
        this.logger.debug('Excel add-in window focused');
      });
      window.addEventListener('blur', () => {
        this.isDomFocused = false;
        this.logger.debug('Excel add-in window blurred');
      });
    }
  }

  /**
   * Load active workbook filename and URL via Office Document API
   */
  public async loadWorkbookProperties(): Promise<{ fileName: string; url: string; platform: string }> {
    return new Promise((resolve) => {
      const platform = (typeof Office !== 'undefined' && Office.context && Office.context.diagnostics)
        ? Office.context.diagnostics.platform
        : 'unknown';

      if (
        typeof Office !== 'undefined' &&
        Office.context &&
        Office.context.document &&
        Office.context.document.getFilePropertiesAsync
      ) {
        Office.context.document.getFilePropertiesAsync(null, (res) => {
          if (res && res.value && res.value.url) {
            this.workbookUrl = res.value.url;
            const rawName = res.value.url.substr(res.value.url.lastIndexOf('/') + 1) || 'Workbook.xlsx';
            this.workbookFileName = decodeURIComponent(rawName);
          } else {
            this.workbookFileName = 'Workbook.xlsx';
          }
          resolve({ fileName: this.workbookFileName, url: this.workbookUrl, platform: String(platform) });
        });
      } else {
        this.workbookFileName = 'Workbook.xlsx';
        resolve({ fileName: this.workbookFileName, url: '', platform: String(platform) });
      }
    });
  }

  /**
   * Check if this Excel instance is currently focused or active
   * @param maxIdleMs Milliseconds threshold of user interaction in the Excel workbook
   */
  public isFocused(maxIdleMs = 120000): boolean {
    if (typeof document !== 'undefined' && document.hasFocus && document.hasFocus()) {
      return true;
    }
    if (this.isDomFocused) {
      return true;
    }
    // Excel worksheet grid focus: if the user clicked cells recently (default within 2 minutes)
    return (Date.now() - this.lastActiveTimestamp) < maxIdleMs;
  }

  /**
   * Get comprehensive workbook status & focus details
   */
  public async getWorkbookStatus(): Promise<{
    fileName: string;
    url: string;
    platform: string;
    isFocused: boolean;
    lastActiveAgoMs: number;
    activeSheet?: string;
    activeRange?: string;
  }> {
    await this.loadWorkbookProperties();
    let activeSheet = '';
    let activeRange = '';

    if (this.isExcelAvailable()) {
      try {
        await Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          sheet.load('name');
          const range = context.workbook.getSelectedRange();
          range.load('address');
          await context.sync();
          activeSheet = sheet.name;
          activeRange = range.address;
        });
      } catch {
        // Fallback for non-cell objects (charts, shapes) or multi-selection where getSelectedRange throws InvalidSelection
        try {
          await Excel.run(async (fallbackContext) => {
            const sheet = fallbackContext.workbook.worksheets.getActiveWorksheet();
            sheet.load('name');
            const cell = fallbackContext.workbook.getActiveCell();
            cell.load('address');
            await fallbackContext.sync();
            activeSheet = sheet.name;
            activeRange = cell.address;
          });
        } catch {
          // Ignore in non-Excel context
        }
      }
    }

    const platform = (typeof Office !== 'undefined' && Office.context && Office.context.diagnostics)
      ? Office.context.diagnostics.platform
      : 'unknown';

    return {
      fileName: this.workbookFileName,
      url: this.workbookUrl,
      platform: String(platform),
      isFocused: this.isFocused(),
      lastActiveAgoMs: Date.now() - this.lastActiveTimestamp,
      activeSheet,
      activeRange
    };
  }

  /**
   * Get cached workbook filename
   */
  public getWorkbookFileName(): string {
    return this.workbookFileName || 'Workbook.xlsx';
  }

  /**
   * Load saved slots from Office Document Settings
   */
  public loadPersistedSettings(): void {
    if (
      typeof Office !== 'undefined' &&
      Office.context &&
      Office.context.document &&
      Office.context.document.settings
    ) {
      try {
        const raw = Office.context.document.settings.get('excel_position_slots');
        if (raw && typeof raw === 'string') {
          const parsed = JSON.parse(raw);
          if (Array.isArray(parsed.workbookSlots)) {
            this.workbookSlots = new Map(parsed.workbookSlots);
          }
          if (Array.isArray(parsed.sheetSlots)) {
            this.sheetSlots = new Map(parsed.sheetSlots);
          }
          this.logger.info(`Loaded ${this.workbookSlots.size} workbook slots and ${this.sheetSlots.size} sheet slots from Office document settings`);
        }
      } catch (err) {
        this.logger.warn('Failed to load position slots from document settings:', err);
      }
    }
  }

  /**
   * Persist current slots to Office Document Settings (Zero Sheet Impact)
   */
  private async persistSettings(): Promise<void> {
    if (
      typeof Office !== 'undefined' &&
      Office.context &&
      Office.context.document &&
      Office.context.document.settings
    ) {
      try {
        const payload = {
          workbookName: this.workbookFileName,
          workbookSlots: Array.from(this.workbookSlots.entries()),
          sheetSlots: Array.from(this.sheetSlots.entries()),
          savedAt: Date.now()
        };
        Office.context.document.settings.set('excel_position_slots', JSON.stringify(payload));
        await new Promise<void>((resolve, reject) => {
          Office.context.document.settings.saveAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              resolve();
            } else {
              reject(result.error);
            }
          });
        });
        this.logger.debug('Saved position slots to Office document settings');
      } catch (err) {
        this.logger.warn('Could not save slots to document settings:', err);
      }
    }
  }

  /**
   * Normalize slot key from alphanumeric or symbol
   */
  private normalizeSlotKey(slot: string | number | undefined): string {
    const raw = String(slot || '1').trim().toLowerCase();
    if (raw === 'comma') return ',';
    if (raw === 'period') return '.';
    return raw;
  }

  /**
   * Check if Excel runtime API is available
   */
  private isExcelAvailable(): boolean {
    return typeof Excel !== 'undefined' && typeof Excel.run === 'function';
  }

  /**
   * Push current position into history breadcrumb stack
   */
  public pushHistory(sheetName: string, rangeAddress: string): void {
    const fullAddress = `${sheetName}!${rangeAddress}`;
    const last = this.historyStack[this.historyStack.length - 1];
    if (last && last.fullAddress === fullAddress) {
      return; // Deduplicate consecutive identical positions
    }

    // Trim any forward history if user jumped and made a new selection
    if (this.historyIndex >= 0 && this.historyIndex < this.historyStack.length - 1) {
      this.historyStack = this.historyStack.slice(0, this.historyIndex + 1);
    }

    this.historyStack.push({
      sheetName,
      rangeAddress,
      fullAddress,
      timestamp: Date.now()
    });

    if (this.historyStack.length > this.MAX_HISTORY) {
      this.historyStack.shift();
    }
    this.historyIndex = this.historyStack.length - 1;
  }

  /**
   * Save current selection into a Workbook (Global) Slot
   */
  public async saveWorkbookSlot(rawSlot: string, targetAddress?: string): Promise<CommandResult> {
    if (!this.isExcelAvailable()) return { ok: false, error: 'Excel API not available' };
    const slot = this.normalizeSlotKey(rawSlot);

    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load('name');
      const range = targetAddress ? sheet.getRange(targetAddress) : context.workbook.getSelectedRange();
      range.load('address');
      await context.sync();

      // Clean sheet name from address if present
      const rawAddress = range.address;
      const cleanRange = rawAddress.includes('!') ? rawAddress.split('!')[1] : rawAddress;
      const fullAddress = `${sheet.name}!${cleanRange}`;

      const slotEntry: PositionSlot = {
        slot,
        sheetName: sheet.name,
        rangeAddress: cleanRange,
        fullAddress,
        timestamp: Date.now()
      };

      this.workbookSlots.set(slot, slotEntry);
      await this.persistSettings();

      this.logger.info(`Workbook slot [${slot}] saved:`, slotEntry);
      return {
        ok: true,
        data: {
          slot,
          scope: 'workbook',
          sheet: sheet.name,
          address: cleanRange,
          fullAddress
        }
      };
    });
  }

  /**
   * Switch / Jump to a Workbook (Global) Slot
   */
  public async switchWorkbookSlot(rawSlot: string): Promise<CommandResult> {
    if (!this.isExcelAvailable()) return { ok: false, error: 'Excel API not available' };
    const slot = this.normalizeSlotKey(rawSlot);
    const entry = this.workbookSlots.get(slot);

    if (!entry) {
      return {
        ok: false,
        error: `Workbook slot '${slot}' is empty. Use save_workbook_slot with slot '${slot}' first.`
      };
    }

    return await Excel.run(async (context) => {
      // Record current position before jumping
      const activeSheet = context.workbook.worksheets.getActiveWorksheet();
      activeSheet.load('name');
      const activeRange = context.workbook.getSelectedRange();
      activeRange.load('address');
      await context.sync();

      const curAddress = activeRange.address.includes('!') ? activeRange.address.split('!')[1] : activeRange.address;
      this.pushHistory(activeSheet.name, curAddress);

      // Jump to target sheet and range
      const targetSheet = context.workbook.worksheets.getItemOrNullObject(entry.sheetName);
      await context.sync();

      if (targetSheet.isNullObject) {
        return {
          ok: false,
          error: `Worksheet '${entry.sheetName}' for slot '${slot}' was deleted or renamed.`
        };
      }

      targetSheet.activate();
      const targetRange = targetSheet.getRange(entry.rangeAddress);
      targetRange.select();
      await context.sync();

      this.pushHistory(entry.sheetName, entry.rangeAddress);

      this.logger.info(`Switched to workbook slot [${slot}]: ${entry.fullAddress}`);
      return {
        ok: true,
        data: {
          slot,
          scope: 'workbook',
          sheet: entry.sheetName,
          address: entry.rangeAddress,
          fullAddress: entry.fullAddress
        }
      };
    });
  }

  /**
   * Save current selection into a Sheet-Local Slot
   */
  public async saveSheetSlot(rawSlot: string, targetAddress?: string): Promise<CommandResult> {
    if (!this.isExcelAvailable()) return { ok: false, error: 'Excel API not available' };
    const slot = this.normalizeSlotKey(rawSlot);

    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load('name');
      const range = targetAddress ? sheet.getRange(targetAddress) : context.workbook.getSelectedRange();
      range.load('address');
      await context.sync();

      const rawAddress = range.address;
      const cleanRange = rawAddress.includes('!') ? rawAddress.split('!')[1] : rawAddress;
      const key = `${sheet.name}:${slot}`;

      const slotEntry: PositionSlot = {
        slot,
        sheetName: sheet.name,
        rangeAddress: cleanRange,
        fullAddress: `${sheet.name}!${cleanRange}`,
        timestamp: Date.now()
      };

      this.sheetSlots.set(key, slotEntry);
      await this.persistSettings();

      this.logger.info(`Sheet slot [${key}] saved:`, slotEntry);
      return {
        ok: true,
        data: {
          slot,
          scope: 'sheet',
          sheet: sheet.name,
          address: cleanRange
        }
      };
    });
  }

  /**
   * Switch / Jump to a Sheet-Local Slot
   */
  public async switchSheetSlot(rawSlot: string): Promise<CommandResult> {
    if (!this.isExcelAvailable()) return { ok: false, error: 'Excel API not available' };
    const slot = this.normalizeSlotKey(rawSlot);

    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load('name');
      const activeRange = context.workbook.getSelectedRange();
      activeRange.load('address');
      await context.sync();

      const key = `${sheet.name}:${slot}`;
      const entry = this.sheetSlots.get(key);

      if (!entry) {
        return {
          ok: false,
          error: `Sheet slot '${slot}' on worksheet '${sheet.name}' is empty.`
        };
      }

      const curAddress = activeRange.address.includes('!') ? activeRange.address.split('!')[1] : activeRange.address;
      this.pushHistory(sheet.name, curAddress);

      const targetRange = sheet.getRange(entry.rangeAddress);
      targetRange.select();
      await context.sync();

      this.pushHistory(sheet.name, entry.rangeAddress);

      this.logger.info(`Switched to sheet slot [${key}]: ${entry.rangeAddress}`);
      return {
        ok: true,
        data: {
          slot,
          scope: 'sheet',
          sheet: sheet.name,
          address: entry.rangeAddress
        }
      };
    });
  }

  /**
   * Jump backwards in history stack across the entire workbook
   */
  public async jumpPreviousPosition(): Promise<CommandResult> {
    if (!this.isExcelAvailable()) return { ok: false, error: 'Excel API not available' };
    if (this.historyStack.length === 0) {
      return { ok: false, error: 'Navigation history is empty' };
    }

    if (this.historyIndex > 0) {
      this.historyIndex--;
    } else {
      // Loop to end or stay at oldest
      this.historyIndex = 0;
    }

    const target = this.historyStack[this.historyIndex];
    if (!target) {
      return { ok: false, error: 'No previous position in history' };
    }

    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItemOrNullObject(target.sheetName);
      await context.sync();

      if (sheet.isNullObject) {
        return { ok: false, error: `Sheet '${target.sheetName}' no longer exists` };
      }

      sheet.activate();
      const range = sheet.getRange(target.rangeAddress);
      range.select();
      await context.sync();

      return {
        ok: true,
        data: {
          jumpedTo: target.fullAddress,
          historyIndex: this.historyIndex,
          totalHistory: this.historyStack.length
        }
      };
    });
  }

  /**
   * Jump backwards in history within current worksheet only
   */
  public async jumpPreviousSheetPosition(): Promise<CommandResult> {
    if (!this.isExcelAvailable()) return { ok: false, error: 'Excel API not available' };

    return await Excel.run(async (context) => {
      const currentSheet = context.workbook.worksheets.getActiveWorksheet();
      currentSheet.load('name');
      const curRange = context.workbook.getSelectedRange();
      curRange.load('address');
      await context.sync();

      const curAddress = curRange.address.includes('!') ? curRange.address.split('!')[1] : curRange.address;

      // Find previous history entry matching current sheet name
      let found: HistoryEntry | null = null;
      for (let i = this.historyIndex - 1; i >= 0; i--) {
        const item = this.historyStack[i];
        if (item.sheetName === currentSheet.name && item.rangeAddress !== curAddress) {
          found = item;
          this.historyIndex = i;
          break;
        }
      }

      if (!found) {
        return { ok: false, error: `No earlier position in worksheet '${currentSheet.name}'` };
      }

      const range = currentSheet.getRange(found.rangeAddress);
      range.select();
      await context.sync();

      return {
        ok: true,
        data: {
          sheet: currentSheet.name,
          address: found.rangeAddress,
          historyIndex: this.historyIndex
        }
      };
    });
  }

  /**
   * Reset navigation history pointer to latest position
   */
  public resetHistoryIndex(): CommandResult {
    this.historyIndex = this.historyStack.length > 0 ? this.historyStack.length - 1 : 0;
    return {
      ok: true,
      data: {
        resetIndex: this.historyIndex,
        totalHistory: this.historyStack.length
      }
    };
  }

  /**
   * Clear slots and history
   */
  public async clearSlots(scope: 'all' | 'workbook' | 'sheet' | 'history' = 'all'): Promise<CommandResult> {
    if (scope === 'all' || scope === 'workbook') {
      this.workbookSlots.clear();
    }
    if (scope === 'all' || scope === 'sheet') {
      this.sheetSlots.clear();
    }
    if (scope === 'all' || scope === 'history') {
      this.historyStack = [];
      this.historyIndex = -1;
    }

    await this.persistSettings();
    return { ok: true, data: { cleared: scope } };
  }

  /**
   * List all stored slots
   */
  public getStoredSlots(): CommandResult {
    return {
      ok: true,
      data: {
        workbookSlots: Array.from(this.workbookSlots.values()),
        sheetSlots: Array.from(this.sheetSlots.values()),
        historyCount: this.historyStack.length,
        currentHistoryIndex: this.historyIndex
      }
    };
  }

  /**
   * Register workbook onSelectionChanged listener to automatically build jump history
   */
  public async registerSelectionTracker(): Promise<void> {
    if (this.isTrackerRegistered || !this.isExcelAvailable()) return;

    try {
      await Excel.run(async (context) => {
        context.workbook.onSelectionChanged.add(async (event) => {
          this.lastActiveTimestamp = Date.now();
          try {
            await Excel.run(async (innerContext) => {
              const sheet = innerContext.workbook.worksheets.getActiveWorksheet();
              sheet.load('name');

              // 1. Always load ActiveCell first (safe, never throws InvalidSelection)
              const activeCell = innerContext.workbook.getActiveCell();
              activeCell.load('rowIndex,columnIndex');

              // 2. Query SelectedRange for bounds & multi-cell detection
              const range = innerContext.workbook.getSelectedRange();
              range.load('address,rowIndex,columnIndex,rowCount,columnCount');
              await innerContext.sync();

              const clean = range.address.includes('!') ? range.address.split('!')[1] : range.address;
              this.pushHistory(sheet.name, clean);

              const rowCount = range.rowCount || 1;
              const columnCount = range.columnCount || 1;

              // Notify Spotlight engine with adaptive selection details
              ExcelSpotlightService.getInstance(this.logger).onSelectionChanged({
                activeRow: activeCell.rowIndex,
                activeCol: activeCell.columnIndex,
                sheetName: sheet.name,
                startRow: range.rowIndex,
                endRow: range.rowIndex + rowCount - 1,
                startCol: range.columnIndex,
                endCol: range.columnIndex + columnCount - 1,
                rowCount,
                columnCount,
                isEntireRow: columnCount >= 16384 && rowCount < 1048576,
                isEntireColumn: rowCount >= 1048576 && columnCount < 16384,
                isEntireSheet: rowCount >= 1048576 && columnCount >= 16384,
                isMultiCell: rowCount > 1 || columnCount > 1
              }).catch(() => {});
            });
          } catch (e) {
            // Fallback for non-contiguous multi-selection (Ctrl+click) where getSelectedRange() throws InvalidSelection
            try {
              await Excel.run(async (fallbackContext) => {
                const sheet = fallbackContext.workbook.worksheets.getActiveWorksheet();
                sheet.load('name');
                const activeCell = fallbackContext.workbook.getActiveCell();
                activeCell.load('rowIndex,columnIndex,address');
                await fallbackContext.sync();

                const clean = activeCell.address.includes('!') ? activeCell.address.split('!')[1] : activeCell.address;
                this.pushHistory(sheet.name, clean);

                ExcelSpotlightService.getInstance(this.logger).onSelectionChanged({
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
                }).catch(() => {});
              });
            } catch {
              // Both getSelectedRange() and getActiveCell() threw InvalidSelection.
              // Occurs when user selects a non-cell object (Chart, Shape, Picture, Slicer).
              // Safely clear active spotlight so we don't leave a ghost crosshair on the sheet.
              ExcelSpotlightService.getInstance(this.logger).clearActiveSpotlight().catch(() => {});
            }
          }
        });
        await context.sync();
        this.isTrackerRegistered = true;
        this.logger.info('Workbook onSelectionChanged tracking handler registered');
      });
    } catch (err) {
      this.logger.warn('Could not register onSelectionChanged tracker:', err);
    }
  }
}
