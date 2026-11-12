/**
 * ExcelCommandExecutor - Executes Excel operations and shared runtime actions
 * Dispatched headlessly via WebSocket commands from Kanata or external clients.
 * Implements operations from excel-shared-runtime-scenario with clean snake_case names
 * and backward-compatible aliases for legacy/sample handlers.
 */

import { CommandResult } from '../protocol/types';
import { Logger } from '../logging/Logger';
import { ConnectionManager } from '../client/ConnectionManager';
import { ExcelPositionSlotManager } from './ExcelPositionSlotManager';
import { ExcelSpotlightService, SpotlightSettings } from './ExcelSpotlightService';

export type CommandHandler = (params: Record<string, unknown>) => Promise<CommandResult>;

export class ExcelCommandExecutor {
  private logger: Logger;
  private slotManager: ExcelPositionSlotManager;
  private connManager?: ConnectionManager;
  private commandHandlers: Map<string, CommandHandler> = new Map();
  private aliasMap: Map<string, string> = new Map();
  private isTaskpaneOpen = false;
  private readonly READ_ONLY_COMMANDS = new Set<string>([
    'ping',
    'query_workbook',
    'get_file_info',
    'get_workbook_info',
    'loadfilename',
    'load_file_name',
    'query_focus',
    'connect_service',
    'get_spotlight_status',
    'open_command_palette',
    'show_taskpane',
    'toggle_taskpane',
    'hide_taskpane',
    'open_palette_dialog',
    'open_dialog',
    'close_dialog'
  ]);

  constructor(logger?: Logger, connManager?: ConnectionManager) {
    this.logger = logger || new Logger();
    this.connManager = connManager;
    this.slotManager = new ExcelPositionSlotManager(this.logger);
    this.registerHandlers();
    this.registerAliases();
  }

  public setConnectionManager(connManager: ConnectionManager): void {
    this.connManager = connManager;
  }

  /**
   * Get the position slot manager
   */
  public getSlotManager(): ExcelPositionSlotManager {
    return this.slotManager;
  }

  /**
   * Execute a command by name with given parameters
   */
  public async execute(rawCommandName: string, params: Record<string, unknown> = {}): Promise<CommandResult> {
    const canonicalName = this.resolveCommandName(rawCommandName);

    // 1. Focus filtering:
    // By default, action/navigation commands only execute on the FOCUSED/ACTIVE Excel window.
    // This allows Kanata to simply broadcast commands without specifying {"onlyWhenFocused": true} everywhere.
    // Diagnostic commands (ping, query_workbook, connect_service) can execute without focus.
    const isReadOnly = this.READ_ONLY_COMMANDS.has(canonicalName);
    const requiresFocus = params.onlyWhenFocused !== undefined
      ? Boolean(params.onlyWhenFocused)
      : !isReadOnly;

    if (requiresFocus && !this.slotManager.isFocused()) {
      this.logger.info(`Command '${canonicalName}' skipped: Excel window '${this.slotManager.getWorkbookFileName()}' is not focused`);
      return { ok: false, error: 'excel_not_focused', skipped: true };
    }

    // 2. Target Workbook filtering: match target workbook name if specified
    if (params.targetWorkbook || params.workbook) {
      const targetWb = String(params.targetWorkbook || params.workbook).toLowerCase();
      const currentWb = this.slotManager.getWorkbookFileName().toLowerCase();
      if (!currentWb.includes(targetWb)) {
        this.logger.info(`Command '${canonicalName}' skipped: workbook mismatch ('${currentWb}' does not match '${targetWb}')`);
        return { ok: false, error: 'workbook_mismatch' };
      }
    }

    // 3. Target Platform filtering: desktop (PC) vs web (OfficeOnline)
    if (params.targetPlatform || params.platform) {
      const targetPlat = String(params.targetPlatform || params.platform).toLowerCase();
      const currentPlat = ((typeof Office !== 'undefined' && Office.context && Office.context.diagnostics)
        ? String(Office.context.diagnostics.platform)
        : '').toLowerCase();

      if (targetPlat === 'desktop' && currentPlat !== 'pc') {
        return { ok: false, error: 'platform_mismatch' };
      }
      if (targetPlat === 'web' && currentPlat !== 'officeonline') {
        return { ok: false, error: 'platform_mismatch' };
      }
    }

    const handler = this.commandHandlers.get(canonicalName);

    if (!handler) {
      this.logger.warn(`Unknown command requested: ${rawCommandName} (resolved: ${canonicalName})`);
      return {
        ok: false,
        error: `Unknown command: '${rawCommandName}'. Supported commands include: ${Array.from(this.commandHandlers.keys()).join(', ')}`
      };
    }

    try {
      this.logger.info(`Executing command: ${canonicalName}`, params);
      const result = await handler(params);
      this.logger.info(`Command ${canonicalName} completed successfully:`, result);
      return result;
    } catch (err: any) {
      const errorMsg = err instanceof Error ? err.message : String(err);
      this.logger.error(`Error executing command ${canonicalName}:`, err);
      return {
        ok: false,
        error: `Failed to execute '${canonicalName}': ${errorMsg}`
      };
    }
  }

  /**
   * Resolve any alias to canonical snake_case command name
   */
  public resolveCommandName(name: string): string {
    const trimmed = (name || '').trim();
    const alias = this.aliasMap.get(trimmed.toLowerCase());
    return alias || trimmed.toLowerCase();
  }

  /**
   * Check if Excel API is available
   */
  private isExcelAvailable(): boolean {
    return typeof Excel !== 'undefined' && typeof Excel.run === 'function';
  }

  /**
   * Check if Office runtime API is available
   */
  private isOfficeAvailable(): boolean {
    return typeof Office !== 'undefined';
  }

  private currentDialog: Office.Dialog | null = null;
  private isOpeningDialog = false;
  private lastOpenPaletteTime = 0;
  private lastDialogCloseTime = 0;

  /**
   * Open or toggle the Command Palette in a floating modal dialog window (center screen)
   */
  public async openPaletteDialog(): Promise<CommandResult> {
    const now = Date.now();
    // 350ms debounce against key bounce / duplicate incoming WebSocket messages
    if (now - this.lastOpenPaletteTime < 350) {
      this.logger.debug('openPaletteDialog debounced');
      return { ok: true, data: { status: 'dialog_debounced' } };
    }
    this.lastOpenPaletteTime = now;

    if (typeof Office === 'undefined' || !Office.context || !Office.context.ui || !Office.context.ui.displayDialogAsync) {
      this.logger.warn('Office.context.ui.displayDialogAsync not supported, falling back to showAsTaskpane');
      if (typeof Office !== 'undefined' && Office.addin && Office.addin.showAsTaskpane) {
        await Office.addin.showAsTaskpane();
        this.isTaskpaneOpen = true;
        return { ok: true, data: { status: 'fallback_to_taskpane' } };
      }
      return { ok: false, error: 'displayDialogAsync_not_supported' };
    }

    // 1. Toggle behavior: If an existing dialog is currently tracked as open, close it cleanly
    if (this.currentDialog) {
      this.logger.info('Dialog is already open; closing on toggle request');
      try {
        this.currentDialog.close();
      } catch (err) {
        this.logger.debug('Error closing existing dialog on toggle:', err);
      }
      this.currentDialog = null;
      this.isOpeningDialog = false;
      this.lastDialogCloseTime = Date.now();
      return { ok: true, data: { status: 'dialog_closed' } };
    }

    // 2. Prevent concurrent execution while displayDialogAsync is actively pending
    if (this.isOpeningDialog) {
      this.logger.info('displayDialogAsync already in progress, skipping duplicated invocation');
      return { ok: true, data: { status: 'dialog_already_opening' } };
    }

    // 3. Proactive cool-down buffer (Microsoft Dialog API rule):
    // After closing a dialog, WebView2 needs ~350-400ms to tear down the Win32 window before
    // Office allows opening another dialog without triggering 12007.
    const timeSinceClose = Date.now() - this.lastDialogCloseTime;
    if (timeSinceClose < 400) {
      const waitMs = 400 - timeSinceClose;
      this.logger.debug(`Cooling down ${waitMs}ms after dialog close to avoid 12007...`);
      await new Promise((r) => setTimeout(r, waitMs));
    }

    this.isOpeningDialog = true;
    let watchdogTimer: any = setTimeout(() => {
      this.isOpeningDialog = false;
      this.currentDialog = null;
    }, 8000);

    const refreshWatchdog = () => {
      clearTimeout(watchdogTimer);
      watchdogTimer = setTimeout(() => {
        this.isOpeningDialog = false;
        this.currentDialog = null;
      }, 5000);
    };

    const origin = typeof window !== 'undefined' && window.location && window.location.origin
      ? window.location.origin
      : 'https://localhost:3000';
    // Static URL allows WebView2 to cache the 1.5MB compiled JS bundle for instant modal popup
    const dialogUrl = `${origin}/taskpane.html?dialog=true`;

    const isError12007 = (err: any): boolean => {
      if (!err) return false;
      const c = err.code !== undefined ? err.code : err.error;
      const m = String(err.message || '').toLowerCase();
      return (
        Number(c) === 12007 ||
        String(c).includes('12007') ||
        m.includes('active dialog') ||
        m.includes('already has an active dialog')
      );
    };

    const executeDisplayDialog = (retriesLeft = 12, delayMs = 150): Promise<CommandResult> => {
      return new Promise((resolve) => {
        try {
          Office.context.ui.displayDialogAsync(
            dialogUrl,
            { height: 60, width: 45, promptBeforeOpen: false, displayInIframe: false },
            (asyncResult: Office.AsyncResult<Office.Dialog>) => {
              if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                // Microsoft Official Best Practice: Error 12007 indicates previous dialog is still closing
                if (isError12007(asyncResult.error) && retriesLeft > 0) {
                  refreshWatchdog();
                  this.logger.info(`displayDialogAsync returned 12007 (dialog closing). Retrying in ${delayMs}ms (${retriesLeft} retries left)...`);
                  setTimeout(() => {
                    executeDisplayDialog(retriesLeft - 1, Math.min(delayMs + 75, 450)).then(resolve);
                  }, delayMs);
                  return;
                }

                clearTimeout(watchdogTimer);
                this.isOpeningDialog = false;
                this.currentDialog = null;
                this.lastDialogCloseTime = Date.now();
                this.logger.warn('displayDialogAsync failed, falling back to taskpane:', asyncResult.error);
                if (typeof Office !== 'undefined' && Office.addin && Office.addin.showAsTaskpane) {
                  Office.addin.showAsTaskpane().catch(() => {});
                  this.isTaskpaneOpen = true;
                }
                resolve({ ok: false, error: asyncResult.error?.message, data: { status: 'fallback_to_taskpane' } });
              } else {
                clearTimeout(watchdogTimer);
                this.isOpeningDialog = false;
                const activeDialog = asyncResult.value;
                this.currentDialog = activeDialog;

                activeDialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg: any) => {
                  this.logger.info('Dialog message received:', arg);

                  // 1. Immediately close the dialog window and update timestamp
                  this.lastDialogCloseTime = Date.now();
                  this.currentDialog = null;
                  this.isOpeningDialog = false;

                  try {
                    activeDialog.close();
                  } catch (closeErr) {
                    this.logger.warn('activeDialog.close warning:', closeErr);
                  }

                  // 2. Dispatch command to Kanata via compatible Command envelope
                  try {
                    const msg = typeof arg.message === 'string' ? JSON.parse(arg.message) : arg.message;
                    if (msg?.action === 'execute_mso' && msg?.idMso) {
                      if (this.connManager) {
                        setTimeout(() => {
                          try {
                            this.connManager.send({
                              type: 'command',
                              command: 'execute_mso',
                              id: `mso_${Date.now()}`,
                              params: {
                                idMso: msg.idMso
                              }
                            } as any);
                            this.logger.info(`Dispatched execute_mso [${msg.idMso}] via WebSocket`);
                          } catch (sendErr) {
                            this.logger.error('Error sending execute_mso via connManager:', sendErr);
                          }
                        }, 80);
                      }
                    }
                  } catch (e) {
                    this.logger.error('Error parsing dialog message:', e);
                  }
                });

                activeDialog.addEventHandler(Office.EventType.DialogEventReceived, (arg: any) => {
                  this.logger.info('DialogEventReceived (dialog closed or error):', arg);
                  this.lastDialogCloseTime = Date.now();
                  this.currentDialog = null;
                  this.isOpeningDialog = false;
                });

                resolve({ ok: true, data: { status: 'dialog_opened' } });
              }
            }
          );
        } catch (syncEx: any) {
          clearTimeout(watchdogTimer);
          this.isOpeningDialog = false;
          this.currentDialog = null;
          this.lastDialogCloseTime = Date.now();
          this.logger.error('displayDialogAsync synchronous exception:', syncEx);
          resolve({ ok: false, error: syncEx?.message || String(syncEx) });
        }
      });
    };

    return await executeDisplayDialog();
  }

  /**
   * Register all canonical command handlers
   */
  private registerHandlers(): void {
    // --- Taskpane Visibility ---
    this.commandHandlers.set('toggle_taskpane', async (params) => {
      if (!this.isOfficeAvailable() || !Office.addin) {
        return { ok: false, error: 'Office.addin runtime API not available' };
      }
      const shouldOpen = params.visible !== undefined ? Boolean(params.visible) : !this.isTaskpaneOpen;
      if (shouldOpen) {
        await Office.addin.showAsTaskpane();
        this.isTaskpaneOpen = true;
      } else {
        await Office.addin.hide();
        this.isTaskpaneOpen = false;
      }
      return { ok: true, data: { taskpaneVisible: this.isTaskpaneOpen } };
    });

    this.commandHandlers.set('show_taskpane', async () => {
      if (!this.isOfficeAvailable() || !Office.addin) {
        return { ok: false, error: 'Office.addin runtime API not available' };
      }
      await Office.addin.showAsTaskpane();
      this.isTaskpaneOpen = true;
      return { ok: true, data: { taskpaneVisible: true } };
    });

    this.commandHandlers.set('hide_taskpane', async () => {
      if (!this.isOfficeAvailable() || !Office.addin) {
        return { ok: false, error: 'Office.addin runtime API not available' };
      }
      await Office.addin.hide();
      this.isTaskpaneOpen = false;
      return { ok: true, data: { taskpaneVisible: false } };
    });

    // --- Command Palette & Dialog Handlers ---
    this.commandHandlers.set('open_command_palette', async (params) => {
      // If user specifically requested 'taskpane' mode, show taskpane
      if (params.mode === 'taskpane') {
        if (typeof Office !== 'undefined' && Office.addin && Office.addin.showAsTaskpane) {
          await Office.addin.showAsTaskpane();
          this.isTaskpaneOpen = true;
          return { ok: true, data: { status: 'taskpane_opened' } };
        }
      }

      // Default: Open floating Command Palette modal dialog window
      return await this.openPaletteDialog();
    });

    this.commandHandlers.set('open_palette_dialog', async () => {
      return await this.openPaletteDialog();
    });

    this.commandHandlers.set('open_dialog', async () => {
      return await this.openPaletteDialog();
    });

    this.commandHandlers.set('show_dialog', async () => {
      return await this.openPaletteDialog();
    });

    this.commandHandlers.set('close_dialog', async () => {
      if (this.currentDialog) {
        try {
          this.currentDialog.close();
        } catch (e: any) {
          this.logger.debug('Error closing dialog:', e);
        }
      }
      this.currentDialog = null;
      this.isOpeningDialog = false;
      this.lastDialogCloseTime = Date.now();
      return { ok: true, data: { status: 'dialog_closed' } };
    });

    // --- Startup Behavior ---
    this.commandHandlers.set('set_startup_behavior', async (params) => {
      if (!this.isOfficeAvailable() || !Office.addin || !Office.addin.setStartupBehavior) {
        return { ok: false, error: 'Office.addin.setStartupBehavior not available' };
      }
      const behavior = params.behavior === 'none' || params.enabled === false
        ? Office.StartupBehavior.none
        : Office.StartupBehavior.load;
      await Office.addin.setStartupBehavior(behavior);
      return { ok: true, data: { startupBehavior: behavior } };
    });

    // --- Worksheet Operations ---
    this.commandHandlers.set('create_sheet', async (params) => {
      if (!this.isExcelAvailable()) return { ok: false, error: 'Excel API not available' };

      return await Excel.run(async (context) => {
        let sheetName = typeof params.name === 'string' ? params.name.trim() : '';

        // If name not provided, get value from selected cell (reference sample behaviour)
        if (!sheetName) {
          const selected = context.workbook.getSelectedRange();
          selected.load('values');
          await context.sync();
          if (selected.values && selected.values[0] && selected.values[0][0]) {
            sheetName = String(selected.values[0][0]).trim();
          }
        }

        if (!sheetName) {
          sheetName = `Sheet_${Date.now().toString().slice(-4)}`;
        }

        const existingSheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
        await context.sync();

        let targetSheet: Excel.Worksheet;
        let created = false;

        if (existingSheet.isNullObject) {
          targetSheet = context.workbook.worksheets.add(sheetName);
          created = true;
        } else {
          targetSheet = existingSheet;
        }

        targetSheet.activate();
        await context.sync();

        return {
          ok: true,
          data: { name: sheetName, created, activated: true }
        };
      });
    });

    this.commandHandlers.set('activate_sheet', async (params) => {
      if (!this.isExcelAvailable()) return { ok: false, error: 'Excel API not available' };
      const targetName = String(params.name || params.sheet || '').trim();
      if (!targetName) return { ok: false, error: 'Parameter "name" or "sheet" is required' };

      return await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(targetName);
        sheet.activate();
        sheet.load('name');
        await context.sync();
        return { ok: true, data: { name: sheet.name, activated: true } };
      });
    });

    this.commandHandlers.set('list_sheets', async () => {
      if (!this.isExcelAvailable()) return { ok: false, error: 'Excel API not available' };

      return await Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        sheets.load('items/name,items/position');
        await context.sync();
        const list = sheets.items.map((s) => ({ name: s.name, position: s.position }));
        return { ok: true, data: { sheets: list, count: list.length } };
      });
    });

    // --- Navigation Table of Contents ("导航") ---
    this.commandHandlers.set('set_navigation', async (params) => {
      if (!this.isExcelAvailable()) return { ok: false, error: 'Excel API not available' };
      const navSheetName = String(params.name || '导航').trim();

      return await Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        sheets.load('items/name');
        await context.sync();

        let navSheet = context.workbook.worksheets.getItemOrNullObject(navSheetName);
        await context.sync();

        if (navSheet.isNullObject) {
          navSheet = context.workbook.worksheets.add(navSheetName);
        }

        // 1. Move navigation sheet to the very first tab
        navSheet.position = 0;

        // 2. Clear old content to prevent ghost rows from deleted/renamed sheets
        const usedRange = navSheet.getUsedRangeOrNullObject();
        await context.sync();
        if (!usedRange.isNullObject) {
          usedRange.clear(Excel.ClearApplyTo.all);
        }

        // 3. Title banner
        const titleCell = navSheet.getCell(0, 0);
        titleCell.values = [['工作表导航目录 (Workbook Navigation)']];
        titleCell.format.font.bold = true;
        titleCell.format.font.size = 14;

        // 4. Column Headers (Row 2)
        const headerRange = navSheet.getRange('A2:C2');
        headerRange.values = [['#', '工作表名称 (Sheet Name)', '快速直达 (Jump Link)']];
        headerRange.format.font.bold = true;
        headerRange.format.fill.color = '#F3F4F6';

        // 5. Populate sheets (skipping navigation sheet itself)
        let currentRow = 2; // 0-indexed: row 3 in Excel
        let sheetIndex = 1;

        for (let i = 0; i < sheets.items.length; i++) {
          const currentSheetName = sheets.items[i].name;
          if (currentSheetName === navSheetName) {
            continue; // Skip self
          }

          const escapedName = currentSheetName.replace(/'/g, "''");
          const cellRef = `'${escapedName}'!A1`;

          // Col A: Index
          navSheet.getCell(currentRow, 0).values = [[sheetIndex++]];

          // Col B: Sheet Name with hyperlink
          const nameCell = navSheet.getCell(currentRow, 1);
          nameCell.values = [[currentSheetName]];
          nameCell.hyperlink = {
            textToDisplay: currentSheetName,
            screenTip: `跳转到工作表: ${currentSheetName}`,
            documentReference: cellRef
          };

          // Col C: Direct A1 Jump link
          const jumpCell = navSheet.getCell(currentRow, 2);
          jumpCell.hyperlink = {
            textToDisplay: `直达 A1 (${currentSheetName})`,
            screenTip: `Navigate to ${cellRef}`,
            documentReference: cellRef
          };

          currentRow++;
        }

        // 6. Format and activate
        navSheet.getRangeByIndexes(1, 0, currentRow, 3).format.autofitColumns();
        navSheet.activate();
        await context.sync();

        return {
          ok: true,
          data: {
            navigationSheet: navSheetName,
            totalWorksheets: sheets.items.length - 1
          }
        };
      });
    });

    // --- Cell & Range Formatting / Selection ---
    this.commandHandlers.set('set_cell_color', async (params) => {
      if (!this.isExcelAvailable()) return { ok: false, error: 'Excel API not available' };
      const color = String(params.color || 'yellow');
      const address = params.range || params.cell || params.address ? String(params.range || params.cell || params.address) : null;

      return await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const targetRange = address ? sheet.getRange(address) : context.workbook.getSelectedRange();
        targetRange.format.fill.color = color;
        targetRange.load('address');
        await context.sync();
        return { ok: true, data: { address: targetRange.address, color } };
      });
    });

    this.commandHandlers.set('set_cell_value', async (params) => {
      if (!this.isExcelAvailable()) return { ok: false, error: 'Excel API not available' };
      const val = params.value !== undefined ? params.value : '';
      const address = params.range || params.cell || params.address ? String(params.range || params.cell || params.address) : null;

      return await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const targetRange = address ? sheet.getRange(address) : context.workbook.getSelectedRange();
        
        if (Array.isArray(val)) {
          targetRange.values = val as any[][];
        } else {
          targetRange.values = [[val]];
        }
        targetRange.load('address');
        await context.sync();
        return { ok: true, data: { address: targetRange.address, value: val } };
      });
    });

    this.commandHandlers.set('select_range', async (params) => {
      if (!this.isExcelAvailable()) return { ok: false, error: 'Excel API not available' };
      const address = String(params.range || params.cell || params.address || 'A1');

      return await Excel.run(async (context) => {
        const sheet = params.sheet
          ? context.workbook.worksheets.getItem(String(params.sheet))
          : context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(address);
        range.select();
        range.load('address');
        await context.sync();
        return { ok: true, data: { selectedRange: range.address } };
      });
    });

    this.commandHandlers.set('set_border', async (params) => {
      if (!this.isExcelAvailable()) return { ok: false, error: 'Excel API not available' };
      const explicitColor = params.color ? String(params.color) : null;
      const address = params.range || params.address ? String(params.range || params.address) : null;

      // Color cycle palette (Coral -> Sage -> Purple -> Teal -> Peach -> Clear)
      const BORDER_PALETTE = ['#E26D5C', '#C7CC7A', '#7560BA', '#9DD9D2', '#FFE1A8', '#FFFFFF'];

      return await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        let range: Excel.Range;
        try {
          range = address ? sheet.getRange(address) : context.workbook.getSelectedRange();
          range.load('address');
          await context.sync();
        } catch {
          return { ok: false, error: '当前未处于有效单元格选中状态 (InvalidSelection)' };
        }

        const bottomBorder = range.format.borders.getItem('EdgeBottom' as any);
        bottomBorder.load('color,style');
        await context.sync();

        // 1. Determine next color in the cycle
        let selectedColor = explicitColor;
        if (!selectedColor) {
          const currentColor = (bottomBorder.color || '').toUpperCase();
          const currentStyle = bottomBorder.style;
          let currentIndex = -1;

          if (currentStyle !== 'None') {
            currentIndex = BORDER_PALETTE.findIndex((c) => c.toUpperCase() === currentColor);
          }

          if (currentIndex === -1 || currentIndex >= BORDER_PALETTE.length - 1) {
            selectedColor = BORDER_PALETTE[0];
          } else {
            selectedColor = BORDER_PALETTE[currentIndex + 1];
          }
        }

        // 2. Apply border to all 4 edges
        const edges = ['EdgeBottom', 'EdgeLeft', 'EdgeRight', 'EdgeTop'];
        const isClear = selectedColor.toUpperCase() === '#FFFFFF';

        edges.forEach((edge) => {
          const item = range.format.borders.getItem(edge as any);
          if (isClear) {
            item.style = 'None';
          } else {
            item.style = 'Continuous';
            item.weight = 'Thick';
            item.color = selectedColor!;
          }
        });

        await context.sync();

        // 3. Option 1: Display temporary floating color badge next to active cell
        try {
          const shapes = sheet.shapes;
          shapes.load('items');
          await context.sync();

          // Remove any existing preview badges
          shapes.items.forEach((s) => {
            if (s.name === '__border_preview_badge__') {
              s.delete();
            }
          });

          // Calculate active cell top-right pixel coordinates
          const activeCell = range.getCell(0, 0);
          activeCell.load('address');
          await context.sync();

          const rawAddr = activeCell.address;
          const cleanAddr = rawAddr.includes('!') ? rawAddr.split('!')[1] : rawAddr;
          const boundingRange = sheet.getRange(`A1:${cleanAddr}`);

          boundingRange.load('height,width');
          activeCell.load('height,width');
          await context.sync();

          const badgeLeft = boundingRange.width + 4;
          const badgeTop = Math.max(0, boundingRange.height - activeCell.height);
          const badgeSize = 22;

          const badge = shapes.addGeometricShape(Excel.GeometricShapeType.roundRectangle);
          badge.name = '__border_preview_badge__';
          badge.left = badgeLeft;
          badge.top = badgeTop;
          badge.width = badgeSize;
          badge.height = badgeSize;

          if (isClear) {
            badge.fill.setSolidColor('#FFFFFF');
            badge.lineFormat.color = '#E26D5C';
            badge.lineFormat.weight = 2;
          } else {
            badge.fill.setSolidColor(selectedColor!);
            badge.lineFormat.color = '#222222';
            badge.lineFormat.weight = 1;
          }

          await context.sync();

          // Auto-remove badge after 1.5 seconds so worksheet stays completely clean
          setTimeout(() => {
            if (this.isExcelAvailable()) {
              Excel.run(async (cleanCtx) => {
                const curSheet = cleanCtx.workbook.worksheets.getActiveWorksheet();
                const curShapes = curSheet.shapes;
                curShapes.load('items');
                await cleanCtx.sync();
                curShapes.items.forEach((s) => {
                  if (s.name === '__border_preview_badge__') {
                    s.delete();
                  }
                });
                await cleanCtx.sync();
              }).catch(() => {});
            }
          }, 1500);
        } catch (badgeErr) {
          this.logger.debug('Preview color badge could not be displayed:', badgeErr);
        }

        return {
          ok: true,
          data: {
            address: range.address,
            borderColor: selectedColor,
            isCleared: isClear
          }
        };
      });
    });

    this.commandHandlers.set('insert_row', async (params) => {
      if (!this.isExcelAvailable()) return { ok: false, error: 'Excel API not available' };
      const address = params.range || params.address ? String(params.range || params.address) : null;

      return await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = address ? sheet.getRange(address) : context.workbook.getSelectedRange();
        range.getEntireRow().insert(Excel.InsertShiftDirection.down);
        await context.sync();
        return { ok: true, data: { rowInserted: true } };
      });
    });

    this.commandHandlers.set('insert_column', async (params) => {
      if (!this.isExcelAvailable()) return { ok: false, error: 'Excel API not available' };
      const address = params.range || params.address ? String(params.range || params.address) : null;

      return await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = address ? sheet.getRange(address) : context.workbook.getSelectedRange();
        range.getEntireColumn().insert(Excel.InsertShiftDirection.right);
        await context.sync();
        return { ok: true, data: { columnInserted: true } };
      });
    });

    this.commandHandlers.set('add_arrow_line', async (params) => {
      if (!this.isExcelAvailable()) return { ok: false, error: 'Excel API not available' };
      const startX = Number(params.startX || 100);
      const startY = Number(params.startY || 100);
      const endX = Number(params.endX || 250);
      const endY = Number(params.endY || 200);
      const color = String(params.color || 'red');

      return await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const shapes = sheet.shapes;
        const lineShape = shapes.addLine(startX, startY, endX, endY, Excel.ConnectorType.straight);
        lineShape.lineFormat.color = color;
        const line = lineShape.line;
        line.beginArrowheadLength = Excel.ArrowheadLength.long;
        line.beginArrowheadWidth = Excel.ArrowheadWidth.wide;
        line.beginArrowheadStyle = Excel.ArrowheadStyle.oval;
        line.endArrowheadLength = Excel.ArrowheadLength.long;
        line.endArrowheadWidth = Excel.ArrowheadWidth.wide;
        line.endArrowheadStyle = Excel.ArrowheadStyle.triangle;
        await context.sync();
        return { ok: true, data: { arrowCreated: true, start: [startX, startY], end: [endX, endY] } };
      });
    });

    // --- Data Export & Import ---
    this.commandHandlers.set('export_selection', async (params) => {
      if (!this.isExcelAvailable()) return { ok: false, error: 'Excel API not available' };
      const address = params.range || params.address ? String(params.range || params.address) : null;

      return await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = address ? sheet.getRange(address) : context.workbook.getSelectedRange();
        range.load('address,values,rowCount,columnCount,text');
        await context.sync();
        return {
          ok: true,
          data: {
            address: range.address,
            rowCount: range.rowCount,
            columnCount: range.columnCount,
            values: range.values,
            text: range.text
          }
        };
      });
    });

    this.commandHandlers.set('export_formula', async (params) => {
      if (!this.isExcelAvailable()) return { ok: false, error: 'Excel API not available' };
      const address = params.range || params.address ? String(params.range || params.address) : null;

      return await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = address ? sheet.getRange(address) : context.workbook.getSelectedRange();
        range.load('address,formulas');
        await context.sync();
        return {
          ok: true,
          data: {
            address: range.address,
            formulas: range.formulas
          }
        };
      });
    });

    this.commandHandlers.set('export_merged_range', async (params) => {
      if (!this.isExcelAvailable()) return { ok: false, error: 'Excel API not available' };
      const address = params.range || params.address ? String(params.range || params.address) : null;

      return await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = address ? sheet.getRange(address) : context.workbook.getSelectedRange();
        range.load('address,isMerged');
        await context.sync();
        return {
          ok: true,
          data: {
            address: range.address,
            isMerged: range.isMerged
          }
        };
      });
    });

    this.commandHandlers.set('import_range', async (params) => {
      if (!this.isExcelAvailable()) return { ok: false, error: 'Excel API not available' };
      const data = params.data as any[][];
      if (!Array.isArray(data) || data.length === 0) {
        return { ok: false, error: 'Parameter "data" must be a non-empty 2D array' };
      }

      return await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        let targetRange: Excel.Range;

        if (params.range || params.address) {
          targetRange = sheet.getRange(String(params.range || params.address));
        } else {
          const selected = context.workbook.getSelectedRange();
          selected.load('rowIndex,columnIndex');
          await context.sync();
          targetRange = sheet.getRangeByIndexes(
            selected.rowIndex,
            selected.columnIndex,
            data.length,
            data[0].length
          );
        }

        targetRange.values = data;
        targetRange.format.autofitColumns();
        targetRange.load('address');
        await context.sync();
        return { ok: true, data: { address: targetRange.address, rows: data.length, cols: data[0].length } };
      });
    });

    this.commandHandlers.set('insert_sample_data', async () => {
      if (!this.isExcelAvailable()) return { ok: false, error: 'Excel API not available' };

      return await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const expensesTable = sheet.tables.add('A1:D1', true /*hasHeaders*/);
        expensesTable.name = `ExpensesTable_${Date.now().toString().slice(-4)}`;

        expensesTable.getHeaderRowRange().values = [['Date', 'Merchant', 'Category', 'Amount']];
        expensesTable.rows.add(null, [
          ['1/1/2026', 'The Phone Company', 'Communications', '120'],
          ['1/2/2026', 'Northwind Electric Cars', 'Transportation', '142'],
          ['1/5/2026', 'Best For You Organics Company', 'Groceries', '27'],
          ['1/10/2026', 'Coho Vineyard', 'Restaurant', '33'],
          ['1/11/2026', 'Bellows College', 'Education', '350'],
          ['1/15/2026', 'Trey Research', 'Other', '135'],
          ['1/15/2026', 'Best For You Organics Company', 'Groceries', '97']
        ]);

        sheet.getUsedRange().format.autofitColumns();
        await context.sync();
        return { ok: true, data: { tableName: expensesTable.name, rowsInserted: 7 } };
      });
    });

    this.commandHandlers.set('sum_selected_data', async (params) => {
      if (!this.isExcelAvailable()) return { ok: false, error: 'Excel API not available' };
      const outputCell = String(params.outputCell || params.cell || 'F1');

      return await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const selected = context.workbook.getSelectedRange();
        selected.load('values,address');
        await context.sync();

        let sum = 0;
        let count = 0;
        if (selected.values) {
          for (const row of selected.values) {
            for (const val of row) {
              const num = Number(val);
              if (!isNaN(num) && val !== null && val !== '') {
                sum += num;
                count++;
              }
            }
          }
        }

        const target = sheet.getRange(outputCell);
        target.values = [[sum]];
        target.format.autofitColumns();
        await context.sync();

        return {
          ok: true,
          data: {
            sourceAddress: selected.address,
            outputCell,
            sum,
            numericCount: count
          }
        };
      });
    });

    // --- Hyperlink & Navigation ---
    this.commandHandlers.set('open_hyperlink', async (params) => {
      if (params.url) {
        window.open(String(params.url), '_blank');
        return { ok: true, data: { openedUrl: params.url } };
      }

      if (!this.isExcelAvailable()) return { ok: false, error: 'Excel API not available' };

      return await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load('hyperlink');
        await context.sync();

        if (!range.hyperlink) {
          return { ok: false, error: 'Selected range does not contain a hyperlink' };
        }

        if (range.hyperlink.address) {
          window.open(range.hyperlink.address, '_blank');
          return { ok: true, data: { openedWebUrl: range.hyperlink.address } };
        } else if (range.hyperlink.documentReference) {
          const docRef = range.hyperlink.documentReference;
          // Document references are in format "'SheetName'!A1" or "SheetName!A1"
          const lastExcl = docRef.lastIndexOf('!');
          if (lastExcl !== -1) {
            const rawSheet = docRef.substring(0, lastExcl);
            const cellAddr = docRef.substring(lastExcl + 1) || 'A1';
            const cleanSheetName = rawSheet.replace(/^'|'$/g, '').replace(/''/g, "'");
            const targetSheet = context.workbook.worksheets.getItemOrNullObject(cleanSheetName);
            await context.sync();
            if (targetSheet.isNullObject) {
              return { ok: false, error: `Target worksheet '${cleanSheetName}' does not exist` };
            }
            targetSheet.activate();
            targetSheet.getRange(cellAddr).select();
            await context.sync();
            return { ok: true, data: { navigatedTo: docRef } };
          }
        }

        return { ok: true, data: { hyperlink: range.hyperlink } };
      });
    });

    // --- Position Slot & History Navigation (Fast Excel Jumps) ---
    this.commandHandlers.set('save_workbook_slot', async (params) => {
      const slot = String(
        params.slot ||
        params.id ||
        params.key ||
        params.recoderange ||
        params.readrecoderange ||
        '1'
      );
      const address = params.range || params.address ? String(params.range || params.address) : undefined;
      return await this.slotManager.saveWorkbookSlot(slot, address);
    });

    this.commandHandlers.set('switch_workbook_slot', async (params) => {
      const slot = String(
        params.slot ||
        params.id ||
        params.key ||
        params.readrecoderange ||
        params.recoderange ||
        '1'
      );
      return await this.slotManager.switchWorkbookSlot(slot);
    });

    this.commandHandlers.set('save_sheet_slot', async (params) => {
      const slot = String(
        params.slot ||
        params.id ||
        params.key ||
        params.recodesheetrange ||
        params.readsheetrange ||
        '1'
      );
      const address = params.range || params.address ? String(params.range || params.address) : undefined;
      return await this.slotManager.saveSheetSlot(slot, address);
    });

    this.commandHandlers.set('switch_sheet_slot', async (params) => {
      const slot = String(
        params.slot ||
        params.id ||
        params.key ||
        params.readsheetrange ||
        params.recodesheetrange ||
        '1'
      );
      return await this.slotManager.switchSheetSlot(slot);
    });

    this.commandHandlers.set('jump_previous_position', async () => {
      return await this.slotManager.jumpPreviousPosition();
    });

    this.commandHandlers.set('jump_previous_sheet_position', async () => {
      return await this.slotManager.jumpPreviousSheetPosition();
    });

    this.commandHandlers.set('reset_history_index', async () => {
      return this.slotManager.resetHistoryIndex();
    });

    this.commandHandlers.set('clear_slots', async (params) => {
      const scope = (params.scope as 'all' | 'workbook' | 'sheet' | 'history') || 'all';
      return await this.slotManager.clearSlots(scope);
    });

    this.commandHandlers.set('get_slots', async () => {
      return this.slotManager.getStoredSlots();
    });

    // --- Specialized Sheet & Range Helpers from Reference Sample ---
    this.commandHandlers.set('createsheetswithselectvalues', async () => {
      if (!this.isExcelAvailable()) return { ok: false, error: 'Excel API not available' };

      return await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load('values,address');
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.load('position,name');
        await context.sync();

        if (!range.values || !range.values[0] || !range.values[0][0]) {
          return { ok: false, error: 'Selected cell is empty; cannot create sheet with empty name' };
        }

        const sheetName = String(range.values[0][0]).trim();
        let dataSheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
        await context.sync();

        let created = false;
        if (dataSheet.isNullObject) {
          dataSheet = context.workbook.worksheets.add(sheetName);
          created = true;
        }

        dataSheet.activate();
        range.values = [['']]; // Clear source cell value per reference sample
        dataSheet.position = sheet.position + 1; // Insert right after current sheet
        await context.sync();

        return {
          ok: true,
          data: {
            name: sheetName,
            created,
            position: dataSheet.position,
            clearedSourceAddress: range.address
          }
        };
      });
    });

    // --- WPS-Style Spotlight / 聚光灯 (Crosshair Highlight) ---
    this.commandHandlers.set('toggle_spotlight', async (params) => {
      const spotlight = ExcelSpotlightService.getInstance(this.logger);
      const overrideState = params.enabled !== undefined ? Boolean(params.enabled) : undefined;
      const result = await spotlight.toggle(overrideState);
      return { ok: true, data: result };
    });

    this.commandHandlers.set('set_spotlight', async (params) => {
      const spotlight = ExcelSpotlightService.getInstance(this.logger);
      const update: Partial<SpotlightSettings> = {};

      if (params.enabled !== undefined) update.enabled = Boolean(params.enabled);
      if (params.mode) update.mode = params.mode as any;
      if (params.style) update.style = params.style as any;
      if (params.color) update.color = String(params.color);
      if (params.secondaryColor) update.secondaryColor = String(params.secondaryColor);
      if (params.multiSelectBehavior) update.multiSelectBehavior = params.multiSelectBehavior as any;
      if (Array.isArray(params.rules)) update.rules = params.rules as any;

      const result = await spotlight.configure(update);
      return { ok: true, data: result };
    });

    this.commandHandlers.set('get_spotlight_status', async () => {
      const spotlight = ExcelSpotlightService.getInstance(this.logger);
      return { ok: true, data: spotlight.getSettings() };
    });

    this.commandHandlers.set('clear_spotlight', async () => {
      const spotlight = ExcelSpotlightService.getInstance(this.logger);
      await spotlight.clearActiveSpotlight();
      return { ok: true, data: { cleared: true } };
    });

    this.commandHandlers.set('purge_workbook_spotlight', async () => {
      const spotlight = ExcelSpotlightService.getInstance(this.logger);
      const purgedCount = await spotlight.purgeAllOrphanSpotlightsAcrossWorkbook();
      return { ok: true, data: { purgedCount } };
    });

    this.commandHandlers.set('select_range_by_address', async (params) => {
      if (!this.isExcelAvailable()) return { ok: false, error: 'Excel API not available' };
      const rawAddress = String(params.address || params.range || '').trim();
      if (!rawAddress) return { ok: false, error: 'Parameter "address" or "range" is required' };

      return await Excel.run(async (context) => {
        if (rawAddress.includes('!')) {
          const parts = rawAddress.split('!');
          const sheetName = parts[0].replace(/'/g, '');
          const cellRef = parts[1];
          const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
          await context.sync();
          if (sheet.isNullObject) {
            return { ok: false, error: `Worksheet '${sheetName}' does not exist` };
          }
          sheet.activate();
          sheet.getRange(cellRef).select();
        } else {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          sheet.getRange(rawAddress).select();
        }
        await context.sync();
        return { ok: true, data: { selectedAddress: rawAddress } };
      });
    });

    this.commandHandlers.set('navigate_home', async (params) => {
      if (!this.isExcelAvailable()) return { ok: false, error: 'Excel API not available' };
      const homeSheetName = String(params.sheet || params.name || '导航');

      return await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItemOrNullObject(homeSheetName);
        await context.sync();
        if (sheet.isNullObject) {
          return { ok: false, error: `Home navigation sheet '${homeSheetName}' does not exist` };
        }
        sheet.activate();
        await context.sync();
        return { ok: true, data: { activatedHomeSheet: homeSheetName } };
      });
    });

    this.commandHandlers.set('query_active_cell', async () => {
      if (!this.isExcelAvailable()) return { ok: false, error: 'Excel API not available' };

      return await Excel.run(async (context) => {
        const activeCell = context.workbook.getActiveCell();
        activeCell.load('address,addressLocal,rowIndex,columnIndex,values');
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.load('name');
        await context.sync();

        const cleanAddress = activeCell.address.includes('!')
          ? activeCell.address.split('!')[1]
          : activeCell.address;

        const boundsRange = sheet.getRange(`A1:${cleanAddress}`);
        boundsRange.load('height,width');
        await context.sync();

        return {
          ok: true,
          data: {
            sheet: sheet.name,
            address: activeCell.address,
            addressLocal: activeCell.addressLocal,
            rowIndex: activeCell.rowIndex,
            columnIndex: activeCell.columnIndex,
            value: activeCell.values[0][0],
            pixelOffset: {
              height: boundsRange.height,
              width: boundsRange.width
            }
          }
        };
      });
    });

    this.commandHandlers.set('query_workbook', async () => {
      const status = await this.slotManager.getWorkbookStatus();
      return { ok: true, data: status };
    });

    // --- Utility & Diagnostics ---
    this.commandHandlers.set('connect_service', async () => {
      return {
        ok: true,
        data: {
          connected: true,
          service: 'excel-shared-runtime-websocket',
          timestamp: Date.now()
        }
      };
    });

    this.commandHandlers.set('ping', async () => {
      return {
        ok: true,
        data: {
          pong: true,
          timestamp: Date.now()
        }
      };
    });
  }

  /**
   * Register legacy and Hungarian button name aliases for maximum backward compatibility
   */
  private registerAliases(): void {
    const aliases: Record<string, string> = {
      // Taskpane
      btnhowtaskpane: 'show_taskpane',
      showtaskpane: 'show_taskpane',
      btnopentaskpane: 'show_taskpane',
      open_taskpane: 'show_taskpane',
      btnclosetaskpane: 'hide_taskpane',
      close_taskpane: 'hide_taskpane',

      // Startup
      btnenableaddinstart: 'set_startup_behavior',
      btndisableaddinstart: 'set_startup_behavior',

      // Sheet ops
      btncreatesheet: 'create_sheet',
      createsheetswithselectvalues: 'create_sheet',
      create_sheet_with_name: 'create_sheet',
      gonagivetion: 'activate_sheet',
      active_sheets: 'activate_sheet',

      // Navigation
      btnsetnavigation: 'set_navigation',
      setnagition: 'set_navigation',

      // Cell & Range ops
      set_color: 'set_cell_color',
      btnsetcolor: 'set_cell_color',
      set_value: 'set_cell_value',
      btnsetvalue: 'set_cell_value',
      setborder: 'set_border',
      btnsetborder: 'set_border',
      insertrow: 'insert_row',
      insertcoloumn: 'insert_column',
      insertcolumn: 'insert_column',
      addarrowline: 'add_arrow_line',
      // Taskpane & Dialog Window Aliases
      open_palette: 'open_command_palette',
      command_palette: 'open_command_palette',
      palette: 'open_command_palette',
      showpane: 'show_taskpane',
      show_pane: 'show_taskpane',
      hidepane: 'hide_taskpane',
      hide_pane: 'hide_taskpane',
      togglepane: 'toggle_taskpane',
      toggle_pane: 'toggle_taskpane',
      opendialog: 'open_palette_dialog',
      open_dialog: 'open_palette_dialog',
      show_dialog: 'open_palette_dialog',
      dialog: 'open_palette_dialog',
      close_dialog: 'close_dialog',
      closedialog: 'close_dialog',

      // Data ops
      btnselectexport: 'export_selection',
      btnselectexportformula: 'export_formula',
      btnselectexportrangemerge: 'export_merged_range',
      btnselectimportrange: 'import_range',
      btninsertdata: 'insert_sample_data',
      insertdata: 'insert_sample_data',
      btnsumdata: 'sum_selected_data',

      // Position Slots (Global Workbook Scope)
      recoderange: 'save_workbook_slot',
      save_slot: 'save_workbook_slot',
      save_position_slot: 'save_workbook_slot',
      readrecoderange: 'switch_workbook_slot',
      switch_slot: 'switch_workbook_slot',
      jump_slot: 'switch_workbook_slot',

      // Position Slots (Sheet-Local Scope)
      recodesheetrange: 'save_sheet_slot',
      readsheetrange: 'switch_sheet_slot',

      // History & Jump List Navigation
      rangeprevios: 'jump_previous_position',
      jump_previous: 'jump_previous_position',
      history_back: 'jump_previous_position',
      rangesheetprevios: 'jump_previous_sheet_position',
      jump_previous_sheet: 'jump_previous_sheet_position',
      resetpreviosindex: 'reset_history_index',
      rangepreviosindex: 'reset_history_index',
      resetrecodeinfomation: 'clear_slots',

      // Hyperlinks & Address Navigation
      btnopenhyperlink: 'open_hyperlink',
      openhyperlink: 'open_hyperlink',
      selecterange_by_address: 'select_range_by_address',
      select_by_address: 'select_range_by_address',

      // Navigation Home
      gonagivetion: 'navigate_home',
      go_home: 'navigate_home',

      // Cell & Workbook Queries
      setsharpposition: 'query_active_cell',
      get_active_cell: 'query_active_cell',
      loadfilename: 'query_workbook',
      load_file_name: 'query_workbook',
      get_file_info: 'query_workbook',
      get_workbook_info: 'query_workbook',
      query_focus: 'query_workbook',

      // Service
      btnconnectservice: 'connect_service',
      connectservice: 'connect_service',

      // Spotlight / 聚光灯 Aliases
      spotlight: 'toggle_spotlight',
      toggle_crosshair: 'toggle_spotlight',
      wps_spotlight: 'toggle_spotlight',
      togglespotlight: 'toggle_spotlight',
      setspotlight: 'set_spotlight',
      purge_spotlight: 'purge_workbook_spotlight',
      clean_spotlight: 'purge_workbook_spotlight',
      purge_all_spotlights: 'purge_workbook_spotlight'
    };

    for (const [alias, canonical] of Object.entries(aliases)) {
      this.aliasMap.set(alias.toLowerCase(), canonical);
    }
  }

  /**
   * Get list of all supported command names
   */
  public getSupportedCommands(): string[] {
    return Array.from(this.commandHandlers.keys());
  }
}
