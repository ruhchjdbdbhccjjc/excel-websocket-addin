# Kanata WebSocket Excel Commands Reference

This document describes all WebSocket commands available in the **Excel Office Add-in** (`excel_websocket_Offic_Add_in`), matching the protocol format used by [`chrome_websocket_extension`](file:///c:/startup/kanata/websocket/chrome_websocket_extension) and [`cursor_websocket_extension`](file:///c:/startup/kanata/websocket/cursor_websocket_extension).

---

## 1. Protocol Architecture & Connection Details

- **Host**: `127.0.0.1`
- **Default Port**: `8890`
- **Default Client Name**: `excel`
- **Default Subscribed Channels**:
  - `excel_websocket_Offic_Add_in` *(Primary — matches extension folder name)*
  - `excel_websocket_extension` *(Standard extension naming)*
  - `excel-addin-websocket` *(Legacy fallback)*

### A. Message Envelope Format (Kanata → Excel Add-in)
```json
{
  "type": "command",
  "id": "req-001",
  "command": "<command_name>",
  "params": { ... }
}
```

### B. Response Envelope Format (Excel Add-in → Kanata)
```json
{
  "type": "response",
  "id": "req-001",
  "status": "success",
  "result": { ... }
}
```
*(On failure: `"status": "error"`, `"error": "Description of error"`).*

### C. Registration Payload on Connect
```json
{
  "type": "register",
  "instanceId": "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx",
  "version": "1.0.0",
  "clientType": "office-addin",
  "clientName": "excel",
  "timestamp": 1728604800000,
  "capabilities": ["receive-messages"],
  "channels": ["excel_websocket_Offic_Add_in", "excel-addin-websocket", "excel_websocket_extension"]
}
```

---

## 2. Command Index

All commands support both **Canonical `snake_case`** (recommended) and **Legacy / Reference Sample Aliases**.

### A. Taskpane & Runtime UI
| Canonical Command | Aliases | Params | Description |
| :--- | :--- | :--- | :--- |
| `toggle_taskpane` | `showtaskpane`, `open_taskpane` | `{ visible?: boolean }` | Toggles or sets taskpane visibility |
| `show_taskpane` | `btnhowtaskpane`, `btnopentaskpane` | `{}` | Shows the shared runtime taskpane |
| `hide_taskpane` | `btnclosetaskpane`, `close_taskpane` | `{}` | Hides the taskpane |
| `set_startup_behavior` | `btnenableaddinstart`, `btndisableaddinstart` | `{ behavior: 'load' \| 'none' }` | Auto-starts add-in on document load |

### B. Worksheet & Navigation
| Canonical Command | Aliases | Params | Description |
| :--- | :--- | :--- | :--- |
| `create_sheet` | `btncreatesheet`, `create_sheet_with_name` | `{ name?: string }` | Creates sheet by name (or selected cell) and activates it |
| `createsheetswithselectvalues` | - | `{}` | Creates sheet from selected cell value, inserts at `sheet.position + 1`, and clears source cell |
| `activate_sheet` | `active_sheets` | `{ name: string }` | Activates worksheet by name |
| `navigate_home` | `gonagivetion`, `go_home` | `{ sheet?: string }` | Jumps to the "导航" navigation TOC sheet |
| `list_sheets` | - | `{}` | Returns list of all worksheets in workbook |
| `set_navigation` | `btnsetnavigation`, `setnagition` | `{ name?: string }` | Creates/updates "导航" TOC sheet with hyperlinks to all sheets |
| `open_hyperlink` | `btnopenhyperlink`, `openhyperlink` | `{ url?: string }` | Opens URL or follows selection's hyperlink / document reference |
| `select_range_by_address` | `selecterange_by_address`, `select_by_address` | `{ address: string }` | Navigates to address with sheet prefix (e.g. `"Sheet2!A1:B10"`) |

### C. Cell & Range Operations
| Canonical Command | Aliases | Params | Description |
| :--- | :--- | :--- | :--- |
| `set_cell_color` | `set_color`, `btnsetcolor` | `{ color: string, range?: string }` | Sets background fill color (yellow, hex, etc.) |
| `set_cell_value` | `set_value`, `btnsetvalue` | `{ value: any, range?: string }` | Writes single value or 2D array to range |
| `select_range` | - | `{ range: string, sheet?: string }` | Selects specified cell/range (e.g. `"A1:D10"`) |
| `set_border` | `setborder`, `btnsetborder` | `{ color?: string, range?: string }` | Applies continuous thick border to edges |
| `insert_row` | `insertrow` | `{ range?: string }` | Inserts row shifting down |
| `insert_column` | `insertcoloumn`, `insertcolumn` | `{ range?: string }` | Inserts column shifting right |
| `add_arrow_line` | `addarrowline`, `addarrow` | `{ startX?: number, startY?: number, endX?: number, endY?: number, color?: string }` | Draws connector arrow shape |
| `query_active_cell` | `setsharpposition`, `get_active_cell` | `{}` | Returns active cell address, local address, values, and pixel offsets |

### D. Position Slots & Fast Jump Navigation (Zero-Sheet Pollution)
Persisted using **Tier 2: Office Document Settings** (`Office.context.document.settings.set('excel_position_slots', ...)`).
Zero pollution of worksheet data, sub-millisecond in-memory lookups.

| Canonical Command | Aliases | Params | Description |
| :--- | :--- | :--- | :--- |
| `save_workbook_slot` | `recoderange`, `save_slot`, `save_position_slot` | `{ slot: string, range?: string }` | Saves active sheet & cell into global slot (e.g. `"1"`, `"a"`, `"u"`) |
| `switch_workbook_slot` | `readrecoderange`, `switch_slot`, `jump_slot` | `{ slot: string }` | Activates sheet and selects range for global slot |
| `save_sheet_slot` | `recodesheetrange` | `{ slot: string, range?: string }` | Saves cell position local to current worksheet |
| `switch_sheet_slot` | `readsheetrange` | `{ slot: string }` | Jumps to slot in current worksheet |
| `jump_previous_position` | `rangeprevios`, `jump_previous`, `history_back` | `{}` | Jumps to previous visited cell across workbook |
| `jump_previous_sheet_position` | `rangesheetprevios`, `jump_previous_sheet` | `{}` | Jumps to previous visited cell within current sheet |
| `reset_history_index` | `resetpreviosindex`, `rangepreviosindex` | `{}` | Resets navigation history pointer to latest |
| `clear_slots` | `resetrecodeinfomation` | `{ scope?: 'all' \| 'workbook' \| 'sheet' \| 'history' }` | Clears stored slots and/or history |
| `get_slots` | - | `{}` | Lists all saved workbook slots, sheet slots, and history depth |

### E. Data Import, Export & Calculations
| Canonical Command | Aliases | Params | Description |
| :--- | :--- | :--- | :--- |
| `export_selection` | `btnselectexport` | `{ range?: string }` | Reads values, address, rowCount, columnCount |
| `export_formula` | `btnselectexportformula` | `{ range?: string }` | Reads formulas of range |
| `export_merged_range` | `btnselectexportrangemerge` | `{ range?: string }` | Returns merge state of range |
| `import_range` | `btnselectimportrange` | `{ data: any[][], range?: string }` | Imports 2D array of values and autofits |
| `insert_sample_data` | `btninsertdata`, `insertdata` | `{}` | Creates `ExpensesTable` with sample records |
| `sum_selected_data` | `btnsumdata` | `{ outputCell?: string }` | Sums numeric values in selection and writes result |

### F. Diagnostic, Focus & Workbook Queries
| Canonical Command | Aliases | Params | Description |
| :--- | :--- | :--- | :--- |
| `query_workbook` | `loadfilename`, `get_file_info`, `get_workbook_info`, `query_focus` | `{}` | Returns active workbook fileName, URL, platform (`PC` vs `OfficeOnline`), isFocused, activeSheet |
| `ping` | - | `{}` | Returns `{ pong: true, timestamp }` |
| `connect_service` | `btnconnectservice` | `{}` | Verifies add-in service availability |

---

## 3. Focus Detection & Multi-Instance Routing (Desktop vs Web)

When multiple Excel instances exist (e.g. Excel Desktop, Excel Online in Chrome, or multiple workbooks open):

### A. Automatic Focus & Activity Detection
Each add-in instance tracks:
1. **DOM Focus**: `window.onfocus`, `window.onblur`, `document.hasFocus()`.
2. **Worksheet Grid Activity**: `context.workbook.onSelectionChanged` records timestamp whenever user interacts with cell grid.
3. If user clicked a cell in an Excel window within the last 60 seconds, `isFocused()` evaluates to `true`.

### B. Targeting Options in WebSocket Commands
You can filter commands from Kanata so only the intended Excel instance executes them:
1. **Execute only on the focused/active window**:
   ```json
   { "command": "set_cell_color", "params": { "color": "yellow", "onlyWhenFocused": true } }
   ```
2. **Target a specific workbook by filename**:
   ```json
   { "command": "create_sheet", "params": { "name": "Report", "workbook": "SalesData.xlsx" } }
   ```
3. **Target a specific platform (`desktop` vs `web`)**:
   ```json
   { "command": "select_range", "params": { "range": "A1", "platform": "desktop" } }
   ```

---

## 3. Kanata `.kbd` Configuration Example

In `c:\startup\kanata\kbd\kanata\excel.kbd`:

### A. Fast Position Slot Template (Single Tap = Switch / Double Tap = Save)
```lisp
;; Fast Excel Position Slot Template (matches chrome-ws-slot and cursor-ws-slot)
(deftemplate excel-ws-slot (slot)
  (defalias
    (concat "excel-save-slot-" $slot)
    (ws-send-channel "excel_websocket_Offic_Add_in"
      (concat r#"{"type":"command","id":"save-"# $slot r#"","command":"save_workbook_slot","params":{"slot":""# $slot r#""}}"#))

    (concat "excel-switch-slot-" $slot)
    (ws-send-channel "excel_websocket_Offic_Add_in"
      (concat r#"{"type":"command","id":"switch-"# $slot r#"","command":"switch_workbook_slot","params":{"slot":""# $slot r#""}}"#))

    ;; Tap = Switch, Double-tap = Save
    (concat "excel-slot-" $slot "-tap")
    (tap-dance 200 ((concat "@excel-switch-slot-" $slot) (concat "@excel-save-slot-" $slot)))
  )
)

;; Instantiate slots for fast navigation keys
(t! excel-ws-slot 1)
(t! excel-ws-slot 2)
(t! excel-ws-slot 3)
(t! excel-ws-slot u)
(t! excel-ws-slot i)
(t! excel-ws-slot o)
(t! excel-ws-slot j)
(t! excel-ws-slot k)
(t! excel-ws-slot l)
```

### B. Navigation & Cell Operations
```lisp
(defalias
  ;; Jump to previous visited position (Breadcrumb history back)
  excel-jump-back (ws-send-channel "excel_websocket_Offic_Add_in"
    r#"{"type":"command","id":"jump-back","command":"jump_previous_position","params":{}}"#)

  ;; Highlight active selection with yellow
  excel-highlight-yellow (ws-send-channel "excel_websocket_Offic_Add_in"
    r#"{"type":"command","id":"hl-yellow","command":"set_cell_color","params":{"color":"yellow"}}"#)

  ;; Toggle Taskpane
  excel-toggle-taskpane (ws-send-channel "excel_websocket_Offic_Add_in"
    r#"{"type":"command","id":"toggle-pane","command":"toggle_taskpane","params":{}}"#)

  ;; Create new Sheet (uses selected cell text if available)
  excel-create-sheet (ws-send-channel "excel_websocket_Offic_Add_in"
    r#"{"type":"command","id":"create-sheet","command":"create_sheet","params":{}}"#)

  ;; Generate Navigation Table of Contents ("导航")
  excel-set-navigation (ws-send-channel "excel_websocket_Offic_Add_in"
    r#"{"type":"command","id":"set-nav","command":"set_navigation","params":{}}"#)

  ;; Insert row down
  excel-insert-row (ws-send-channel "excel_websocket_Offic_Add_in"
    r#"{"type":"command","id":"ins-row","command":"insert_row","params":{}}"#)

  ;; Insert column right
  excel-insert-col (ws-send-channel "excel_websocket_Offic_Add_in"
    r#"{"type":"command","id":"ins-col","command":"insert_column","params":{}}"#)

  ;; Open Command Palette (Ctrl+Shift+P primary / Ctrl+Shift+K alias)
  excel-open-command-palette (ws-send-channel "excel_websocket_Offic_Add_in"
    r#"{"type":"command","id":"open-palette","command":"open_command_palette"}"#)
)

;; Overrides to intercept Ctrl+Shift+P and Ctrl+Shift+K directly in Excel before native dialogs
(defoverrides
  (lctl lsft p) @excel-open-command-palette
  (rctl rsft p) @excel-open-command-palette
  (lctl lsft k) @excel-open-command-palette
  (rctl rsft k) @excel-open-command-palette
)
```

