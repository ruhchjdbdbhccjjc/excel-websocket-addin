# Excel Spotlight Adaptive Selection Engine (形态自适应多选引擎)

## 1. Overview & Problem Definition

In spreadsheets, user selections vary widely between single-cell navigation, multi-cell rectangular ranges, non-contiguous multi-selections (`Ctrl+Click`), whole-row selections (`5:5`), whole-column selections (`C:C`), and whole-sheet selections (`Ctrl+A`).

Prior to this implementation:
1. Multi-cell selections (`B3:D8`) only read the top-left cell (`B3`), ignoring the rest of the selected area.
2. Clicking row headers (`5:5`) erroneously highlighted Column A.
3. Clicking column headers (`C:C`) erroneously highlighted Row 1.
4. Non-contiguous selections (`Ctrl+Click`) caused `getSelectedRange()` to throw `InvalidSelection`, crashing passive tracking.
5. Whole-sheet selection (`Ctrl+A`) painted the entire spreadsheet in spotlight yellow.

The **Adaptive Selection Engine** resolves all these problems by coupling **ActiveCell Anchor Tracking** with **Context-Aware Range Detection** and a customizable **Multi-Cell Selection Behavior** switch.

---

## 2. Selection Behavior Modes (多选选区行为)

Users can configure the multi-cell behavior directly in the Taskpane panel or via remote WebSocket commands:

| Mode | Identifier | Visual Behavior | Primary Use Case |
| :--- | :--- | :--- | :--- |
| **🎯 活动焦点 (Active Cell)** | `'active_cell'` | The spotlight crosshair tracks the white active cell (`ActiveCell`) within the selection. | Daily data entry & review; aligns with WPS Office native behavior. |
| **🔲 选区包围 (Banded Area)** | `'banded_area'` | The crosshair dynamically expands to span the entire selection width (`startCol..endCol`) and height (`startRow..endRow`). | Range verification, presentation demos, and block data auditing. |
| **⏸️ 自动避让 (Auto-Hide)** | `'auto_hide'` | Spotlight automatically clears/suspends when `cellCount > 1`, and seamlessly re-appears when a single cell is clicked. | Heavy formula editing and copying/pasting large datasets without visual distractions. |

---

## 3. Architecture & Logic Flow Map (ASCII Art)

```
[ User Selection Event in Excel ]
                 |
                 v
   context.workbook.onSelectionChanged
                 |
                 +------------------------------------------------------+
                 | 1. Always load ActiveCell (never throws InvalidSelection)
                 | 2. Try loading SelectedRange (rowCount, columnCount) |
                 +------------------------------------------------------+
                 |
        +--------+--------+
        |                 |
 [ Success ]      [ Error: InvalidSelection (Ctrl+Click) ]
        |                 |
        |                 v
        |         Fallback to ActiveCell (single cell 1x1)
        v
 [ Build SpotlightSelectionInfo ]
  - activeRow, activeCol
  - startRow, endRow, startCol, endCol
  - rowCount, columnCount
  - isEntireRow (columnCount >= 16384)
  - isEntireColumn (rowCount >= 1048576)
  - isEntireSheet (1048576 x 16384)
  - isMultiCell (rowCount > 1 || columnCount > 1)
                 |
                 v
     [ ExcelSpotlightService.applySpotlightForSelection ]
                 |
                 +---> [ isEntireSheet? ] ---------------------> YES ---> Clear spotlight (Suspend)
                 |
                 +---> [ auto_hide && isMultiCell? ] ---------> YES ---> Clear spotlight (Temporary hide)
                 |
                 +---> [ isEntireRow? ] -----------------------> YES ---> Downgrade to mode: 'row_only' (No Col A!)
                 |
                 +---> [ isEntireColumn? ] --------------------> YES ---> Downgrade to mode: 'col_only' (No Row 1!)
                 |
                 v
  [ Assemble Conditional Formatting Formulas ]
  - If banded_area: =OR(AND(ROW()>=r1,ROW()<=r2), AND(COLUMN()>=c1,COLUMN()<=c2))
  - If active_cell: =OR(ROW()=excelRow, COLUMN()=excelCol)
  - Anchor Rules:   =AND(ROW_PART, COLUMN_PART)
                 |
                 v
  [ Safe In-Place Rule Update & Cleanup ]
```

---

## 4. Edge-Case Matrix

| Scenario | Raw Selection | Engine Handling | Visual Output |
| :--- | :--- | :--- | :--- |
| **Single Cell** | `B7` | Standard single-cell crosshair. | Row 7 and Col B highlighted. |
| **Row Header Click** | `5:5` | `isEntireRow: true` $\rightarrow$ forced `row_only`. | Row 5 highlighted; Col A is NOT highlighted. |
| **Column Header Click** | `C:C` | `isEntireColumn: true` $\rightarrow$ forced `col_only`. | Col C highlighted; Row 1 is NOT highlighted. |
| **Whole Sheet (`Ctrl+A`)** | `1048576 x 16384` | `isEntireSheet: true` $\rightarrow$ auto suspended. | Spotlight cleanly removed; 0 yellow glare. |
| **Non-contiguous (`Ctrl+Click`)** | `A1, C5, E9` | `getSelectedRange()` throws $\rightarrow$ graceful fallback to `getActiveCell()`. | Pinned cleanly to the currently active cell. |
| **Multi-cell Block (`B3:D8`)** | `6 rows x 3 cols` | Follows `multiSelectBehavior`: `active_cell` vs `banded_area` vs `auto_hide`. | Pinned to active cell, banded over rows 3-8 & cols B-D, or hidden. |

---

## 5. Remote WebSocket Protocol API

External controllers (such as Kanata) can query or modify the multi-selection behavior:

### Setting Multi-Select Behavior
```json
{
  "action": "set_spotlight",
  "multiSelectBehavior": "banded_area"
}
```

Supported values for `multiSelectBehavior`:
- `"active_cell"`
- `"banded_area"`
- `"auto_hide"`

### Querying Status
```json
{
  "action": "get_spotlight_status"
}
```
Response:
```json
{
  "ok": true,
  "data": {
    "enabled": true,
    "mode": "both",
    "style": "fill",
    "color": "#FEF08A",
    "multiSelectBehavior": "banded_area",
    "rules": []
  }
}
```
