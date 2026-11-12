# Position Slots & Border Color Cycle Feature Details

This document provides technical details for the **Zero-Pollution Position Slots** and **Border Color Cycling with Non-Intrusive Floating Shape Preview Badge** in the Excel Office Add-in.

---

## 1. Overview & Objectives

1. **Zero-Pollution Position Slots**:
   - Allows saving and jumping between cell coordinates across sheets and workbooks.
   - Dual-scope support:
     - **Global Workbook Slots**: Saved in `Office.context.document.settings` under key `__KANATA_POS_SLOTS__`.
     - **Worksheet-Local Slots**: Saved per sheet name under key `__KANATA_SHEET_SLOTS__<SheetName>`.
   - Never modifies worksheet cells, formulas, or cell formats.
   - Synced across Excel Desktop (Win32) and Excel Web.

2. **Border Color Cycling & Floating Shape Badge**:
   - Cycles borders across 6 curated colors:
     `#E26D5C` (Coral Red) -> `#C7CC7A` (Sage Green) -> `#7560BA` (Royal Purple) -> `#9DD9D2` (Teal) -> `#FFE1A8` (Warm Peach) -> `#FFFFFF` (Clear / Style None).
   - Solves Excel's green selection marquee masking the active cell edges by rendering an in-memory 22x22 floating rounded rectangle badge next to the cell for 1.5 seconds, then automatically deleting it.

3. **Window Focus Auto-Filtering**:
   - Commands default to `onlyWhenFocused: true` for mutating operations, ensuring Kanata broadcast messages only affect the active Excel window.

---

## 2. Configuration & Integration Links

- **Kanata Excel Configuration**: [excel.kbd](file:///c:/startup/kanata/kbd/kanata/excel.kbd)
- **Kanata Config File Mirror**: [kanata_config_file/excel.kbd](file:///c:/startup/kanata/kbd/kanata_config_file/excel.kbd)
- **Kanata Layer Switcher**: [layer.kbd](file:///c:/startup/kanata/kbd/kanata/layer.kbd)
- **Kanata App Layer SmartMap**: [defsmartmap.kbd](file:///c:/startup/kanata/kbd/kanata/defsmartmap.kbd)

---

## 3. Function Logic Map (ASCII Art)

```
===================================================================================
                   KANATA TO EXCEL POSITION SLOTS & BORDER DISPATCH
===================================================================================

 [Kanata Engine]
        |
        | ws-send-channel "excel_websocket_Offic_Add_in"
        v
 [WebSocketClient.ts]
        |
        | onmessage -> KanataProtocol.parseMessage()
        v
 [MessageHandler.ts] -> handleMessage()
        |
        v
 [ExcelCommandExecutor.ts] -> execute(rawCommandName, params)
        |
        +---> Focus Check: slotManager.isFocused()?
        |       |-- [NO]  --> Skip mutating action (return ok: false, error: 'excel_not_focused')
        |       +-- [YES] --> Continue to Handler
        |
        +---> Case 1: save_slot / switch_slot / clear_slot
        |       |
        |       v
        |     [ExcelPositionSlotManager.ts]
        |       |
        |       +---> saveSlot()   ---> Office.context.document.settings.set(...)
        |       |                        Office.context.document.settings.saveAsync()
        |       +---> switchSlot() ---> sheet.activate() -> range.select()
        |
        +---> Case 2: set_border (setboard)
                |
                v
              [ExcelCommandExecutor.ts] -> set_border
                |
                +---> Read current EdgeBottom border color & style
                +---> Select next color from 6-color palette
                +---> Apply Thick Continuous border to all 4 edges (or None if Clear)
                +---> Compute active cell (x, y) coordinates via bounding range
                +---> shapes.addGeometricShape(__border_preview_badge__)
                +---> Fill badge with selectedColor (Thick contrast border)
                +---> setTimeout(1500ms) -> Delete shape (__border_preview_badge__)
```

---

## 4. Key Functions & Code References

- **[ExcelCommandExecutor.execute](../src/websocket/service/ExcelCommandExecutor.ts#L48-L113)**: Command dispatch with automatic focus enforcement.
- **[ExcelCommandExecutor.set_border](../src/websocket/service/ExcelCommandExecutor.ts#L362-L483)**: 6-color cycle palette calculation, border styling, and floating badge generation.
- **[ExcelPositionSlotManager.saveSlot](../src/websocket/service/ExcelPositionSlotManager.ts#L107-L149)**: Stores coordinates into `Office.context.document.settings`.
- **[ExcelPositionSlotManager.switchSlot](../src/websocket/service/ExcelPositionSlotManager.ts#L151-L200)**: Restores worksheet selection and activates sheet without cell mutations.
- **[ExcelPositionSlotManager.isFocused](../src/websocket/service/ExcelPositionSlotManager.ts#L320-L335)**: Determines if the Excel document is active.

---

## 5. Kanata Layer Mapping

The dedicated Kanata layer `excel-fast-range-switch` binds keys matching Chrome window switching:

| Key | Slot | Tap Action | Double Tap Action | Shift + Tap (Local Sheet) |
| :--- | :--- | :--- | :--- | :--- |
| `u` | `u` | Switch to Global Slot `u` | Save to Global Slot `u` | Switch / Save Local Sheet Slot `u` |
| `i` | `i` | Switch to Global Slot `i` | Save to Global Slot `i` | Switch / Save Local Sheet Slot `i` |
| `o` | `o` | Switch to Global Slot `o` | Save to Global Slot `o` | Switch / Save Local Sheet Slot `o` |
| `j` | `j` | Switch to Global Slot `j` | Save to Global Slot `j` | Switch / Save Local Sheet Slot `j` |
| `k` | `k` | Switch to Global Slot `k` | Save to Global Slot `k` | Switch / Save Local Sheet Slot `k` |
| `l` | `l` | Switch to Global Slot `l` | Save to Global Slot `l` | Switch / Save Local Sheet Slot `l` |
| `m` | `m` | Switch to Global Slot `m` | Save to Global Slot `m` | Switch / Save Local Sheet Slot `m` |
| `,` | `comma` | Switch to Global Slot `,` | Save to Global Slot `,` | Switch / Save Local Sheet Slot `,` |
| `.` | `period` | Switch to Global Slot `.` | Save to Global Slot `.` | Switch / Save Local Sheet Slot `.` |
