# WPS-Style Excel Spotlight with Anchor Cell & Column Matching Engine

## Overview

The WPS-Style Excel Spotlight (聚光灯) engine provides real-time cursor tracking and row/column crosshair highlighting in Excel via native Conditional Formatting. It features zero workbook pollution (preserves user cell fills and values, with zero undo-stack pollution) and introduces **Anchor Cell Key Column Highlighting** (同行重点列/首列锚点高亮) and **Dynamic Peer Matching** (同类值跨行联动).

When a user sets a target column (e.g. Column `A`, Column `1`, or header name `ID`), and activates any cell on that row (e.g. `B7`), the cell in that target column on the same row (`A7`) is automatically highlighted with a user-configured accent color.

---

## Function Logic Map (ASCII Art)

```
[ Active Cell Clicked: e.g. B7 (Row 7, Col 2) ]
                      |
                      v
      Workbook.onSelectionChanged Event
                      |
                      v
         [ ExcelSpotlightService ]
       applySpotlightForCell(row=6, col=1)
                      |
        +-------------+-------------+
        |                           |
        v                           v
  [ Extra Rules ]            [ Master Rule ]
  - Parse column targets     - Mode: crosshair / row / col
  - resolveColumnInfo()      - Color: classic pastel yellow
  - Formula:                 - Formula:
    =AND(ROW()=7,COLUMN()=1)   =OR(ROW()=7,COLUMN()=2)
        |                           |
        +-------------+-------------+
                      |
                      v
      [ Desired Rules Assembly (Priority Sorted) ]
       Index 0: Specific Anchor / Peer Rules (HIGH PRIORITY)
       Index 1: Master Crosshair / Spotlight (BASE PRIORITY)
                      |
                      v
      [ In-Place Safe Rule Reuse & Isolation ]
       getSpotlightItems() -> Filter custom formulas
       Delete surplus -> Update existing -> Add new
                      |
                      v
          [ Native Excel Rendering ]
   Cell A7 (Priority 0) ----------> User Accent Fill / 4-Sided Box
   Cells B7..Z7 (Priority 1) -----> Soft Pastel Spotlight Fill
   Cells A1..A6 (No Match) -------> Original Workbook Cell Format
```

---

## Key Features

1. **Flexible Column Targeting**:
   - **Column Letter**: Standard Excel column letters (e.g., `A`, `B`, `AA`). Formula: `COLUMN()=COLUMN($A$1)`.
   - **Column Number**: 1-based numeric index (e.g., `1`, `27`). Formula: `COLUMN()=1` (clamped between `1` and `16384`).
   - **Header Title**: Name match in row 1 (e.g., `#ID`, `姓名`). Formula: `COLUMN()=IFERROR(MATCH("ID",$1:$1,0),0)`.
   - **Dynamic Active Column**: `@` tracks current cursor column dynamically.

2. **Dual Targeting Modes**:
   - **Anchor Key Cell (`key_column`)**: Highlights exclusively the target column's cell on the active row (e.g. `A7` when on `B7`), with an optional 4-sided enclosing border box.
   - **Peer Value Matching (`peer_match`)**: Highlights all other rows across the sheet sharing the same key value, while also highlighting the active row's source cell (`A7`) in that color.

3. **Safe Rule Isolation & Clean Removal**:
   - Only manages formulas created by the spotlight engine (`=ROW()=`, `=COLUMN()=`, `=OR(ROW()=`, `=AND(ROW()=`, `=INDEX(`).
   - Preserves all pre-existing user conditional formatting rules and cell backgrounds.
   - 0 Undo-stack disruption (clean formula manipulation).

---

## Code References

- **[ExcelSpotlightService](../src/websocket/service/ExcelSpotlightService.ts)**: Core spotlight engine singleton.
  - [`getInstance`](../src/websocket/service/ExcelSpotlightService.ts#L66) - Singleton access.
  - [`applySpotlightForCell`](../src/websocket/service/ExcelSpotlightService.ts#L168) - Builds prioritized rule collection and syncs conditional formats.
  - [`resolveColumnInfo`](../src/websocket/service/ExcelSpotlightService.ts#L398) - Parses column input into formula expressions with bounds clamping.
  - [`getSpotlightItems`](../src/websocket/service/ExcelSpotlightService.ts#L305) - Isolates spotlight conditional format items.
  - [`clearActiveSpotlight`](../src/websocket/service/ExcelSpotlightService.ts#L273) - Removes all active spotlight rules cleanly.
- **[SpotlightTab](../src/taskpane/components/spotlight/SpotlightTab.tsx#L13)**: SolidJS Taskpane panel for spotlight master toggle and rule builder.
- **[SpotlightRuleCard](../src/taskpane/components/spotlight/SpotlightRuleCard.tsx#L11)**: Rule card with type selector, column input, scope input, and color pills.
- **[ColorPills](../src/taskpane/components/spotlight/ColorPills.tsx#L16)**: Pastel color picker component supporting WPS iconic colors.
