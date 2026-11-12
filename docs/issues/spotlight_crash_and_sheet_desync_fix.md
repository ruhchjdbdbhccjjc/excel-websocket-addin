# Issue Resolution: Spotlight Crash Residuals & Multi-Sheet Desync Fix

## 1. Executive Summary

During testing and adversarial code review of the WPS Spotlight engine and selection tracker, six interdependent reliability and safety defects were identified:
1. **User Conditional Format Deletion**: `f.includes('INDEX(')` falsely matched and deleted legitimate user lookup formulas.
2. **Sheet-Switching Ghost Crosshair Leak**: Switching from Sheet1 to Sheet2 rendered crosshairs on Sheet2 without purging Sheet1.
3. **Rapid Navigation Crosshair Freezing**: Fast arrow-key movement dropped intermediate selection events when `isUpdating === true`, leaving the crosshair stuck on an earlier cell.
4. **Crash Detox Bypass**: Persisted `enabled: true` in `localStorage` caused startup detox to be skipped on reboot after an Excel crash.
5. **Multi-Sheet Purge IPC Bottleneck**: Sequential per-sheet format inspection created $2N$ round-trips over the Office.js bridge.
6. **Hardcoded UI Strings in Rule Card**: `SpotlightRuleCard.tsx` did not connect to the i18n dictionary system.

---

## 2. Issue Logic Map (Failure Path vs. Success Path)

```text
===============================================================================
FAILURE PATH (BEFORE FIX):
===============================================================================

User Navigates Rapidly (A1 -> A2 -> A3)
       |
       v
[Selection: A1->A2 starts] ----> isUpdating = true (IPC sync running)
       |
[Selection: A2->A3 fires] -----> Check: if (isUpdating) return; ---> [A3 DROPPED!]
       |
       v
[A1->A2 sync completes] -------> isUpdating = false
                                 [RESULT: Spotlight STUCK on A2, cursor on A3]

User Switches Sheet1 -> Sheet2
       |
       v
[applySpotlight(Sheet2)] ------> Updates active sheet only
                                 [RESULT: Ghost crosshair frozen on Sheet1 forever]

Excel Crashes & Restarts
       |
       v
[loadPersistedSettings()] -----> Restores enabled: true
       |
[BackgroundService boot] ------> Check: if (!settings.enabled) purge();
                                 Condition is FALSE!
                                 [RESULT: Startup detox skipped; stale formats remain]

===============================================================================
SUCCESS PATH (AFTER FIX):
===============================================================================

User Navigates Rapidly (A1 -> A2 -> A3)
       |
       v
[Selection: A1->A2 starts] ----> isUpdating = true
       |
[Selection: A2->A3 fires] -----> Check: if (isUpdating) { pendingSelection = A3; return; }
       |
       v
[A1->A2 sync completes] -------> finally { isUpdating = false; }
                                 Check: if (pendingSelection) { consume and trigger A3! }
                                 [RESULT: Spotlight cleanly catches up to A3]

User Switches Sheet1 -> Sheet2
       |
       v
[applySpotlight(Sheet2, 'Sheet1')]
       |
       v
[getItemOrNullObject('Sheet1')] -> Purges Sheet1 spotlight items in same batch
[getActiveWorksheet() (Sheet2)] -> Applies fresh Spotlight on Sheet2
                                 [RESULT: Zero crosshair leak across worksheets]

Excel Crashes & Restarts
       |
       v
[BackgroundService boot] ------> Unconditionally runs purgeAllOrphanSpotlightsAcrossWorkbook()
                                 (Batched 3-roundtrip pipeline cleans entire workbook)
                                 If enabled: refreshes active cell on current sheet!
                                 [RESULT: Complete auto-detox; 0 residual rules]
```

---

## 3. Root Causes & Technical Solutions

### A. False-Positive Detection of User Formulas
- **Root Cause**: `getSpotlightItems` used `f.includes('INDEX(')`. Any user workbook containing `=INDEX($A$1:$A$100, MATCH(...))` was incorrectly flagged as an internal spotlight rule.
- **Solution**: Replaced with [`isSpotlightFormula()`](../src/websocket/service/ExcelSpotlightService.ts#L571-L593) which requires either spotlight coordinate anchors (`ROW()=`, `COLUMN()=`, `OR(ROW()=`, `AND(ROW()>=`) or the exact peer-match signature (`INDEX(1:1048576,ROW(),MATCH(`).

### B. Cross-Sheet Ghost Crosshairs
- **Root Cause**: [`applySpotlightForSelection()`](../src/websocket/service/ExcelSpotlightService.ts#L330-L373) only operated on `context.workbook.worksheets.getActiveWorksheet()`. When a user clicked a different sheet tab, the previous sheet's conditional format collection was untouched.
- **Solution**: Recorded `prevSheetName` in [`onSelectionChanged()`](../src/websocket/service/ExcelSpotlightService.ts#L225-L255) and added pre-flight cleanup of `context.workbook.worksheets.getItemOrNullObject(prevSheetName)`.

### C. Trailing Selection Drop During Fast Navigation
- **Root Cause**: Rapid arrow-key navigation fired `onSelectionChanged` while `isUpdating === true`, unconditionally dropping trailing selections without scheduling follow-up execution.
- **Solution**: Added `pendingSelection: SpotlightSelectionInfo | null` slot in [`ExcelSpotlightService`](../src/websocket/service/ExcelSpotlightService.ts#L88) and drained it in the `finally` block of [`applySpotlightForSelection()`](../src/websocket/service/ExcelSpotlightService.ts#L487-L495).

### D. Startup Detox Bypass
- **Root Cause**: In [`BackgroundService.ts`](../src/websocket/service/BackgroundService.ts#L105-L113), the detox call was wrapped in `if (!spotlight.getSettings().enabled)`. Because `localStorage` restored `enabled: true`, the detox was skipped after crashes.
- **Solution**: Made `purgeAllOrphanSpotlightsAcrossWorkbook()` unconditional on startup, followed by `refreshSpotlight()` if enabled.

---

## 4. Modified Files & Functions

- **[`ExcelSpotlightService.ts`](../src/websocket/service/ExcelSpotlightService.ts)**:
  - [`onSelectionChanged()`](../src/websocket/service/ExcelSpotlightService.ts#L206-L258): Added `pendingSelection` queue and `prevSheetName` tracking.
  - [`applySpotlightForSelection()`](../src/websocket/service/ExcelSpotlightService.ts#L330-L495): Added previous sheet purge and trailing execution slot drain.
  - [`isSpotlightFormula()`](../src/websocket/service/ExcelSpotlightService.ts#L571-L593): Safe formula discrimination pattern.
  - [`purgeAllOrphanSpotlightsAcrossWorkbook()`](../src/websocket/service/ExcelSpotlightService.ts#L535-L568): Batched 3-roundtrip workbook purge.
- **[`BackgroundService.ts`](../src/websocket/service/BackgroundService.ts)**:
  - [`initialize()`](../src/websocket/service/BackgroundService.ts#L105-L113): Unconditional startup detox.
- **[`SpotlightRuleCard.tsx`](../src/taskpane/components/spotlight/SpotlightRuleCard.tsx)**:
  - Integrated `useI18n()` for complete Chinese and English localization.
- **[`src/i18n/*`](../src/i18n)**:
  - Expanded `ruleCard` schema and locale entries.
