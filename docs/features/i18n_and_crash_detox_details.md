# Split-File i18n & Whole-Workbook Crash Detox Architecture

## 1. Overview

This document details two major enterprise capabilities implemented in the Excel Flow Add-in:
1. **Split-File Reactive i18n Architecture**: Multi-locale dictionary system separating TypeScript type contracts from individual locale files (`zh-CN` and `en-US`), featuring reactive language switching via SolidJS signals and automatic host environment language detection.
2. **Whole-Workbook Crash Detox & Lifecycle Purge**: Bulletproof conditional formatting cleanup preventing workbook pollution when Excel is abruptly terminated, closed without saving, or restarted after a crash.

---

## 2. Split-File i18n Architecture

### Key Components

- **Type Contract**: [`TranslationSchema`](../src/i18n/types.ts#L3-L65) defines the strongly typed tree for all UI labels, tooltips, placeholders, and feedback messages.
- **Locale Dictionaries**:
  - Simplified Chinese: [`zhCN`](../src/i18n/locales/zh-CN.ts#L3-L65)
  - American English: [`enUS`](../src/i18n/locales/en-US.ts#L3-L65)
- **Reactive i18n Hook**: [`useI18n()`](../src/i18n/index.ts#L48-L68) providing reactive `t()`, `locale()`, and `setLocale()`.

### Detection & Resolution Hierarchy

```text
+-------------------------------------------------------------+
|             i18n INITIAL LOCALE RESOLUTION HIERARCHY        |
+-------------------------------------------------------------+
                              |
                              v
                 [Check localStorage cache]
                 ('excel_addin_locale')
                   /                \
             (Found)                (None)
               /                        \
              v                          v
      [Apply User Locale]       [Query Office.context]
      ('zh-CN' / 'en-US')     (.displayLanguage.toLowerCase())
                                 /              \
                           (StartsWith 'zh')   (Other)
                               /                  \
                              v                    v
                       [Default: 'zh-CN']    [Query navigator.language]
                                                /              \
                                          (StartsWith 'zh')   (Other)
                                              /                  \
                                             v                    v
                                      [Default: 'zh-CN']   [Fallback: 'en-US']
```

---

## 3. Whole-Workbook Crash Detox & Purge Engine

### The Workbook Pollution Challenge

In Microsoft Excel, conditional formatting rules applied to worksheets are persisted to the `.xlsx` document structure upon save or auto-save. If Excel crashes, closes abruptly, or unloads the add-in runtime while the WPS Spotlight crosshair is active, conditional formatting formulas (`=ROW()=...`, `=COLUMN()=...`) remain embedded in the worksheet cells.

### Detox Execution Flow (Function Logic Map)

```text
+-----------------------------------------------------------------------------+
|               WHOLE-WORKBOOK CRASH DETOX & DETOX EXECUTION FLOW              |
+-----------------------------------------------------------------------------+

  [Add-in Boot (Office.onReady)]
                 |
                 v
  [BackgroundService.initialize()] ----------------------------+
                 |                                             |
                 v                                             v
  [purgeAllOrphanSpotlightsAcrossWorkbook()]          [Window Event Hooks]
                 |                                     (beforeunload / pagehide)
                 v                                             |
  +----------------------------------------------------+       v
  |  Phase 1: Query All Sheets in Workbook             |  [clearActiveSpotlight()]
  |  context.workbook.worksheets.load('items/name')    |
  |  (1st IPC sync round-trip)                         |
  +----------------------------------------------------+
                 |
                 v
  +----------------------------------------------------+
  |  Phase 2: Batched Conditional Format Collection    |
  |  For each sheet: sheet.getRange().conditionalFormats|
  |  cfs.load('items/type')                            |
  |  (2nd IPC sync round-trip - Batched across sheets) |
  +----------------------------------------------------+
                 |
                 v
  +----------------------------------------------------+
  |  Phase 3: Inspect Custom Rule Formulas             |
  |  For custom formats: cf.custom.rule.load('formula')|
  |  (3rd IPC sync round-trip - Parallelized)          |
  +----------------------------------------------------+
                 |
                 v
  +----------------------------------------------------+
  |  Phase 4: Strict Signature Discrimination          |
  |  Execute isSpotlightFormula(f):                    |
  |  - Requires ROW()=/COLUMN()= OR                    |
  |  - Requires INDEX(1:1048576,ROW(),MATCH(...)       |
  |  - Preserves user custom INDEX/MATCH formulas      |
  +----------------------------------------------------+
                 |
                 v
  +----------------------------------------------------+
  |  Phase 5: Queue Deletions & Sync                   |
  |  surplusItems[i].delete()                          |
  |  (Final sync round-trip)                           |
  +----------------------------------------------------+
                 |
                 v
  [Check if settings.enabled == true]
     /                           \
  (Yes)                          (No)
    /                              \
   v                                v
[spotlight.refreshSpotlight()]    [Detox Complete: 0 Residuals]
(Re-apply clean crosshair to
 active cell on current sheet)
```

---

## 4. Key Functions & Code References

- [`useI18n()`](../src/i18n/index.ts#L48-L68): Reactive translation getter and locale setter.
- [`purgeAllOrphanSpotlightsAcrossWorkbook()`](../src/websocket/service/ExcelSpotlightService.ts#L535-L568): 3-phase atomic batch purge across all sheets.
- [`isSpotlightFormula()`](../src/websocket/service/ExcelSpotlightService.ts#L571-L593): Zero false-positive formula discrimination.
- [`applySpotlightForSelection()`](../src/websocket/service/ExcelSpotlightService.ts#L330-L373): Cross-sheet previous worksheet ghost cleanup.
- [`initialize()`](../src/websocket/service/BackgroundService.ts#L58-L114): Startup auto-detox orchestration in headless runtime.
