# xlsx-ts

[中文 README](README.md)

A prototype XLSX reader/writer built around a "lossless first" principle.

The goal is not to map the entire Excel object model into a huge JS object graph first.
The first baseline is simpler and stricter:

`read(xlsx) -> write(xlsx)`

After that roundtrip, the extracted package parts should stay byte-for-byte identical unless a part was intentionally edited.

Once that baseline holds, styles, themes, comments, relationship files, and unknown extension nodes are preserved naturally.
Then higher-level APIs can be added on top with much lower risk.

## Design

The library is split into two layers:

1. `Lossless package layer`
   - Treat an `.xlsx` file as a zip package.
   - Keep every entry as raw bytes first.
   - Write untouched entries back exactly as they were, without re-serialization.

2. `Editable workbook layer`
   - Apply targeted XML patches only to the parts that actually need to change.
   - The current prototype supports:
     - reading sheet lists
     - reading cells
     - writing cells
     - reading formulas
     - writing formulas
   - Style-related `s="..."` attributes are preserved, so styles are not lost when values change.

## Why This Works For Style Preservation

Most style loss is not caused by failing to parse `styles.xml`.
It usually happens because the workbook is regenerated wholesale on write, which tends to break things like:

- unknown nodes
- attribute ordering
- namespaces and extension markers
- relationship file ordering
- coupling between shared strings, worksheets, and styles

The lossless-first direction flips that approach:

- preserve every package part first
- edit only the parts that must change
- write untouched parts back exactly as-is

That makes it much easier to satisfy a strict "roundtrip without diffs" requirement.

## Current API

- `Workbook.open(path)`
- `Workbook.fromEntries(entries)`
- `workbook.listEntries()`
- `workbook.getSheets()`
- `workbook.getSheet(name)`
- `workbook.getActiveSheet()`
- `workbook.getSheetVisibility(name)`
- `workbook.getDefinedNames()`
- `workbook.getDefinedName(name, scope?)`
- `workbook.setDefinedName(name, value, options?)`
- `workbook.deleteDefinedName(name, scope?)`
- `workbook.renameSheet(currentName, nextName)`
- `workbook.moveSheet(name, targetIndex)`
- `workbook.addSheet(name)`
- `workbook.deleteSheet(name)`
- `workbook.setSheetVisibility(name, visibility)`
- `workbook.setActiveSheet(name)`
- `sheet.cell(address)`
- `sheet.cell(rowNumber, column)`
- `sheet.rename(name)`
- `sheet.getCell(address)`
- `sheet.getCell(rowNumber, column)`
- `sheet.getStyleId(address)`
- `sheet.getStyleId(rowNumber, column)`
- `sheet.getCellEntries()`
- `sheet.iterCellEntries()`
- `sheet.rowCount`
- `sheet.columnCount`
- `sheet.getHeaders(headerRowNumber?)`
- `sheet.getRecord(rowNumber, headerRowNumber?)`
- `sheet.getRecords(headerRowNumber?)`
- `sheet.getColumn(column)`
- `sheet.getColumnEntries(column)`
- `sheet.getRow(rowNumber)`
- `sheet.getRowEntries(rowNumber)`
- `sheet.getRange(range)`
- `sheet.getUsedRange()`
- `sheet.getMergedRanges()`
- `sheet.getAutoFilter()`
- `sheet.getFreezePane()`
- `sheet.getSelection()`
- `sheet.getDataValidations()`
- `sheet.getTables()`
- `sheet.getHyperlinks()`
- `sheet.addTable(range, options?)`
- `sheet.removeTable(name)`
- `sheet.setHyperlink(address, target, options?)`
- `sheet.removeHyperlink(address)`
- `sheet.setAutoFilter(range)`
- `sheet.freezePane(columnCount, rowCount?)`
- `sheet.unfreezePane()`
- `sheet.setSelection(activeCell, range?)`
- `sheet.removeAutoFilter()`
- `sheet.setDataValidation(range, options?)`
- `sheet.removeDataValidation(range)`
- `sheet.setCell(address, value)`
- `sheet.setCell(rowNumber, column, value)`
- `sheet.setStyleId(address, styleId)`
- `sheet.setStyleId(rowNumber, column, styleId)`
- `sheet.deleteCell(address)`
- `sheet.deleteCell(rowNumber, column)`
- `sheet.deleteRow(row, count?)`
- `sheet.deleteColumn(column, count?)`
- `sheet.insertRow(row, count?)`
- `sheet.insertColumn(column, count?)`
- `sheet.setHeaders(headers, headerRowNumber?, startColumn?)`
- `sheet.setRecord(rowNumber, record, headerRowNumber?)`
- `sheet.setRecords(records, headerRowNumber?)`
- `sheet.deleteRecord(rowNumber, headerRowNumber?)`
- `sheet.deleteRecords(rowNumbers, headerRowNumber?)`
- `sheet.addRecord(record, headerRowNumber?)`
- `sheet.addRecords(records, headerRowNumber?)`
- `sheet.appendRow(values, startColumn?)`
- `sheet.appendRows(rows, startColumn?)`
- `sheet.setColumn(column, values, startRow?)`
- `sheet.setRow(rowNumber, values, startColumn?)`
- `sheet.setRange(startAddress, values)`
- `sheet.addMergedRange(range)`
- `sheet.removeMergedRange(range)`
- `sheet.getFormula(address)`
- `sheet.getFormula(rowNumber, column)`
- `sheet.setFormula(address, formula, options?)`
- `sheet.setFormula(rowNumber, column, formula, options?)`
- `workbook.save(path)`

Example:

```ts
const workbook = await Workbook.open("input.xlsx");
const sheet = workbook.getSheet("Sheet1");
const scoreCell = sheet.cell("B2");
const scoreValue = sheet.getCell(2, 2);
const scoreStyleId = sheet.getStyleId(2, 2);
const detailSheet = workbook.addSheet("Detail");
const activeSheet = workbook.getActiveSheet();

workbook.setDefinedName("Scores", "Summary!$A$1:$B$10");
workbook.setDefinedName("LocalScore", "$B$2", { scope: "Summary" });
workbook.renameSheet("Sheet1", "Summary");
workbook.moveSheet("Summary", 0);
workbook.setActiveSheet("Summary");
workbook.setSheetVisibility("Summary", "hidden");
detailSheet.rename("Detail 2026");
console.log(sheet.getTables());
console.log(sheet.getHyperlinks());
console.log(sheet.rowCount, sheet.columnCount);
console.log(sheet.getFreezePane(), sheet.getSelection(), activeSheet.name);
sheet.addTable("A1:B10", { name: "Scores" });
sheet.setHyperlink("A1", "https://example.com", { text: "Hello", tooltip: "Open link" });
sheet.setHyperlink("B2", "#Summary!A1");
sheet.setAutoFilter("A1:F20");
sheet.freezePane(1, 1);
sheet.setSelection("B2", "B2:C4");
sheet.setDataValidation("B2:B100", { type: "whole", operator: "between", formula1: "0", formula2: "100" });
sheet.setCell(3, 2, 98);
sheet.setStyleId(3, 2, scoreStyleId);
sheet.setCell("A1", "Hello");
sheet.deleteRow(8);
sheet.deleteColumn("G");
sheet.insertRow(2);
sheet.setHeaders(["Name", "Score"]);
sheet.insertColumn("B");
sheet.setRecord(2, { Name: "Alice", Score: 98 });
sheet.setRecords([
  { Name: "Alice", Score: 98 },
  { Name: "Bob", Score: 87 },
]);
sheet.deleteRecord(4);
sheet.deleteRecords([6, 7]);
sheet.addRecord({ Name: "Alice", Score: 98 });
sheet.addRecords([
  { Name: "Bob", Score: 87 },
  { Name: "Cara", Score: 91 },
]);
sheet.appendRow(["tail", 1]);
sheet.appendRows([
  ["tail-2", 2],
  ["tail-3", 3],
]);
sheet.setColumn("F", ["Q1", "Q2"], 2);
sheet.setRow(5, ["Name", "Score"], 2);
sheet.setRange("B2", [
  [1, 2],
  [3, 4],
]);
sheet.addMergedRange("D1:E1");
sheet.setFormula("B1", "SUM(1,2)", { cachedValue: 3 });
sheet.setFormula(4, 3, "SUM(A4:B4)", { cachedValue: 12 });
sheet.removeHyperlink("B2");
sheet.unfreezePane();
sheet.removeAutoFilter();
sheet.removeDataValidation("B2:B100");
sheet.removeTable("Scores");
detailSheet.setCell("A1", "created");
workbook.setSheetVisibility("Summary", "visible");
console.log(workbook.getDefinedNames(), workbook.getDefinedName("LocalScore", "Summary"));
workbook.deleteDefinedName("LocalScore", "Summary");
workbook.deleteSheet("Temp");
console.log(scoreCell.value, scoreCell.styleId, scoreCell.formula);

await workbook.save("output.xlsx");
```

Notes:

- On first read/write access, a sheet scans `sheetData` once and builds indexes for rows and cells.
- `sheet.cell(address)` returns a reusable `Cell` handle whose parsed value/formula/style state is cached by sheet revision.
- `sheet.cell()`, `getCell()`, `setCell()`, `getFormula()`, and `setFormula()` now support both `A1` addresses and `(rowNumber, column)` calls. Row and column indexes are 1-based.
- Later `getCell()` and `getFormula()` calls use those indexes directly instead of running a full string match on every read.
- `sheet.rowCount` and `sheet.columnCount` currently mean the maximum used row number and maximum used column number. Empty sheets return `0`.
- `sheet.getCellEntries()`, `iterCellEntries()`, `getRowEntries()`, and `getColumnEntries()` expose the real worksheet `<c>` nodes with address, row/column indexes, type, style id, and value, which is useful for large or sparse sheet iteration.
- `sheet.deleteCell()` removes the worksheet `<c>` node entirely; if you want to keep a styled placeholder but clear the value, continue using `setCell(..., null)`.
- `sheet.getStyleId()` and `setStyleId()` currently read and write the cell-level `s="..."` style index; both `A1` and `(rowNumber, column)` calls are supported, but `styles.xml` is not edited directly yet.
- `sheet.getFreezePane()`, `freezePane()`, and `unfreezePane()` currently manage worksheet `sheetViews/sheetView/pane`; `topLeftCell` keeps tracking row and column insert/delete operations.
- `sheet.getSelection()` and `setSelection()` currently read and write worksheet `sheetViews/sheetView/selection`; when a frozen pane exists, they target the selection for the current active pane.
- After each write, the sheet index is rebuilt so later reads always see the latest content.
- Worksheet edits keep `<dimension ref="...">` in sync so used-range metadata does not go stale.
- `deleteRow()` and `deleteColumn()` currently update cell coordinates, formulas, merged ranges, worksheet `dimension`, common `ref` and `sqref` attributes, `definedNames`, and explicit formulas in other sheets that reference the edited sheet.
- `insertRow()` currently updates cell coordinates, formulas, merged ranges, worksheet `dimension`, common `ref` and `sqref` attributes, `definedNames`, and explicit formulas in other sheets that reference the edited sheet.
- `insertColumn()` currently updates cell coordinates, formulas, merged ranges, worksheet `dimension`, common `ref` and `sqref` attributes, `definedNames`, and explicit formulas in other sheets that reference the edited sheet.
- `sheet.getTables()` currently reads existing table names, display names, ranges, and part paths.
- `sheet.getHyperlinks()` currently reads internal and external hyperlinks from the sheet; external link targets are resolved through the sheet relationships part.
- `sheet.getAutoFilter()`, `sheet.setAutoFilter()`, and `sheet.removeAutoFilter()` currently manage the worksheet-level `autoFilter`; removing it also clears the top-level `sortState`.
- `sheet.getDataValidations()`, `sheet.setDataValidation()`, and `sheet.removeDataValidation()` currently manage worksheet-level `dataValidations`, including common attributes plus `formula1` and `formula2`, and keep `sqref` updated during row and column edits.
- `sheet.addTable()` currently creates the basic table part, sheet relationship, `[Content_Types].xml` override, and table XML. Column names default to the first row in the range, and blank names fall back to `ColumnN`.
- `sheet.removeTable()` currently removes the current sheet's `tableParts`, sheet relationship, table XML, and matching content type override.
- Existing linked tables keep their own `ref` and `autoFilter` updated during row and column insert/delete operations. If a table becomes empty, its `tableParts` entry is removed from the sheet.
- `sheet.setHyperlink()` and `sheet.removeHyperlink()` currently manage worksheet `<hyperlinks>` plus the matching sheet relationship for external links. Internal targets use a format like `#Sheet1!A1`.
- `workbook.getDefinedNames()`, `getDefinedName()`, `setDefinedName()`, and `deleteDefinedName()` currently support both global and local defined names.
- `workbook.getSheetVisibility()` and `setSheetVisibility()` currently support `visible`, `hidden`, and `veryHidden`, and prevent hiding the last visible sheet in the workbook.
- `workbook.getActiveSheet()` and `setActiveSheet()` currently read and write `workbookView.activeTab`; if the workbook does not yet contain `bookViews`, they are created automatically, and hidden sheets cannot be activated.
- `workbook.renameSheet()` and `sheet.rename()` currently update sheet names, explicit formula references in other sheets, `definedNames`, internal hyperlink locations, and document properties.
- `workbook.moveSheet()` currently uses a 0-based `targetIndex` and keeps workbook `<sheets>` order, worksheet order in `docProps/app.xml`, local defined-name `localSheetId` values, and `workbookView.activeTab` aligned.
- `workbook.addSheet()` and `workbook.deleteSheet()` currently maintain `workbook.xml`, workbook rels, and `[Content_Types].xml`, and adjust remaining formulas and `definedNames` when a sheet is deleted.

## Benchmarking

The repo now includes a sanitized large benchmark workbook at [res/monster.xlsx](/Users/codetypes/Desktop/Github/xlsx-ts/res/monster.xlsx), intended for repeatable performance regression checks.

Common commands:

- `npm run bench:monster`
  - Run a 3-iteration comparison on `res/monster.xlsx` against `xlsx-ts` and `xlsx dense`
- `npm run bench:check`
  - Run a 5-iteration comparison on `res/monster.xlsx` and validate correctness plus performance thresholds from `benchmarks/monster-baseline.json`
- `npm run bench:compare`
  - Equivalent wrapper around the compare script kept in the repo
- `node --import tsx scripts/benchmark.ts res/monster.xlsx 5`
  - Run the benchmark with a custom file path and iteration count
- `node --import tsx scripts/benchmark.ts res/monster.xlsx 5 --check benchmarks/monster-baseline.json`
  - Run the regression check against any benchmark file; the process exits non-zero when the thresholds are exceeded

## Current Limits

- The zip backend now uses pure JS via `fflate`, so it no longer depends on system `python3` or `zip`.
- The full zip package and all entries are still loaded into memory today, so peak memory usage for very large files can still be improved.
- String writes use `inlineStr` to avoid rebuilding `sharedStrings.xml` for simple value updates.
- APIs for merged comments, rich text, images, and similar parts are still missing.
- XML writes are implemented as local patches, not as a full OOXML object model.

## Development

```bash
npm run build
npm test
npm run validate:task
```

Where:

- `npm test` runs the TypeScript tests through `tsx`
- `npm run validate:task` runs the TypeScript validation script through `tsx`
- `npm run build` only produces `dist`

The test suite currently checks two things:

1. After an untouched roundtrip, every package part remains byte-for-byte identical.
2. After editing a styled cell, the style index is still preserved and `styles.xml` stays unchanged.

## Real File Validation

[`res/task.xlsx`](/Users/codetypes/Desktop/Github/xlsx-ts/res/task.xlsx) in the repository is a useful regression sample.

```bash
npm run validate:task
```

To validate any other file:

```bash
npm run validate:roundtrip -- path/to/file.xlsx
```
