# xlsx-ts

[English README](README.en.md)

一个以“无损优先”为核心的 XLSX 读写库原型。

目标不是先把 Excel 全模型映射成庞大的 JS 对象，而是先保证这一条底线：

`read(xlsx) -> write(xlsx)` 之后，解压出来的各个部件文件内容保持一致。

这条底线一旦成立，样式、主题、批注、关系文件、未知扩展节点都能被天然保住。后续再往上叠加单元格、公式、批注、图片等 API，风险会低很多。

## 设计思路

库分成两层：

1. `Lossless package layer`
   - 把 xlsx 当成 zip 包处理。
   - 所有 entry 先按原始字节保存。
   - 未修改的 entry 永远原样写回，不做重新序列化。

2. `Editable workbook layer`
   - 只针对确实需要改动的 XML 部件做局部 patch。
   - 当前原型先支持：
     - 读取工作表列表
     - 读取单元格
     - 修改单元格
     - 读取公式
     - 修改公式
   - 样式依赖的 `s="..."` 属性保留不动，因此样式不会因为写值被丢掉。

## 为什么这条路线适合“保样式”

大多数样式丢失，根源都不是 `styles.xml` 不会读，而是“写回时整个工作簿被重新生成”，导致：

- 未知节点丢失
- 属性顺序变化
- namespace/扩展标记被清洗
- 关系文件被重排
- shared strings / worksheet / styles 之间的耦合被误改

无损优先的路线反过来做：

- 先完整保住 zip 内所有 part
- 再只改需要改的 part
- 没改的 part 一律原样回写

这样就更容易通过“解压后内容一致”的验收标准。

## 当前能力

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

示例：

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

说明：

- 同一张工作表首次读写时会扫描一次 `sheetData`，建立单元格与行的位置索引
- `sheet.cell(address)` 返回可复用的 `Cell` 句柄，值/公式/样式索引会按工作表 revision 缓存
- `sheet.cell()` / `getCell()` / `setCell()` / `getFormula()` / `setFormula()` 现在同时支持 `A1` 地址和 `(rowNumber, column)` 两种调用方式；行列索引是从 `1` 开始
- 后续 `getCell` / `getFormula` 会直接走索引查找，不再每次整张表做字符串匹配
- `sheet.rowCount` / `sheet.columnCount` 当前表示已用区域的最大行号 / 最大列号；空表返回 `0`
- `sheet.getCellEntries()` / `iterCellEntries()` / `getRowEntries()` / `getColumnEntries()` 会按 worksheet 中真实存在的 `<c>` 节点返回带地址、行列号、类型、样式索引和值的对象，适合大表和稀疏表遍历
- `sheet.deleteCell()` 会真正移除 worksheet 里的 `<c>` 节点；如果你只是想保留样式占位但把值清空，继续用 `setCell(..., null)`
- `sheet.getStyleId()` / `setStyleId()` 当前读写单元格上的 `s="..."` 样式索引；支持 `A1` 和 `(rowNumber, column)` 两种调用，但还不直接编辑 `styles.xml`
- `sheet.getFreezePane()` / `freezePane()` / `unfreezePane()` 当前维护 worksheet `sheetViews/sheetView/pane`；插删行列时 `topLeftCell` 也会继续跟随更新
- `sheet.getSelection()` / `setSelection()` 当前读写 worksheet `sheetViews/sheetView/selection`；冻结窗格存在时会优先落在当前 active pane 对应的 selection 上
- 每次写入后会重建该表索引，保证后续读取拿到的是最新结果
- 修改工作表后会同步维护 `<dimension ref="...">`，避免使用范围信息过期
- `deleteRow()` / `deleteColumn()` 当前会同步更新本 sheet 的单元格坐标、公式引用、合并区域、`dimension`、常见 `ref/sqref` 属性、`definedNames`，以及其它 sheet 里显式引用它的公式
- `insertRow()` 当前会同步更新本 sheet 的单元格坐标、公式引用、合并区域、`dimension`、常见 `ref/sqref` 属性、`definedNames`，以及其它 sheet 里显式引用它的公式
- `insertColumn()` 当前会同步更新本 sheet 的单元格坐标、公式引用、合并区域、`dimension`、常见 `ref/sqref` 属性、`definedNames`，以及其它 sheet 里显式引用它的公式
- `sheet.getTables()` 当前可以读取已有 table 的名称、显示名、范围和部件路径
- `sheet.getHyperlinks()` 当前可以读取当前 sheet 上的内部和外部超链接；外部链接会解析 sheet rel 里的目标地址
- `sheet.getAutoFilter()` / `sheet.setAutoFilter()` / `sheet.removeAutoFilter()` 当前支持读写 worksheet 顶层 `autoFilter`，移除时会一并清掉顶层 `sortState`
- `sheet.getDataValidations()` / `sheet.setDataValidation()` / `sheet.removeDataValidation()` 当前支持读写 worksheet 顶层 `dataValidations`，包括常见属性与 `formula1/formula2`，并继续跟随插删行列维护 `sqref`
- `sheet.addTable()` 当前会创建最基础的 table part、sheet rel、`[Content_Types].xml` override 和 table XML；列名默认取范围首行，空列名会回退到 `ColumnN`
- `sheet.removeTable()` 当前会同步移除当前 sheet 的 `tableParts`、sheet rel、table XML 和对应的 content type override
- `sheet.setHyperlink()` / `sheet.removeHyperlink()` 当前支持维护 worksheet `<hyperlinks>` 与外部链接对应的 sheet rel，内部链接 target 用 `#Sheet1!A1` 这种格式
- 已有关联 table 在插删行列时会同步维护它们自己的 `ref` / `autoFilter`；如果整块 table 被删空，会从当前 sheet 的 `tableParts` 里移除
- `workbook.getDefinedNames()` / `getDefinedName()` / `setDefinedName()` / `deleteDefinedName()` 当前支持读写全局和本地 `definedNames`
- `workbook.getSheetVisibility()` / `setSheetVisibility()` 当前支持 `visible` / `hidden` / `veryHidden`；并会阻止把最后一张可见 sheet 隐藏掉
- `workbook.getActiveSheet()` / `setActiveSheet()` 当前读写 `workbookView.activeTab`；如果 workbook 里还没有 `bookViews`，会自动补上；隐藏 sheet 不允许设为 active
- `workbook.renameSheet()` / `sheet.rename()` 当前会同步维护 sheet 名、其它 sheet 的显式公式引用、`definedNames`、内部超链接位置和文档属性
- `workbook.moveSheet()` 当前使用 0-based `targetIndex`，会同步维护 workbook 里的 `<sheets>` 顺序、`docProps/app.xml` 里的工作表顺序、本地 `definedNames` 的 `localSheetId`，以及 `workbookView.activeTab`
- `workbook.addSheet()` / `workbook.deleteSheet()` 当前会同步维护 `workbook.xml`、rels、`[Content_Types].xml`，并在删除 sheet 时修正剩余公式与 `definedNames`

## 基准测试

仓库内现在包含一份已脱敏的大型基准文件 [res/monster.xlsx](/Users/codetypes/Desktop/Github/xlsx-ts/res/monster.xlsx)，可直接用于性能回归对比。

常用命令：

- `npm run bench:monster`
  - 对 `res/monster.xlsx` 运行 3 轮对比基准，比较 `xlsx-ts` 和 `xlsx dense`
- `npm run bench:check`
  - 对 `res/monster.xlsx` 运行 5 轮对比，并校验 `benchmarks/monster-baseline.json` 里的正确性与性能阈值
- `npm run bench:compare`
  - 等价于运行仓库里的对比脚本
- `node --import tsx scripts/benchmark.ts res/monster.xlsx 5`
  - 自定义文件路径和迭代次数
- `node --import tsx scripts/benchmark.ts res/monster.xlsx 5 --check benchmarks/monster-baseline.json`
  - 对任意基准文件执行回归检查；超出阈值时进程会以非零状态退出

 ## 当前限制

- zip 读写后端现在使用纯 JS 的 `fflate`，不再依赖系统里的 `python3` 与 `zip`
- 当前仍会把整个 zip 包与各个 entry 一起放进内存，对超大文件的峰值内存还可以继续优化
- 字符串写入使用 `inlineStr`，避免为了简单写值而重建 `sharedStrings.xml`
- 合并单元格、批注、富文本、图片等 API 还没加
- 对 XML 的写入是“局部 patch”，不是完整 OOXML 模型

## 开发

```bash
npm run build
npm test
npm run validate:task
```

其中：

- `npm test` 直接通过 `tsx` 运行 TypeScript 测试
- `npm run validate:task` 直接通过 `tsx` 运行 TypeScript 验证脚本
- `npm run build` 只负责产出 `dist`

测试里包含两件事：

1. 无修改 roundtrip 后，包内各个 part 的内容逐字节一致
2. 修改一个带样式的单元格后，样式索引仍被保留，`styles.xml` 不变

## 真实文件验证

仓库里的 [`res/task.xlsx`](/Users/codetypes/Desktop/Github/xlsx-ts/res/task.xlsx) 可以作为后续回归验证样本。

```bash
npm run validate:task
```

如果想验证任意文件：

```bash
npm run validate:roundtrip -- path/to/file.xlsx
```
