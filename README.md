# xlsx-ts

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
- `sheet.getCell(address)`
- `sheet.getRange(range)`
- `sheet.getUsedRange()`
- `sheet.setCell(address, value)`
- `sheet.setRange(startAddress, values)`
- `sheet.getFormula(address)`
- `sheet.setFormula(address, formula, options?)`
- `workbook.save(path)`

示例：

```ts
const workbook = await Workbook.open("input.xlsx");
const sheet = workbook.getSheet("Sheet1");

sheet.setCell("A1", "Hello");
sheet.setRange("B2", [
  [1, 2],
  [3, 4],
]);
sheet.setFormula("B1", "SUM(1,2)", { cachedValue: 3 });

await workbook.save("output.xlsx");
```

说明：

- 同一张工作表首次读写时会扫描一次 `sheetData`，建立单元格与行的位置索引
- 后续 `getCell` / `getFormula` 会直接走索引查找，不再每次整张表做字符串匹配
- 每次写入后会重建该表索引，保证后续读取拿到的是最新结果

## 当前限制

- zip 读写后端暂时依赖系统里的 `python3` 与 `zip`
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
