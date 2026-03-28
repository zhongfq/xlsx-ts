import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, readdir, rm, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";

import { Workbook } from "../src/index.ts";

test("roundtrip keeps extracted parts identical", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const tempRoot = await mkdtemp(join(tmpdir(), "xlsx-ts-test-"));

  try {
    const inputPath = join(tempRoot, "input.xlsx");
    const outputPath = join(tempRoot, "output.xlsx");
    const expectedEntries = await loadFixtureEntries(fixtureDir);

    const sourceDocument = Workbook.fromEntries(expectedEntries);
    await sourceDocument.save(inputPath);

    const reopened = await Workbook.open(inputPath);
    await reopened.save(outputPath);

    const actualEntries = await Workbook.open(outputPath);
    assertEntryMapsEqual(toEntryMap(expectedEntries), toEntryMap(actualEntries.toEntries()));
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("editing a styled cell keeps its style index and leaves styles.xml untouched", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);
  const originalStyles = entryText(entries, "xl/styles.xml");
  const sheet = workbook.getSheet("Sheet1");

  assert.equal(sheet.getCell("A1"), "Hello");

  sheet.setCell("A1", "World");

  const nextEntries = workbook.toEntries();
  const sheetXml = entryText(nextEntries, "xl/worksheets/sheet1.xml");
  const stylesXml = entryText(nextEntries, "xl/styles.xml");

  assert.match(sheetXml, /<c r="A1" t="inlineStr" s="1">/);
  assert.match(sheetXml, /<t>World<\/t>/);
  assert.equal(stylesXml, originalStyles);
});

test("sheet reads stay coherent after repeated writes", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.setCell("B1", 1);
  assert.equal(sheet.getCell("B1"), 1);

  sheet.setCell("B1", 2);
  assert.equal(sheet.getCell("B1"), 2);

  sheet.setCell("A2", "Tail");
  assert.equal(sheet.getCell("A2"), "Tail");
});

test("formula cells can be read and updated without dropping styles", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" s="1"><f>SUM(1,2)</f><v>3</v></c>
    </row>
  </sheetData>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  assert.equal(sheet.getFormula("A1"), "SUM(1,2)");
  assert.equal(sheet.getCell("A1"), 3);

  sheet.setFormula("A1", 'CONCAT("He","llo")', { cachedValue: "Hello" });

  assert.equal(sheet.getFormula("A1"), 'CONCAT("He","llo")');
  assert.equal(sheet.getCell("A1"), "Hello");

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<c r="A1" t="str" s="1">/);
  assert.match(sheetXml, /<f>CONCAT\(&quot;He&quot;,&quot;llo&quot;\)<\/f>/);
  assert.match(sheetXml, /<v>Hello<\/v>/);
});

test("cell handle objects cache parsed state and refresh after writes", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");
  const first = sheet.cell("A1");
  const second = sheet.cell("A1");

  assert.equal(first, second);
  assert.equal(first.exists, true);
  assert.equal(first.type, "string");
  assert.equal(first.styleId, 1);
  assert.equal(first.formula, null);
  assert.equal(first.value, "Hello");

  first.setValue("World");

  assert.equal(first.value, "World");
  assert.equal(first.type, "string");

  sheet.setFormula("A1", "SUM(1,2)", { cachedValue: 3 });

  assert.equal(first.formula, "SUM(1,2)");
  assert.equal(first.type, "formula");
  assert.equal(first.value, 3);

  const missing = sheet.cell("C9");
  assert.equal(missing.exists, false);
  assert.equal(missing.type, "missing");
  assert.equal(missing.value, null);
});

test("range APIs read and write rectangular values", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  assert.equal(sheet.getUsedRange(), "A1");
  assert.deepEqual(sheet.getRange("A1:B2"), [["Hello", null], [null, null]]);

  sheet.setRange("B2", [
    [1, 2],
    [3, 4],
  ]);

  assert.equal(sheet.getUsedRange(), "A1:C3");
  assert.deepEqual(sheet.getRange("A1:C3"), [
    ["Hello", null, null],
    [null, 1, 2],
    [null, 3, 4],
  ]);

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<row r="2"><c r="B2"><v>1<\/v><\/c><c r="C2"><v>2<\/v><\/c><\/row>/);
  assert.match(sheetXml, /<row r="3"><c r="B3"><v>3<\/v><\/c><c r="C3"><v>4<\/v><\/c><\/row>/);
});

test("sheet rowCount and columnCount track the used bounds", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData></sheetData>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  assert.equal(sheet.rowCount, 0);
  assert.equal(sheet.columnCount, 0);

  sheet.setCell("C5", 1);

  assert.equal(sheet.rowCount, 5);
  assert.equal(sheet.columnCount, 3);
  assert.equal(sheet.getUsedRange(), "C5");

  sheet.setCell("A1", "Top");

  assert.equal(sheet.rowCount, 5);
  assert.equal(sheet.columnCount, 3);
  assert.equal(sheet.getUsedRange(), "A1:C5");

  sheet.deleteColumn("B");

  assert.equal(sheet.rowCount, 5);
  assert.equal(sheet.columnCount, 2);
  assert.equal(sheet.getUsedRange(), "A1:B5");
});

test("insertColumn shifts cell addresses, formulas, and merged ranges together", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:C2"/>
  <sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
      <c r="B1"><v>2</v></c>
      <c r="C1"><f>SUM(A1:B1)</f><v>3</v></c>
    </row>
    <row r="2">
      <c r="A2"><f>Sheet1!B1</f><v>2</v></c>
      <c r="B2"><v>4</v></c>
      <c r="C2"><v>5</v></c>
    </row>
  </sheetData>
  <mergeCells count="1"><mergeCell ref="B2:C2"/></mergeCells>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.insertColumn("B");

  assert.equal(sheet.getCell("A1"), 1);
  assert.equal(sheet.getCell("B1"), null);
  assert.equal(sheet.getCell("C1"), 2);
  assert.equal(sheet.getCell("D1"), 3);
  assert.equal(sheet.getFormula("D1"), "SUM(A1:C1)");
  assert.equal(sheet.getFormula("A2"), "Sheet1!C1");
  assert.deepEqual(sheet.getMergedRanges(), ["C2:D2"]);
  assert.equal(sheet.getUsedRange(), "A1:D2");

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<c r="C1"><v>2<\/v><\/c>/);
  assert.match(sheetXml, /<c r="D1"><f>SUM\(A1:C1\)<\/f><v>3<\/v><\/c>/);
  assert.match(sheetXml, /<c r="A2"><f>Sheet1!C1<\/f><v>2<\/v><\/c>/);
  assert.match(sheetXml, /<mergeCell ref="C2:D2"\/>/);
  assert.match(sheetXml, /<dimension ref="A1:D2"\/>/);
});

test("insertRow shifts cell addresses, formulas, and merged ranges together", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:B3"/>
  <sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
      <c r="B1"><f>SUM(A2:B2)</f><v>3</v></c>
    </row>
    <row r="2">
      <c r="A2"><v>2</v></c>
      <c r="B2"><v>4</v></c>
    </row>
    <row r="3">
      <c r="A3"><f>Sheet1!A2</f><v>2</v></c>
      <c r="B3"><v>5</v></c>
    </row>
  </sheetData>
  <mergeCells count="1"><mergeCell ref="A2:B3"/></mergeCells>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.insertRow(2);

  assert.equal(sheet.getCell("A1"), 1);
  assert.equal(sheet.getCell("A2"), null);
  assert.equal(sheet.getCell("A3"), 2);
  assert.equal(sheet.getCell("A4"), 2);
  assert.equal(sheet.getFormula("B1"), "SUM(A3:B3)");
  assert.equal(sheet.getFormula("A4"), "Sheet1!A3");
  assert.deepEqual(sheet.getMergedRanges(), ["A3:B4"]);
  assert.equal(sheet.getUsedRange(), "A1:B4");

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<row r="3">[\s\S]*<c r="A3"><v>2<\/v><\/c>[\s\S]*<c r="B3"><v>4<\/v><\/c>[\s\S]*<\/row>/);
  assert.match(sheetXml, /<row r="4">[\s\S]*<c r="A4"><f>Sheet1!A3<\/f><v>2<\/v><\/c>[\s\S]*<\/row>/);
  assert.match(sheetXml, /<c r="B1"><f>SUM\(A3:B3\)<\/f><v>3<\/v><\/c>/);
  assert.match(sheetXml, /<mergeCell ref="A3:B4"\/>/);
  assert.match(sheetXml, /<dimension ref="A1:B4"\/>/);
});

test("insertColumn updates worksheet ref attributes and defined names", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    replaceEntryText(
      await loadFixtureEntries(fixtureDir),
      "xl/workbook.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
  <definedNames>
    <definedName name="_xlnm.Print_Area" localSheetId="0">$A$1:$C$4</definedName>
    <definedName name="DataRange">Sheet1!$B$2:$C$4</definedName>
  </definedNames>
</workbook>`,
    ),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews><sheetView workbookViewId="0"><selection activeCell="B2" sqref="B2:C2"/></sheetView></sheetViews>
  <sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
      <c r="B1"><v>2</v></c>
      <c r="C1"><v>3</v></c>
    </row>
    <row r="2">
      <c r="A2"><v>4</v></c>
      <c r="B2"><v>5</v></c>
      <c r="C2"><v>6</v></c>
    </row>
  </sheetData>
  <autoFilter ref="A1:C4"/>
  <sortState ref="A2:C4"/>
  <conditionalFormatting sqref="B2:C4"><cfRule type="expression" priority="1"><formula>B2&gt;0</formula></cfRule></conditionalFormatting>
  <dataValidations count="1"><dataValidation type="whole" sqref="A2:B4"/></dataValidations>
  <hyperlinks><hyperlink ref="C2" location="#Sheet1!A1"/></hyperlinks>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.insertColumn("B");

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  const workbookXml = entryText(workbook.toEntries(), "xl/workbook.xml");

  assert.match(sheetXml, /<selection activeCell="C2" sqref="C2:D2"\/>/);
  assert.match(sheetXml, /<autoFilter ref="A1:D4"\/>/);
  assert.match(sheetXml, /<sortState ref="A2:D4"\/>/);
  assert.match(sheetXml, /<conditionalFormatting sqref="C2:D4">/);
  assert.match(sheetXml, /<dataValidations count="1"><dataValidation type="whole" sqref="A2:C4"\/><\/dataValidations>/);
  assert.match(sheetXml, /<hyperlinks><hyperlink ref="D2" location="#Sheet1!A1"\/><\/hyperlinks>/);
  assert.match(workbookXml, /<definedName name="_xlnm.Print_Area" localSheetId="0">\$A\$1:\$D\$4<\/definedName>/);
  assert.match(workbookXml, /<definedName name="DataRange">Sheet1!\$C\$2:\$D\$4<\/definedName>/);
});

test("insertColumn updates formulas in other sheets that reference the edited sheet", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = withSecondSheet(
    replaceEntryText(
      await loadFixtureEntries(fixtureDir),
      "xl/worksheets/sheet1.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
      <c r="B1"><v>2</v></c>
    </row>
  </sheetData>
</worksheet>`,
    ),
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><f>SUM(Sheet1!A1:B1)</f><v>3</v></c>
    </row>
    <row r="2">
      <c r="A2"><f>Sheet1!B1</f><v>2</v></c>
    </row>
  </sheetData>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet1 = workbook.getSheet("Sheet1");
  const sheet2 = workbook.getSheet("Sheet2");

  sheet1.insertColumn("B");

  assert.equal(sheet2.getFormula("A1"), "SUM(Sheet1!A1:C1)");
  assert.equal(sheet2.getFormula("A2"), "Sheet1!C1");

  const sheet2Xml = entryText(workbook.toEntries(), "xl/worksheets/sheet2.xml");
  assert.match(sheet2Xml, /<c r="A1"><f>SUM\(Sheet1!A1:C1\)<\/f><v>3<\/v><\/c>/);
  assert.match(sheet2Xml, /<c r="A2"><f>Sheet1!C1<\/f><v>2<\/v><\/c>/);
});

test("deleteColumn shifts cells, shrinks ranges, and emits #REF! for deleted refs", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:D2"/>
  <sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
      <c r="B1"><v>2</v></c>
      <c r="C1"><f>SUM(A1:D1)</f><v>10</v></c>
      <c r="D1"><f>B1</f><v>2</v></c>
    </row>
    <row r="2">
      <c r="B2"><v>4</v></c>
      <c r="C2"><v>5</v></c>
      <c r="D2"><v>6</v></c>
    </row>
  </sheetData>
  <mergeCells count="1"><mergeCell ref="B2:D2"/></mergeCells>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.deleteColumn("B");

  assert.equal(sheet.getCell("A1"), 1);
  assert.equal(sheet.getCell("B1"), 10);
  assert.equal(sheet.getFormula("B1"), "SUM(A1:C1)");
  assert.equal(sheet.getCell("C1"), 2);
  assert.equal(sheet.getFormula("C1"), "#REF!");
  assert.deepEqual(sheet.getMergedRanges(), ["B2:C2"]);
  assert.equal(sheet.getUsedRange(), "A1:C2");

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.doesNotMatch(sheetXml, /<c r="B2"><v>4<\/v><\/c>/);
  assert.match(sheetXml, /<c r="B1"><f>SUM\(A1:C1\)<\/f><v>10<\/v><\/c>/);
  assert.match(sheetXml, /<c r="C1"><f>#REF!<\/f><v>2<\/v><\/c>/);
  assert.match(sheetXml, /<mergeCell ref="B2:C2"\/>/);
});

test("deleteRow updates formulas in other sheets that reference the edited sheet", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = withSecondSheet(
    replaceEntryText(
      await loadFixtureEntries(fixtureDir),
      "xl/worksheets/sheet1.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
      <c r="B1"><v>2</v></c>
    </row>
    <row r="2">
      <c r="A2"><v>3</v></c>
      <c r="B2"><v>4</v></c>
    </row>
    <row r="3">
      <c r="A3"><v>5</v></c>
      <c r="B3"><v>6</v></c>
    </row>
    <row r="4">
      <c r="A4"><v>7</v></c>
      <c r="B4"><v>8</v></c>
    </row>
  </sheetData>
</worksheet>`,
    ),
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><f>SUM(Sheet1!A1:B4)</f><v>36</v></c>
    </row>
    <row r="2">
      <c r="A2"><f>Sheet1!A2</f><v>3</v></c>
    </row>
  </sheetData>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet1 = workbook.getSheet("Sheet1");
  const sheet2 = workbook.getSheet("Sheet2");

  sheet1.deleteRow(2);

  assert.equal(sheet2.getFormula("A1"), "SUM(Sheet1!A1:B3)");
  assert.equal(sheet2.getFormula("A2"), "#REF!");

  const sheet2Xml = entryText(workbook.toEntries(), "xl/worksheets/sheet2.xml");
  assert.match(sheet2Xml, /<c r="A1"><f>SUM\(Sheet1!A1:B3\)<\/f><v>36<\/v><\/c>/);
  assert.match(sheet2Xml, /<c r="A2"><f>#REF!<\/f><v>3<\/v><\/c>/);
});

test("deleteRow shifts cells, shrinks ranges, and emits #REF! for deleted refs", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:B4"/>
  <sheetData>
    <row r="1">
      <c r="A1"><f>SUM(A1:B4)</f><v>1</v></c>
    </row>
    <row r="2">
      <c r="A2"><v>2</v></c>
    </row>
    <row r="3">
      <c r="A3"><f>A2</f><v>2</v></c>
      <c r="B3"><v>3</v></c>
    </row>
    <row r="4">
      <c r="A4"><v>4</v></c>
      <c r="B4"><v>5</v></c>
    </row>
  </sheetData>
  <mergeCells count="1"><mergeCell ref="A2:B4"/></mergeCells>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.deleteRow(2);

  assert.equal(sheet.getFormula("A1"), "SUM(A1:B3)");
  assert.equal(sheet.getFormula("A2"), "#REF!");
  assert.equal(sheet.getCell("A3"), 4);
  assert.deepEqual(sheet.getMergedRanges(), ["A2:B3"]);
  assert.equal(sheet.getUsedRange(), "A1:B3");

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.doesNotMatch(sheetXml, /<row r="4">/);
  assert.match(sheetXml, /<row r="2">[\s\S]*<c r="A2"><f>#REF!<\/f><v>2<\/v><\/c>[\s\S]*<c r="B2"><v>3<\/v><\/c>[\s\S]*<\/row>/);
  assert.match(sheetXml, /<mergeCell ref="A2:B3"\/>/);
});

test("deleteRow updates worksheet ref attributes and defined names", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    replaceEntryText(
      await loadFixtureEntries(fixtureDir),
      "xl/workbook.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
  <definedNames>
    <definedName name="_xlnm.Print_Area" localSheetId="0">$A$1:$C$4</definedName>
    <definedName name="DataRange">Sheet1!$B$2:$C$4</definedName>
  </definedNames>
</workbook>`,
    ),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews><sheetView workbookViewId="0"><selection activeCell="B3" sqref="B3:C3"/></sheetView></sheetViews>
  <sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
      <c r="B1"><v>2</v></c>
      <c r="C1"><v>3</v></c>
    </row>
    <row r="2">
      <c r="A2"><v>4</v></c>
      <c r="B2"><v>5</v></c>
      <c r="C2"><v>6</v></c>
    </row>
    <row r="3">
      <c r="A3"><v>7</v></c>
      <c r="B3"><v>8</v></c>
      <c r="C3"><v>9</v></c>
    </row>
    <row r="4">
      <c r="A4"><v>10</v></c>
      <c r="B4"><v>11</v></c>
      <c r="C4"><v>12</v></c>
    </row>
  </sheetData>
  <autoFilter ref="A1:C4"/>
  <sortState ref="A2:C4"/>
  <conditionalFormatting sqref="B2:C4"><cfRule type="expression" priority="1"><formula>B2&gt;0</formula></cfRule></conditionalFormatting>
  <dataValidations count="1"><dataValidation type="whole" sqref="A2:B4"/></dataValidations>
  <hyperlinks><hyperlink ref="C3" location="#Sheet1!A1"/></hyperlinks>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.deleteRow(2);

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  const workbookXml = entryText(workbook.toEntries(), "xl/workbook.xml");

  assert.match(sheetXml, /<selection activeCell="B2" sqref="B2:C2"\/>/);
  assert.match(sheetXml, /<autoFilter ref="A1:C3"\/>/);
  assert.match(sheetXml, /<sortState ref="A2:C3"\/>/);
  assert.match(sheetXml, /<conditionalFormatting sqref="B2:C3">/);
  assert.match(sheetXml, /<dataValidations count="1"><dataValidation type="whole" sqref="A2:B3"\/><\/dataValidations>/);
  assert.match(sheetXml, /<hyperlinks><hyperlink ref="C2" location="#Sheet1!A1"\/><\/hyperlinks>/);
  assert.match(workbookXml, /<definedName name="_xlnm.Print_Area" localSheetId="0">\$A\$1:\$C\$3<\/definedName>/);
  assert.match(workbookXml, /<definedName name="DataRange">Sheet1!\$B\$2:\$C\$3<\/definedName>/);
});

test("sheet getTables reads existing tables and insertColumn updates table refs", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = withSheetTable(
    replaceEntryText(
      await loadFixtureEntries(fixtureDir),
      "xl/worksheets/sheet1.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
      <c r="B1"><v>2</v></c>
    </row>
    <row r="2">
      <c r="A2"><v>3</v></c>
      <c r="B2"><v>4</v></c>
    </row>
    <row r="3">
      <c r="A3"><v>5</v></c>
      <c r="B3"><v>6</v></c>
    </row>
  </sheetData>
  <tableParts count="1"><tablePart r:id="rIdTable1"/></tableParts>
</worksheet>`,
    ),
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Sales" displayName="Sales" ref="A1:B3" totalsRowShown="0">
  <autoFilter ref="A1:B3"/>
  <tableColumns count="2">
    <tableColumn id="1" name="A"/>
    <tableColumn id="2" name="B"/>
  </tableColumns>
</table>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  assert.deepEqual(sheet.getTables(), [
    { name: "Sales", displayName: "Sales", range: "A1:B3", path: "xl/tables/table1.xml" },
  ]);

  sheet.insertColumn("B");

  assert.deepEqual(sheet.getTables(), [
    { name: "Sales", displayName: "Sales", range: "A1:C3", path: "xl/tables/table1.xml" },
  ]);

  const tableXml = entryText(workbook.toEntries(), "xl/tables/table1.xml");
  assert.match(tableXml, /<table [^>]*ref="A1:C3"[^>]*>/);
  assert.match(tableXml, /<autoFilter ref="A1:C3"\/>/);
});

test("deleteRow removes table parts when the full table range is deleted", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = withSheetTable(
    replaceEntryText(
      await loadFixtureEntries(fixtureDir),
      "xl/worksheets/sheet1.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
      <c r="B1"><v>2</v></c>
    </row>
    <row r="2">
      <c r="A2"><v>3</v></c>
      <c r="B2"><v>4</v></c>
    </row>
    <row r="3">
      <c r="A3"><v>5</v></c>
      <c r="B3"><v>6</v></c>
    </row>
  </sheetData>
  <tableParts count="1"><tablePart r:id="rIdTable1"/></tableParts>
</worksheet>`,
    ),
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Sales" displayName="Sales" ref="A1:B3" totalsRowShown="0">
  <autoFilter ref="A1:B3"/>
  <tableColumns count="2">
    <tableColumn id="1" name="A"/>
    <tableColumn id="2" name="B"/>
  </tableColumns>
</table>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.deleteRow(1, 3);

  assert.deepEqual(sheet.getTables(), []);
  assert.equal(workbook.listEntries().includes("xl/tables/table1.xml"), false);

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  const relsXml = entryText(workbook.toEntries(), "xl/worksheets/_rels/sheet1.xml.rels");
  const contentTypesXml = entryText(workbook.toEntries(), "[Content_Types].xml");
  assert.doesNotMatch(sheetXml, /<tableParts\b/);
  assert.doesNotMatch(relsXml, /relationships\/table/);
  assert.doesNotMatch(contentTypesXml, /spreadsheetml\.table\+xml/);
});

test("sheet addTable and removeTable manage package parts", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.setRow(1, ["Name", "Score"]);
  sheet.setRow(2, ["Alice", 98]);
  sheet.setRow(3, ["Bob", 87]);

  const table = sheet.addTable("A1:B3", { name: "Scores" });

  assert.deepEqual(table, {
    name: "Scores",
    displayName: "Scores",
    range: "A1:B3",
    path: "xl/tables/table1.xml",
  });
  assert.deepEqual(sheet.getTables(), [table]);

  let sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  let relsXml = entryText(workbook.toEntries(), "xl/worksheets/_rels/sheet1.xml.rels");
  let contentTypesXml = entryText(workbook.toEntries(), "[Content_Types].xml");
  let tableXml = entryText(workbook.toEntries(), "xl/tables/table1.xml");

  assert.match(sheetXml, /<tableParts count="1"><tablePart r:id="rId1"\/><\/tableParts>/);
  assert.match(relsXml, /<Relationship Id="rId1" Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/table" Target="\.\.\/tables\/table1\.xml"\/>/);
  assert.match(contentTypesXml, /<Override PartName="\/xl\/tables\/table1\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.spreadsheetml\.table\+xml"\/>/);
  assert.match(tableXml, /<table [^>]*name="Scores" displayName="Scores" ref="A1:B3"[^>]*>/);
  assert.match(tableXml, /<tableColumn id="1" name="Name"\/>/);
  assert.match(tableXml, /<tableColumn id="2" name="Score"\/>/);

  sheet.removeTable("Scores");

  assert.deepEqual(sheet.getTables(), []);
  assert.equal(workbook.listEntries().includes("xl/tables/table1.xml"), false);

  sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  relsXml = entryText(workbook.toEntries(), "xl/worksheets/_rels/sheet1.xml.rels");
  contentTypesXml = entryText(workbook.toEntries(), "[Content_Types].xml");

  assert.doesNotMatch(sheetXml, /<tableParts\b/);
  assert.doesNotMatch(relsXml, /relationships\/table/);
  assert.doesNotMatch(contentTypesXml, /spreadsheetml\.table\+xml/);
});

test("row APIs read sparse rows and write from a column offset", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  assert.deepEqual(sheet.getRow(1), ["Hello"]);
  assert.deepEqual(sheet.getRow(4), []);

  sheet.setRow(4, ["Name", null, "Score"], 2);

  assert.deepEqual(sheet.getRow(4), [null, "Name", null, "Score"]);
  assert.equal(sheet.getUsedRange(), "A1:D4");

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<row r="4"><c r="B4" t="inlineStr"><is><t>Name<\/t><\/is><\/c><c r="C4"\/><c r="D4" t="inlineStr"><is><t>Score<\/t><\/is><\/c><\/row>/);
  assert.match(sheetXml, /<dimension ref="A1:D4"\/>/);
});

test("column APIs read sparse columns and write from a row offset", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  assert.deepEqual(sheet.getColumn("A"), ["Hello"]);
  assert.deepEqual(sheet.getColumn(3), []);

  sheet.setColumn("C", ["Q1", null, "Q3"], 2);

  assert.deepEqual(sheet.getColumn("C"), [null, "Q1", null, "Q3"]);
  assert.equal(sheet.getUsedRange(), "A1:C4");

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<row r="2"><c r="C2" t="inlineStr"><is><t>Q1<\/t><\/is><\/c><\/row>/);
  assert.match(sheetXml, /<row r="3"><c r="C3"\/><\/row>/);
  assert.match(sheetXml, /<row r="4"><c r="C4" t="inlineStr"><is><t>Q3<\/t><\/is><\/c><\/row>/);
});

test("header APIs read and write header rows", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.setHeaders(["Name", "Score"]);

  assert.deepEqual(sheet.getHeaders(), ["Name", "Score"]);

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(
    sheetXml,
    /<row r="1">[\s\S]*<c r="A1" t="inlineStr" s="1"><is><t>Name<\/t><\/is><\/c>[\s\S]*<c r="B1" t="inlineStr"><is><t>Score<\/t><\/is><\/c>[\s\S]*<\/row>/,
  );
});

test("append row APIs add rows at the sheet tail", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  const firstRow = sheet.appendRow(["Tail", 1], 2);
  const nextRows = sheet.appendRows([
    ["Tail-2", 2],
    ["Tail-3", 3],
  ], 2);

  assert.equal(firstRow, 2);
  assert.deepEqual(nextRows, [3, 4]);
  assert.deepEqual(sheet.getRow(4), [null, "Tail-3", 3]);

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<row r="2"><c r="B2" t="inlineStr"><is><t>Tail<\/t><\/is><\/c><c r="C2"><v>1<\/v><\/c><\/row>/);
  assert.match(sheetXml, /<row r="4"><c r="B4" t="inlineStr"><is><t>Tail-3<\/t><\/is><\/c><c r="C4"><v>3<\/v><\/c><\/row>/);
});

test("record APIs map rows by header cells", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.setRow(1, ["Name", "Score"]);
  sheet.setRow(2, ["Alice", 98]);
  sheet.setRow(4, ["Bob", 87]);

  assert.deepEqual(sheet.getRecords(), [
    { Name: "Alice", Score: 98 },
    { Name: "Bob", Score: 87 },
  ]);

  sheet.addRecord({ Name: "Cara", Score: 91 });

  assert.deepEqual(sheet.getRecords(), [
    { Name: "Alice", Score: 98 },
    { Name: "Bob", Score: 87 },
    { Name: "Cara", Score: 91 },
  ]);

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<row r="5"><c r="A5" t="inlineStr"><is><t>Cara<\/t><\/is><\/c><c r="B5"><v>91<\/v><\/c><\/row>/);
});

test("record APIs can append multiple records in order", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.setRow(1, ["Name", "Score"]);
  sheet.addRecords([
    { Name: "Alice", Score: 98 },
    { Name: "Bob", Score: 87 },
  ]);

  assert.deepEqual(sheet.getRecords(), [
    { Name: "Alice", Score: 98 },
    { Name: "Bob", Score: 87 },
  ]);

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<row r="2"><c r="A2" t="inlineStr"><is><t>Alice<\/t><\/is><\/c><c r="B2"><v>98<\/v><\/c><\/row>/);
  assert.match(sheetXml, /<row r="3"><c r="A3" t="inlineStr"><is><t>Bob<\/t><\/is><\/c><c r="B3"><v>87<\/v><\/c><\/row>/);
});

test("record APIs can read and update a specific record row", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.setRow(1, ["Name", "Score"]);
  sheet.setRow(2, ["Alice", 98]);

  assert.deepEqual(sheet.getRecord(2), { Name: "Alice", Score: 98 });

  sheet.setRecord(2, { Name: "Alicia", Score: 99 });

  assert.deepEqual(sheet.getRecord(2), { Name: "Alicia", Score: 99 });

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<row r="2"><c r="A2" t="inlineStr"><is><t>Alicia<\/t><\/is><\/c><c r="B2"><v>99<\/v><\/c><\/row>/);
});

test("record APIs can replace the full record set", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.setRow(1, ["Name", "Score"]);
  sheet.addRecords([
    { Name: "Alice", Score: 98 },
    { Name: "Bob", Score: 87 },
    { Name: "Cara", Score: 91 },
  ]);

  sheet.setRecords([
    { Name: "Zoe", Score: 100 },
    { Name: "Yan" },
  ]);

  assert.deepEqual(sheet.getRecords(), [
    { Name: "Zoe", Score: 100 },
    { Name: "Yan", Score: null },
  ]);

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<row r="2"><c r="A2" t="inlineStr"><is><t>Zoe<\/t><\/is><\/c><c r="B2"><v>100<\/v><\/c><\/row>/);
  assert.match(sheetXml, /<row r="3"><c r="A3" t="inlineStr"><is><t>Yan<\/t><\/is><\/c><c r="B3"\/><\/row>/);
  assert.doesNotMatch(sheetXml, /<row r="4">/);
});

test("record APIs can delete a record row", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.setRow(1, ["Name", "Score"]);
  sheet.addRecords([
    { Name: "Alice", Score: 98 },
    { Name: "Bob", Score: 87 },
  ]);

  sheet.deleteRecord(2);

  assert.equal(sheet.getRecord(2), null);
  assert.deepEqual(sheet.getRecords(), [{ Name: "Bob", Score: 87 }]);

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.doesNotMatch(sheetXml, /<row r="2">/);
  assert.match(sheetXml, /<dimension ref="A1:B3"\/>/);
});

test("record APIs can delete multiple record rows", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.setRow(1, ["Name", "Score"]);
  sheet.addRecords([
    { Name: "Alice", Score: 98 },
    { Name: "Bob", Score: 87 },
    { Name: "Cara", Score: 91 },
  ]);

  sheet.deleteRecords([2, 4, 2]);

  assert.equal(sheet.getRecord(2), null);
  assert.equal(sheet.getRecord(4), null);
  assert.deepEqual(sheet.getRecords(), [{ Name: "Bob", Score: 87 }]);

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.doesNotMatch(sheetXml, /<row r="2">/);
  assert.match(sheetXml, /<row r="3"><c r="A3" t="inlineStr"><is><t>Bob<\/t><\/is><\/c><c r="B3"><v>87<\/v><\/c><\/row>/);
  assert.doesNotMatch(sheetXml, /<row r="4">/);
});

test("merged range APIs patch mergeCells without touching unrelated parts", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  assert.deepEqual(sheet.getMergedRanges(), []);

  sheet.addMergedRange("B2:A1");
  sheet.addMergedRange("C3:D4");
  sheet.addMergedRange("A1:B2");

  assert.deepEqual(sheet.getMergedRanges(), ["A1:B2", "C3:D4"]);

  let sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(
    sheetXml,
    /<\/sheetData><mergeCells count="2"><mergeCell ref="A1:B2"\/><mergeCell ref="C3:D4"\/><\/mergeCells>\s*<\/worksheet>/,
  );

  sheet.removeMergedRange("A1:B2");
  assert.deepEqual(sheet.getMergedRanges(), ["C3:D4"]);

  sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<mergeCells count="1"><mergeCell ref="C3:D4"\/><\/mergeCells>/);

  sheet.removeMergedRange("C3:D4");
  assert.deepEqual(sheet.getMergedRanges(), []);

  sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.doesNotMatch(sheetXml, /<mergeCells\b/);
});

test("writing cells keeps worksheet dimension ref in sync", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetData><row r="1"><c r="A1" s="1" t="inlineStr"><is><t>Hello</t></is></c></row></sheetData></worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.setCell("C4", 9);

  const sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<dimension ref="A1:C4"\/>/);
});

test("workbook can add a sheet and wire workbook metadata", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "docProps/app.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>xlsx-ts</Application>
  <HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs>
  <TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>Sheet1</vt:lpstr></vt:vector></TitlesOfParts>
</Properties>`,
  );
  const workbook = Workbook.fromEntries(entries);

  const newSheet = workbook.addSheet("Sheet2");
  newSheet.setCell("A1", "New");

  assert.deepEqual(workbook.getSheets().map((sheet) => sheet.name), ["Sheet1", "Sheet2"]);
  assert.equal(newSheet.getCell("A1"), "New");

  const workbookXml = entryText(workbook.toEntries(), "xl/workbook.xml");
  const relsXml = entryText(workbook.toEntries(), "xl/_rels/workbook.xml.rels");
  const contentTypesXml = entryText(workbook.toEntries(), "[Content_Types].xml");
  const sheet2Xml = entryText(workbook.toEntries(), "xl/worksheets/sheet2.xml");
  const appXml = entryText(workbook.toEntries(), "docProps/app.xml");

  assert.match(workbookXml, /<sheet name="Sheet2" sheetId="2" r:id="rId3"\/>/);
  assert.match(relsXml, /<Relationship Id="rId3" Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/worksheet" Target="worksheets\/sheet2\.xml"\/>/);
  assert.match(contentTypesXml, /<Override PartName="\/xl\/worksheets\/sheet2\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.spreadsheetml\.worksheet\+xml"\/>/);
  assert.match(sheet2Xml, /<row r="1"><c r="A1" t="inlineStr"><is><t>New<\/t><\/is><\/c><\/row>/);
  assert.match(appXml, /<vt:i4>2<\/vt:i4>/);
  assert.match(appXml, /<vt:lpstr>Sheet1<\/vt:lpstr><vt:lpstr>Sheet2<\/vt:lpstr>/);
});

test("workbook can delete a sheet and rewrite remaining references", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    replaceEntryText(
      replaceEntryText(
        withSecondSheet(
          await loadFixtureEntries(fixtureDir),
          `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>9</v></c>
    </row>
  </sheetData>
</worksheet>`,
        ),
        "xl/workbook.xml",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="Sheet2" sheetId="2" r:id="rId3"/>
  </sheets>
  <definedNames>
    <definedName name="ExternalRef">Sheet2!$A$1</definedName>
    <definedName name="LocalToSheet2" localSheetId="1">$A$1</definedName>
  </definedNames>
</workbook>`,
      ),
      "xl/worksheets/sheet1.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><f>Sheet2!A1</f><v>9</v></c>
    </row>
  </sheetData>
</worksheet>`,
    ),
    "docProps/app.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>xlsx-ts</Application>
  <HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>2</vt:i4></vt:variant></vt:vector></HeadingPairs>
  <TitlesOfParts><vt:vector size="2" baseType="lpstr"><vt:lpstr>Sheet1</vt:lpstr><vt:lpstr>Sheet2</vt:lpstr></vt:vector></TitlesOfParts>
</Properties>`,
  );
  const workbook = Workbook.fromEntries(entries);

  workbook.deleteSheet("Sheet2");

  assert.deepEqual(workbook.getSheets().map((sheet) => sheet.name), ["Sheet1"]);
  assert.equal(workbook.getSheet("Sheet1").getFormula("A1"), "#REF!");
  assert.equal(workbook.listEntries().includes("xl/worksheets/sheet2.xml"), false);

  const workbookXml = entryText(workbook.toEntries(), "xl/workbook.xml");
  const relsXml = entryText(workbook.toEntries(), "xl/_rels/workbook.xml.rels");
  const contentTypesXml = entryText(workbook.toEntries(), "[Content_Types].xml");
  const appXml = entryText(workbook.toEntries(), "docProps/app.xml");

  assert.doesNotMatch(workbookXml, /Sheet2/);
  assert.match(workbookXml, /<definedName name="ExternalRef">#REF!<\/definedName>/);
  assert.doesNotMatch(workbookXml, /LocalToSheet2/);
  assert.doesNotMatch(relsXml, /Target="worksheets\/sheet2\.xml"/);
  assert.doesNotMatch(contentTypesXml, /PartName="\/xl\/worksheets\/sheet2\.xml"/);
  assert.match(appXml, /<vt:i4>1<\/vt:i4>/);
  assert.match(appXml, /<vt:lpstr>Sheet1<\/vt:lpstr>/);
  assert.doesNotMatch(appXml, /Sheet2/);
});

test("workbook sheet visibility APIs read and write hidden states", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    replaceEntryText(
      withSecondSheet(
        await loadFixtureEntries(fixtureDir),
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
</worksheet>`,
      ),
      "xl/workbook.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="Sheet2" sheetId="2" r:id="rId3" state="hidden"/>
  </sheets>
</workbook>`,
    ),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);

  assert.equal(workbook.getSheetVisibility("Sheet1"), "visible");
  assert.equal(workbook.getSheetVisibility("Sheet2"), "hidden");

  workbook.setSheetVisibility("Sheet2", "veryHidden");
  assert.equal(workbook.getSheetVisibility("Sheet2"), "veryHidden");

  assert.throws(
    () => workbook.setSheetVisibility("Sheet1", "hidden"),
    /Workbook must contain at least one visible sheet/,
  );

  workbook.setSheetVisibility("Sheet2", "visible");
  workbook.setSheetVisibility("Sheet1", "hidden");

  const workbookXml = entryText(workbook.toEntries(), "xl/workbook.xml");
  assert.match(workbookXml, /<sheet name="Sheet1" sheetId="1" r:id="rId1" state="hidden"\/>/);
  assert.match(workbookXml, /<sheet name="Sheet2" sheetId="2" r:id="rId3"\/>/);
  assert.equal(workbook.getSheetVisibility("Sheet1"), "hidden");
  assert.equal(workbook.getSheetVisibility("Sheet2"), "visible");
});

test("sheet rename updates workbook metadata, formulas, and hyperlink locations", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    replaceEntryText(
      replaceEntryText(
        withSecondSheet(
          await loadFixtureEntries(fixtureDir),
          `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><f>Sheet1!A1</f><v>1</v></c>
    </row>
  </sheetData>
  <hyperlinks><hyperlink ref="A1" location="#Sheet1!A1"/></hyperlinks>
</worksheet>`,
        ),
        "xl/workbook.xml",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="Sheet2" sheetId="2" r:id="rId3"/>
  </sheets>
  <definedNames>
    <definedName name="ExternalRef">Sheet1!$A$1</definedName>
  </definedNames>
</workbook>`,
      ),
      "docProps/app.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>xlsx-ts</Application>
  <HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>2</vt:i4></vt:variant></vt:vector></HeadingPairs>
  <TitlesOfParts><vt:vector size="2" baseType="lpstr"><vt:lpstr>Sheet1</vt:lpstr><vt:lpstr>Sheet2</vt:lpstr></vt:vector></TitlesOfParts>
</Properties>`,
    ),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><f>Sheet1!A1</f><v>1</v></c>
    </row>
  </sheetData>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.rename("Data Set");

  assert.equal(sheet.name, "Data Set");
  assert.deepEqual(workbook.getSheets().map((candidate) => candidate.name), ["Data Set", "Sheet2"]);
  assert.equal(workbook.getSheet("Sheet2").getFormula("A1"), "'Data Set'!A1");

  const workbookXml = entryText(workbook.toEntries(), "xl/workbook.xml");
  const appXml = entryText(workbook.toEntries(), "docProps/app.xml");
  const sheet1Xml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  const sheet2Xml = entryText(workbook.toEntries(), "xl/worksheets/sheet2.xml");

  assert.match(workbookXml, /<sheet name="Data Set" sheetId="1" r:id="rId1"\/>/);
  assert.match(workbookXml, /<definedName name="ExternalRef">'Data Set'!\$A\$1<\/definedName>/);
  assert.match(sheet1Xml, /<c r="A1"><f>'Data Set'!A1<\/f><v>1<\/v><\/c>/);
  assert.match(sheet2Xml, /<c r="A1"><f>'Data Set'!A1<\/f><v>1<\/v><\/c>/);
  assert.match(sheet2Xml, /<hyperlinks><hyperlink ref="A1" location="#'Data Set'!A1"\/><\/hyperlinks>/);
  assert.match(appXml, /<vt:lpstr>Data Set<\/vt:lpstr><vt:lpstr>Sheet2<\/vt:lpstr>/);
});

test("sheet hyperlink APIs read, write, replace, and delete hyperlinks", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    withSecondSheet(
      await loadFixtureEntries(fixtureDir),
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
</worksheet>`,
    ),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><is><t>Old</t></is></c></row>
    <row r="2"><c r="B2"><v>1</v></c></row>
  </sheetData>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  sheet.setHyperlink("A1", "https://example.com", { text: "Open", tooltip: "Go" });
  sheet.setHyperlink("B2", "#Sheet2!A1");

  assert.deepEqual(sheet.getHyperlinks(), [
    { address: "A1", target: "https://example.com", tooltip: "Go", type: "external" },
    { address: "B2", target: "#Sheet2!A1", tooltip: null, type: "internal" },
  ]);
  assert.equal(sheet.getCell("A1"), "Open");

  let sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  let relsXml = entryText(workbook.toEntries(), "xl/worksheets/_rels/sheet1.xml.rels");
  assert.match(
    sheetXml,
    /<hyperlinks><hyperlink ref="A1" r:id="rId1" tooltip="Go"\/><hyperlink ref="B2" location="#Sheet2!A1"\/><\/hyperlinks>/,
  );
  assert.match(
    relsXml,
    /<Relationship Id="rId1" Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/hyperlink" Target="https:\/\/example\.com" TargetMode="External"\/>/,
  );

  sheet.setHyperlink("A1", "#Sheet2!B3");

  assert.deepEqual(sheet.getHyperlinks(), [
    { address: "A1", target: "#Sheet2!B3", tooltip: null, type: "internal" },
    { address: "B2", target: "#Sheet2!A1", tooltip: null, type: "internal" },
  ]);

  sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  relsXml = entryText(workbook.toEntries(), "xl/worksheets/_rels/sheet1.xml.rels");
  assert.match(
    sheetXml,
    /<hyperlinks><hyperlink ref="A1" location="#Sheet2!B3"\/><hyperlink ref="B2" location="#Sheet2!A1"\/><\/hyperlinks>/,
  );
  assert.doesNotMatch(relsXml, /relationships\/hyperlink/);

  sheet.removeHyperlink("A1");

  assert.deepEqual(sheet.getHyperlinks(), [
    { address: "B2", target: "#Sheet2!A1", tooltip: null, type: "internal" },
  ]);

  sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<hyperlinks><hyperlink ref="B2" location="#Sheet2!A1"\/><\/hyperlinks>/);

  sheet.removeHyperlink("B2");

  sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.doesNotMatch(sheetXml, /<hyperlinks>/);
  assert.deepEqual(sheet.getHyperlinks(), []);
});

test("sheet autoFilter APIs read, write, shift, and remove filters", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c><c r="C1"><v>3</v></c></row>
    <row r="2"><c r="A2"><v>4</v></c><c r="B2"><v>5</v></c><c r="C2"><v>6</v></c></row>
  </sheetData>
  <mergeCells count="1"><mergeCell ref="E1:F1"/></mergeCells>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);
  const sheet = workbook.getSheet("Sheet1");

  assert.equal(sheet.getAutoFilter(), null);

  sheet.setAutoFilter("A1:C2");
  assert.equal(sheet.getAutoFilter(), "A1:C2");

  let sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(
    sheetXml,
    /<\/sheetData>\s*<autoFilter ref="A1:C2"\/><mergeCells count="1"><mergeCell ref="E1:F1"\/><\/mergeCells>/,
  );

  sheet.insertColumn("B");
  assert.equal(sheet.getAutoFilter(), "A1:D2");

  sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<autoFilter ref="A1:D2"\/>/);

  workbook.writeEntryText(
    "xl/worksheets/sheet1.xml",
    sheetXml.replace(
      /<autoFilter ref="A1:D2"\/>/,
      `<autoFilter ref="A1:D2"/><sortState ref="A2:D2"/>`,
    ),
  );

  assert.equal(sheet.getAutoFilter(), "A1:D2");

  sheet.removeAutoFilter();

  sheetXml = entryText(workbook.toEntries(), "xl/worksheets/sheet1.xml");
  assert.equal(sheet.getAutoFilter(), null);
  assert.doesNotMatch(sheetXml, /<autoFilter\b/);
  assert.doesNotMatch(sheetXml, /<sortState\b/);
});

test("workbook defined name APIs read, write, and delete global and local names", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    replaceEntryText(
      await loadFixtureEntries(fixtureDir),
      "xl/workbook.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
  <definedNames>
    <definedName name="GlobalValue">Sheet1!$A$1</definedName>
    <definedName name="LocalValue" localSheetId="0">$B$2</definedName>
  </definedNames>
</workbook>`,
    ),
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
    <row r="2"><c r="B2"><v>2</v></c></row>
  </sheetData>
</worksheet>`,
  );
  const workbook = Workbook.fromEntries(entries);

  assert.deepEqual(workbook.getDefinedNames(), [
    { hidden: false, name: "GlobalValue", scope: null, value: "Sheet1!$A$1" },
    { hidden: false, name: "LocalValue", scope: "Sheet1", value: "$B$2" },
  ]);
  assert.equal(workbook.getDefinedName("GlobalValue"), "Sheet1!$A$1");
  assert.equal(workbook.getDefinedName("LocalValue", "Sheet1"), "$B$2");

  workbook.setDefinedName("GlobalValue", "Sheet1!$C$3");
  workbook.setDefinedName("NewLocal", "$D$4", { scope: "Sheet1" });

  let workbookXml = entryText(workbook.toEntries(), "xl/workbook.xml");
  assert.match(workbookXml, /<definedName name="GlobalValue">Sheet1!\$C\$3<\/definedName>/);
  assert.match(workbookXml, /<definedName name="NewLocal" localSheetId="0">\$D\$4<\/definedName>/);

  workbook.deleteDefinedName("LocalValue", "Sheet1");
  workbook.deleteDefinedName("GlobalValue");

  workbookXml = entryText(workbook.toEntries(), "xl/workbook.xml");
  assert.doesNotMatch(workbookXml, /LocalValue/);
  assert.doesNotMatch(workbookXml, /GlobalValue/);
  assert.match(workbookXml, /<definedName name="NewLocal" localSheetId="0">\$D\$4<\/definedName>/);
  assert.deepEqual(workbook.getDefinedNames(), [
    { hidden: false, name: "NewLocal", scope: "Sheet1", value: "$D$4" },
  ]);
});

test("deleting the last defined name removes the definedNames container", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = replaceEntryText(
    await loadFixtureEntries(fixtureDir),
    "xl/workbook.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
  <definedNames>
    <definedName name="OnlyOne">Sheet1!$A$1</definedName>
  </definedNames>
</workbook>`,
  );
  const workbook = Workbook.fromEntries(entries);

  workbook.deleteDefinedName("OnlyOne");

  const workbookXml = entryText(workbook.toEntries(), "xl/workbook.xml");
  assert.doesNotMatch(workbookXml, /<definedNames>/);
  assert.deepEqual(workbook.getDefinedNames(), []);
});

async function loadFixtureEntries(rootDirectory: string): Promise<Array<{ path: string; data: Uint8Array }>> {
  const entries: Array<{ path: string; data: Uint8Array }> = [];
  const stack = [rootDirectory];

  while (stack.length > 0) {
    const current = stack.pop();
    if (!current) {
      continue;
    }

    const names = await readdir(current);

    for (const name of names) {
      const absolutePath = join(current, name);
      const info = await stat(absolutePath);

      if (info.isDirectory()) {
        stack.push(absolutePath);
        continue;
      }

      const relativePath = absolutePath.slice(rootDirectory.length + 1).replaceAll("\\", "/");
      entries.push({
        path: relativePath,
        data: await readFile(absolutePath),
      });
    }
  }

  entries.sort((left, right) => left.path.localeCompare(right.path));
  return entries;
}

function toEntryMap(
  entries: Array<{ path: string; data: Uint8Array }>,
): Map<string, string> {
  return new Map(entries.map((entry) => [entry.path, Buffer.from(entry.data).toString("utf8")]));
}

function assertEntryMapsEqual(expected: Map<string, string>, actual: Map<string, string>): void {
  assert.deepEqual([...actual.keys()].sort(), [...expected.keys()].sort());

  for (const [path, text] of expected) {
    assert.equal(actual.get(path), text, `content mismatch for ${path}`);
  }
}

function entryText(entries: Array<{ path: string; data: Uint8Array }>, path: string): string {
  const entry = entries.find((candidate) => candidate.path === path);
  if (!entry) {
    throw new Error(`Missing entry: ${path}`);
  }

  return Buffer.from(entry.data).toString("utf8");
}

function replaceEntryText(
  entries: Array<{ path: string; data: Uint8Array }>,
  path: string,
  text: string,
): Array<{ path: string; data: Uint8Array }> {
  const encoder = new TextEncoder();
  let replaced = false;
  const nextEntries = entries.map((entry) => {
    if (entry.path !== path) {
      return entry;
    }

    replaced = true;
    return {
      path,
      data: encoder.encode(text),
    };
  });

  if (!replaced) {
    throw new Error(`Missing entry for replacement: ${path}`);
  }

  return nextEntries;
}

function withSecondSheet(
  entries: Array<{ path: string; data: Uint8Array }>,
  sheetXml: string,
): Array<{ path: string; data: Uint8Array }> {
  const encoder = new TextEncoder();

  return [
    ...replaceEntryText(
      replaceEntryText(
        replaceEntryText(
          entries,
          "xl/workbook.xml",
          `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="Sheet2" sheetId="2" r:id="rId3"/>
  </sheets>
</workbook>`,
        ),
        "xl/_rels/workbook.xml.rels",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>`,
      ),
      "[Content_Types].xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
    ),
    {
      path: "xl/worksheets/sheet2.xml",
      data: encoder.encode(sheetXml),
    },
  ].sort((left, right) => left.path.localeCompare(right.path));
}

function withSheetTable(
  entries: Array<{ path: string; data: Uint8Array }>,
  tableXml: string,
): Array<{ path: string; data: Uint8Array }> {
  const encoder = new TextEncoder();

  return [
    ...replaceEntryText(
      entries,
      "[Content_Types].xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>
</Types>`,
    ),
    {
      path: "xl/worksheets/_rels/sheet1.xml.rels",
      data: encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdTable1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
</Relationships>`),
    },
    {
      path: "xl/tables/table1.xml",
      data: encoder.encode(tableXml),
    },
  ].sort((left, right) => left.path.localeCompare(right.path));
}
