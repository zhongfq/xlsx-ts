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
