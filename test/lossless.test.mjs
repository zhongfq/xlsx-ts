import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";

import { Workbook } from "../dist/src/index.js";

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

async function loadFixtureEntries(rootDirectory) {
  const { readdir, stat } = await import("node:fs/promises");
  const entries = [];
  const stack = [rootDirectory];

  while (stack.length > 0) {
    const current = stack.pop();
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

function toEntryMap(entries) {
  return new Map(entries.map((entry) => [entry.path, Buffer.from(entry.data).toString("utf8")]));
}

function assertEntryMapsEqual(expected, actual) {
  assert.deepEqual([...actual.keys()].sort(), [...expected.keys()].sort());

  for (const [path, text] of expected) {
    assert.equal(actual.get(path), text, `content mismatch for ${path}`);
  }
}

function entryText(entries, path) {
  const entry = entries.find((candidate) => candidate.path === path);
  if (!entry) {
    throw new Error(`Missing entry: ${path}`);
  }

  return Buffer.from(entry.data).toString("utf8");
}
