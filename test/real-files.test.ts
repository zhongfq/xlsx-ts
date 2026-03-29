import test from "node:test";
import assert from "node:assert/strict";
import { resolve } from "node:path";

import { Workbook, validateRoundtripFile } from "../src/index.ts";

test("task.xlsx exposes stable workbook structure", async () => {
  const workbook = await Workbook.open(resolve("res/task.xlsx"));

  assert.equal(workbook.listEntries().length, 39);
  assert.equal(workbook.getActiveSheet().name, "define");
  assert.deepEqual(workbook.getDefinedNames(), [
    {
      hidden: true,
      name: "_xlnm._FilterDatabase",
      scope: "branch",
      value: "branch!$F$1:$F$16",
    },
    {
      hidden: true,
      name: "_xlnm._FilterDatabase",
      scope: "main",
      value: "main!$G$1:$G$17",
    },
  ]);
  assert.deepEqual(
    workbook.getSheets().map((sheet) => ({
      name: sheet.name,
      rowCount: sheet.rowCount,
      columnCount: sheet.columnCount,
      usedRange: sheet.getUsedRange(),
    })),
    [
      { name: "define", rowCount: 24, columnCount: 9, usedRange: "A1:I24" },
      { name: "conf", rowCount: 11, columnCount: 5, usedRange: "A1:E11" },
      { name: "main", rowCount: 17, columnCount: 24, usedRange: "A1:X17" },
      { name: "branch", rowCount: 16, columnCount: 24, usedRange: "A1:X16" },
      { name: "weekly", rowCount: 19, columnCount: 9, usedRange: "A1:I19" },
      { name: "events", rowCount: 19, columnCount: 10, usedRange: "A1:J19" },
      { name: "exchange", rowCount: 9, columnCount: 12, usedRange: "A1:L9" },
    ],
  );
});

test("task.xlsx roundtrips without entry diffs", async () => {
  const result = await validateRoundtripFile(resolve("res/task.xlsx"));

  assert.equal(result.ok, true);
  assert.equal(result.entries, 39);
  assert.deepEqual(result.diffs, []);
});

test("monster.xlsx opens with stable workbook metadata and roundtrips cleanly", async () => {
  const workbook = await Workbook.open(resolve("res/monster.xlsx"));

  assert.equal(workbook.listEntries().length, 51);
  assert.equal(workbook.getActiveSheet().name, "pvp_troop");
  assert.equal(workbook.getDefinedNames().length, 3);
  assert.deepEqual(
    workbook.getSheets().map((sheet) => ({
      name: sheet.name,
      rowCount: sheet.rowCount,
      columnCount: sheet.columnCount,
      usedRange: sheet.getUsedRange(),
    })),
    [
      { name: "troop", rowCount: 965, columnCount: 83, usedRange: "A1:CE965" },
      { name: "td_troop", rowCount: 3874, columnCount: 81, usedRange: "A1:CC3874" },
      { name: "td_soldier", rowCount: 4334, columnCount: 71, usedRange: "A1:BS4334" },
      { name: "prop", rowCount: 327, columnCount: 17, usedRange: "A1:Q327" },
      { name: "attr", rowCount: 4725, columnCount: 39, usedRange: "A1:AM4725" },
      { name: "drop", rowCount: 2094, columnCount: 4, usedRange: "A1:D2094" },
      { name: "pvp_troop", rowCount: 1246, columnCount: 59, usedRange: "A1:BG1246" },
      { name: "scenario_troop", rowCount: 15, columnCount: 79, usedRange: "A1:CA15" },
      { name: "dungeon_troop", rowCount: 1227, columnCount: 59, usedRange: "A1:BG1227" },
    ],
  );

  const result = await validateRoundtripFile(resolve("res/monster.xlsx"));
  assert.equal(result.ok, true);
  assert.equal(result.entries, 51);
  assert.deepEqual(result.diffs, []);
});
