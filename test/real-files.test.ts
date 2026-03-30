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

test("openpyxl sample opens with stable workbook structure and roundtrips cleanly", async () => {
  const workbook = await Workbook.open(resolve("res/producers/openpyxl-sample.xlsx"));
  const dataSheet = workbook.getSheet("Data");

  assert.equal(workbook.listEntries().length, 11);
  assert.deepEqual(
    workbook.getSheets().map((sheet) => sheet.name),
    ["Data", "Links"],
  );
  assert.deepEqual(workbook.getDefinedNames(), [
    {
      hidden: false,
      name: "Scores",
      scope: null,
      value: "Data!$A$1:$C$3",
    },
    {
      hidden: true,
      name: "_xlnm._FilterDatabase",
      scope: "Data",
      value: "'Data'!$A$1:$C$3",
    },
  ]);
  assert.equal(dataSheet.getUsedRange(), "A1:D3");
  assert.equal(dataSheet.getAutoFilter(), "A1:C3");
  assert.deepEqual(dataSheet.getFreezePane(), {
    columnCount: 1,
    rowCount: 1,
    topLeftCell: "B2",
    activePane: "bottomRight",
  });
  assert.deepEqual(dataSheet.getMergedRanges(), ["D1:E1"]);
  assert.equal(dataSheet.getFormula("C2"), "B2*2");
  assert.equal(dataSheet.getFormula("C3"), "B3*2");
  assert.deepEqual(dataSheet.getDataValidations(), [
    {
      allowBlank: true,
      error: null,
      errorStyle: null,
      errorTitle: null,
      formula1: "0",
      formula2: "100",
      imeMode: null,
      operator: "between",
      prompt: null,
      promptTitle: null,
      range: "B2:B3",
      showDropDown: false,
      showErrorMessage: false,
      showInputMessage: false,
      type: "whole",
    },
  ]);
  assert.deepEqual(dataSheet.getHyperlinks(), [
    {
      address: "A2",
      target: "#Links!A1",
      tooltip: null,
      type: "external",
    },
  ]);

  const result = await validateRoundtripFile(resolve("res/producers/openpyxl-sample.xlsx"));
  assert.equal(result.ok, true);
  assert.equal(result.entries, 11);
  assert.deepEqual(result.diffs, []);
});

test("xlsxwriter sample opens with stable workbook structure and roundtrips cleanly", async () => {
  const workbook = await Workbook.open(resolve("res/producers/xlsxwriter-sample.xlsx"));
  const dataSheet = workbook.getSheet("Data");

  assert.equal(workbook.listEntries().length, 11);
  assert.deepEqual(
    workbook.getSheets().map((sheet) => sheet.name),
    ["Data", "Links"],
  );
  assert.deepEqual(workbook.getDefinedNames(), [
    {
      hidden: true,
      name: "_xlnm._FilterDatabase",
      scope: "Data",
      value: "Data!$A$1:$C$3",
    },
    {
      hidden: false,
      name: "Scores",
      scope: null,
      value: "Data!$A$1:$C$3",
    },
  ]);
  assert.equal(dataSheet.getUsedRange(), "A1:E3");
  assert.equal(dataSheet.getAutoFilter(), "A1:C3");
  assert.deepEqual(dataSheet.getFreezePane(), {
    columnCount: 1,
    rowCount: 1,
    topLeftCell: "B2",
    activePane: "bottomRight",
  });
  assert.deepEqual(dataSheet.getMergedRanges(), ["D1:E1"]);
  assert.equal(dataSheet.getFormula("C2"), "B2*2");
  assert.equal(dataSheet.getFormula("C3"), "B3*2");
  assert.deepEqual(dataSheet.getDataValidations(), [
    {
      allowBlank: true,
      error: null,
      errorStyle: null,
      errorTitle: null,
      formula1: "0",
      formula2: "100",
      imeMode: null,
      operator: null,
      prompt: null,
      promptTitle: null,
      range: "B2:B3",
      showDropDown: null,
      showErrorMessage: true,
      showInputMessage: true,
      type: "whole",
    },
  ]);
  assert.deepEqual(dataSheet.getHyperlinks(), [
    {
      address: "A2",
      target: "Links!A1",
      tooltip: null,
      type: "internal",
    },
  ]);

  const result = await validateRoundtripFile(resolve("res/producers/xlsxwriter-sample.xlsx"));
  assert.equal(result.ok, true);
  assert.equal(result.entries, 11);
  assert.deepEqual(result.diffs, []);
});
