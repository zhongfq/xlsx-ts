import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, readdir, rm, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";

import { runCli } from "../src/cli.ts";
import { Workbook } from "../src/index.ts";

test("inspect reports workbook structure as JSON", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "xlsx-ts-cli-test-"));

  try {
    const inputPath = await writeFixtureWorkbook(tempRoot);
    const result = await runCliCapture(["inspect", inputPath]);

    assert.equal(result.exitCode, 0);

    const output = JSON.parse(result.stdout);
    assert.equal(output.file, inputPath);
    assert.equal(output.activeSheet, "Sheet1");
    assert.deepEqual(output.definedNames, []);
    assert.deepEqual(output.sheets, [
      {
        columnCount: 1,
        headers: ["Hello"],
        name: "Sheet1",
        rowCount: 1,
        usedRange: "A1",
        visibility: "visible",
      },
    ]);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("set writes a cell value to a new workbook and preserves the style id", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "xlsx-ts-cli-test-"));

  try {
    const inputPath = await writeFixtureWorkbook(tempRoot);
    const outputPath = join(tempRoot, "set-output.xlsx");
    const result = await runCliCapture([
      "set",
      inputPath,
      "--sheet",
      "Sheet1",
      "--cell",
      "A1",
      "--text",
      "World",
      "--output",
      outputPath,
    ]);

    assert.equal(result.exitCode, 0);

    const payload = JSON.parse(result.stdout);
    assert.equal(payload.output, outputPath);
    assert.equal(payload.result.value, "World");
    assert.equal(payload.result.styleId, 1);

    const workbook = await Workbook.open(outputPath);
    const sheet = workbook.getSheet("Sheet1");
    assert.equal(sheet.getCell("A1"), "World");
    assert.equal(sheet.getStyleId("A1"), 1);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("add-sheet creates a new worksheet through the direct CLI command", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "xlsx-ts-cli-test-"));

  try {
    const inputPath = await writeFixtureWorkbook(tempRoot);
    const outputPath = join(tempRoot, "add-sheet-output.xlsx");
    const result = await runCliCapture([
      "add-sheet",
      inputPath,
      "--sheet",
      "Config",
      "--output",
      outputPath,
    ]);

    assert.equal(result.exitCode, 0);

    const payload = JSON.parse(result.stdout);
    assert.deepEqual(payload.sheets, ["Sheet1", "Config"]);

    const workbook = await Workbook.open(outputPath);
    assert.deepEqual(
      workbook.getSheets().map((sheet) => sheet.name),
      ["Sheet1", "Config"],
    );
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("rename-sheet and delete-sheet manage worksheets through direct commands", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "xlsx-ts-cli-test-"));

  try {
    const inputPath = await writeFixtureWorkbook(tempRoot);
    const withExtraSheetPath = join(tempRoot, "with-extra-sheet.xlsx");
    const renamedPath = join(tempRoot, "renamed.xlsx");
    const deletedPath = join(tempRoot, "deleted.xlsx");

    let result = await runCliCapture([
      "add-sheet",
      inputPath,
      "--sheet",
      "Scratch",
      "--output",
      withExtraSheetPath,
    ]);
    assert.equal(result.exitCode, 0);

    result = await runCliCapture([
      "rename-sheet",
      withExtraSheetPath,
      "--from",
      "Sheet1",
      "--to",
      "Config",
      "--output",
      renamedPath,
    ]);
    assert.equal(result.exitCode, 0);

    result = await runCliCapture([
      "delete-sheet",
      renamedPath,
      "--sheet",
      "Scratch",
      "--output",
      deletedPath,
    ]);
    assert.equal(result.exitCode, 0);

    const workbook = await Workbook.open(deletedPath);
    assert.deepEqual(
      workbook.getSheets().map((sheet) => sheet.name),
      ["Config"],
    );
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("record commands manage header-based sheet data through the CLI", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "xlsx-ts-cli-test-"));

  try {
    const inputPath = await writeFixtureWorkbook(tempRoot);
    const headersPath = join(tempRoot, "headers.xlsx");
    const recordsPath = join(tempRoot, "records.xlsx");
    const replacedPath = join(tempRoot, "replaced.xlsx");
    const deletedPath = join(tempRoot, "deleted.xlsx");

    let result = await runCliCapture([
      "set-headers",
      inputPath,
      "--sheet",
      "Sheet1",
      "--headers",
      '["Key","Value"]',
      "--output",
      headersPath,
    ]);
    assert.equal(result.exitCode, 0);

    result = await runCliCapture([
      "add-record",
      headersPath,
      "--sheet",
      "Sheet1",
      "--record",
      '{"Key":"alpha","Value":"1"}',
      "--output",
      recordsPath,
    ]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).records, [{ Key: "alpha", Value: "1" }]);

    result = await runCliCapture([
      "set-records",
      recordsPath,
      "--sheet",
      "Sheet1",
      "--records",
      '[{"Key":"alpha","Value":"10"},{"Key":"beta","Value":"20"}]',
      "--output",
      replacedPath,
    ]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).records, [
      { Key: "alpha", Value: "10" },
      { Key: "beta", Value: "20" },
    ]);

    result = await runCliCapture([
      "delete-record",
      replacedPath,
      "--sheet",
      "Sheet1",
      "--row",
      "2",
      "--output",
      deletedPath,
    ]);
    assert.equal(result.exitCode, 0);

    result = await runCliCapture(["records", deletedPath, "--sheet", "Sheet1"]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).records, [{ Key: "beta", Value: "20" }]);

    const workbook = await Workbook.open(deletedPath);
    assert.deepEqual(workbook.getSheet("Sheet1").getHeaders(), ["Key", "Value"]);
    assert.deepEqual(workbook.getSheet("Sheet1").getRecords(), [{ Key: "beta", Value: "20" }]);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("config-table command group supports high-level config workflows", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "xlsx-ts-cli-test-"));

  try {
    const inputPath = await writeFixtureWorkbook(tempRoot);
    const initializedPath = join(tempRoot, "config-init.xlsx");
    const upsertedPath = join(tempRoot, "config-upserted.xlsx");
    const deletedPath = join(tempRoot, "config-deleted.xlsx");
    const replacedPath = join(tempRoot, "config-replaced.xlsx");

    let result = await runCliCapture([
      "config-table",
      "init",
      inputPath,
      "--sheet",
      "Config",
      "--headers",
      '["Key","Value"]',
      "--output",
      initializedPath,
    ]);
    assert.equal(result.exitCode, 0);

    result = await runCliCapture([
      "config-table",
      "upsert",
      initializedPath,
      "--sheet",
      "Config",
      "--field",
      "Key",
      "--record",
      '{"Key":"timeout","Value":"30"}',
      "--output",
      upsertedPath,
    ]);
    assert.equal(result.exitCode, 0);

    result = await runCliCapture([
      "config-table",
      "upsert",
      upsertedPath,
      "--sheet",
      "Config",
      "--field",
      "Key",
      "--record",
      '{"Key":"timeout","Value":"60"}',
      "--in-place",
    ]);
    assert.equal(result.exitCode, 0);

    result = await runCliCapture([
      "config-table",
      "get",
      upsertedPath,
      "--sheet",
      "Config",
      "--field",
      "Key",
      "--text",
      "timeout",
    ]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).record, {
      record: { Key: "timeout", Value: "60" },
      row: 2,
    });

    result = await runCliCapture([
      "config-table",
      "delete",
      upsertedPath,
      "--sheet",
      "Config",
      "--field",
      "Key",
      "--text",
      "timeout",
      "--output",
      deletedPath,
    ]);
    assert.equal(result.exitCode, 0);
    assert.equal(JSON.parse(result.stdout).deleted, true);

    result = await runCliCapture([
      "config-table",
      "replace",
      deletedPath,
      "--sheet",
      "Config",
      "--records",
      '[{"Key":"region","Value":"cn"},{"Key":"retries","Value":"3"}]',
      "--output",
      replacedPath,
    ]);
    assert.equal(result.exitCode, 0);

    result = await runCliCapture([
      "config-table",
      "list",
      replacedPath,
      "--sheet",
      "Config",
    ]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).rows, [
      { row: 2, record: { Key: "region", Value: "cn" } },
      { row: 3, record: { Key: "retries", Value: "3" } },
    ]);

    const workbook = await Workbook.open(replacedPath);
    assert.deepEqual(workbook.getSheet("Config").getHeaders(), ["Key", "Value"]);
    assert.deepEqual(workbook.getSheet("Config").getRecords(), [
      { Key: "region", Value: "cn" },
      { Key: "retries", Value: "3" },
    ]);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("config-table sync imports JSON config objects in replace and upsert modes", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "xlsx-ts-cli-test-"));

  try {
    const inputPath = await writeFixtureWorkbook(tempRoot);
    const replaceJsonPath = join(tempRoot, "replace.json");
    const replaceOutputPath = join(tempRoot, "sync-replace.xlsx");
    const upsertJsonPath = join(tempRoot, "upsert.json");

    await writeFile(
      replaceJsonPath,
      JSON.stringify(
        {
          timeout: "30",
          region: "cn",
        },
        null,
        2,
      ),
    );

    let result = await runCliCapture([
      "config-table",
      "sync",
      inputPath,
      "--sheet",
      "Config",
      "--from-json",
      replaceJsonPath,
      "--output",
      replaceOutputPath,
    ]);
    assert.equal(result.exitCode, 0);
    assert.equal(JSON.parse(result.stdout).mode, "replace");

    await writeFile(
      upsertJsonPath,
      JSON.stringify(
        {
          timeout: "60",
          retries: "3",
        },
        null,
        2,
      ),
    );

    result = await runCliCapture([
      "config-table",
      "sync",
      replaceOutputPath,
      "--sheet",
      "Config",
      "--from-json",
      upsertJsonPath,
      "--mode",
      "upsert",
      "--in-place",
    ]);
    assert.equal(result.exitCode, 0);
    assert.equal(JSON.parse(result.stdout).mode, "upsert");

    result = await runCliCapture([
      "config-table",
      "list",
      replaceOutputPath,
      "--sheet",
      "Config",
    ]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).rows, [
      { row: 2, record: { Key: "timeout", Value: "60" } },
      { row: 3, record: { Key: "region", Value: "cn" } },
      { row: 4, record: { Key: "retries", Value: "3" } },
    ]);

    const workbook = await Workbook.open(replaceOutputPath);
    assert.deepEqual(workbook.getSheet("Config").getHeaders(), ["Key", "Value"]);
    assert.deepEqual(workbook.getSheet("Config").getRecords(), [
      { Key: "timeout", Value: "60" },
      { Key: "region", Value: "cn" },
      { Key: "retries", Value: "3" },
    ]);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("table command group respects explicit data row boundaries", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "xlsx-ts-cli-test-"));

  try {
    const inputPath = await writeStructuredTableWorkbook(tempRoot);
    const upsertedPath = join(tempRoot, "table-upsert.xlsx");
    const syncJsonPath = join(tempRoot, "table-sync.json");
    const syncedPath = join(tempRoot, "table-sync.xlsx");

    let result = await runCliCapture([
      "table",
      "inspect",
      inputPath,
      "--sheet",
      "Sheet1",
      "--header-row",
      "1",
      "--data-start-row",
      "6",
    ]);
    assert.equal(result.exitCode, 0);
    assert.equal(JSON.parse(result.stdout).dataRowCount, 2);

    result = await runCliCapture([
      "table",
      "get",
      inputPath,
      "--sheet",
      "Sheet1",
      "--header-row",
      "1",
      "--data-start-row",
      "6",
      "--key-field",
      "id",
      "--key",
      "1001",
    ]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).row, {
      row: 6,
      record: { id: 1001, name: "Alpha" },
    });

    result = await runCliCapture([
      "table",
      "upsert",
      inputPath,
      "--sheet",
      "Sheet1",
      "--header-row",
      "1",
      "--data-start-row",
      "6",
      "--key-field",
      "id",
      "--record",
      '{"id":1002,"name":"Beta-2"}',
      "--output",
      upsertedPath,
    ]);
    assert.equal(result.exitCode, 0);

    await writeFile(
      syncJsonPath,
      JSON.stringify(
        [
          { id: 1001, name: "Alpha-2" },
          { id: 1003, name: "Gamma" },
        ],
        null,
        2,
      ),
    );

    result = await runCliCapture([
      "table",
      "sync",
      upsertedPath,
      "--sheet",
      "Sheet1",
      "--header-row",
      "1",
      "--data-start-row",
      "6",
      "--key-field",
      "id",
      "--from-json",
      syncJsonPath,
      "--output",
      syncedPath,
    ]);
    assert.equal(result.exitCode, 0);

    result = await runCliCapture([
      "table",
      "list",
      syncedPath,
      "--sheet",
      "Sheet1",
      "--header-row",
      "1",
      "--data-start-row",
      "6",
    ]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).rows, [
      { row: 6, record: { id: 1001, name: "Alpha-2" } },
      { row: 7, record: { id: 1003, name: "Gamma" } },
    ]);

    const workbook = await Workbook.open(syncedPath);
    const sheet = workbook.getSheet("Sheet1");
    assert.deepEqual(sheet.getRow(2), ["int", "string"]);
    assert.deepEqual(sheet.getRow(3), [">>", "client"]);
    assert.deepEqual(sheet.getRow(4), ["!!!", "x"]);
    assert.deepEqual(sheet.getRow(5), ["###", "display"]);
    assert.deepEqual(sheet.getRecord(6, 1), { id: 1001, name: "Alpha-2" });
    assert.deepEqual(sheet.getRecord(7, 1), { id: 1003, name: "Gamma" });
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("table command group supports profile presets for structured sheets", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "xlsx-ts-cli-test-"));

  try {
    const inputPath = await writeStructuredTableWorkbook(tempRoot);
    const profilesPath = join(tempRoot, "table-profiles.json");
    const upsertedPath = join(tempRoot, "profile-upsert.xlsx");

    await writeFile(
      profilesPath,
      JSON.stringify(
        {
          profiles: {
            demo: {
              sheet: "Sheet1",
              headerRow: 1,
              dataStartRow: 6,
              keyFields: ["id"],
            },
          },
        },
        null,
        2,
      ),
    );

    let result = await runCliCapture([
      "table",
      "list",
      inputPath,
      "--profile",
      "demo",
      "--profiles-file",
      profilesPath,
    ]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).rows, [
      { row: 6, record: { id: 1001, name: "Alpha" } },
      { row: 7, record: { id: 1002, name: "Beta" } },
    ]);

    result = await runCliCapture([
      "table",
      "get",
      inputPath,
      "--profile",
      "demo",
      "--profiles-file",
      profilesPath,
      "--key",
      "1002",
    ]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).row, {
      row: 7,
      record: { id: 1002, name: "Beta" },
    });

    result = await runCliCapture([
      "table",
      "upsert",
      inputPath,
      "--profile",
      "demo",
      "--profiles-file",
      profilesPath,
      "--record",
      '{"id":1002,"name":"Beta-2"}',
      "--output",
      upsertedPath,
    ]);
    assert.equal(result.exitCode, 0);

    const payload = JSON.parse(result.stdout);
    assert.deepEqual(payload.keyFields, ["id"]);
    assert.deepEqual(payload.rows, [
      { row: 6, record: { id: 1001, name: "Alpha" } },
      { row: 7, record: { id: 1002, name: "Beta-2" } },
    ]);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("explicit table options override profile values", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "xlsx-ts-cli-test-"));

  try {
    const inputPath = await writeStructuredTableWorkbook(tempRoot);
    const profilesPath = join(tempRoot, "table-profiles.json");

    await writeFile(
      profilesPath,
      JSON.stringify(
        {
          profiles: {
            demo: {
              sheet: "Sheet1",
              headerRow: 1,
              dataStartRow: 7,
              keyFields: ["id"],
            },
          },
        },
        null,
        2,
      ),
    );

    const result = await runCliCapture([
      "table",
      "list",
      inputPath,
      "--profile",
      "demo",
      "--profiles-file",
      profilesPath,
      "--data-start-row",
      "6",
    ]);
    assert.equal(result.exitCode, 0);

    const payload = JSON.parse(result.stdout);
    assert.equal(payload.dataStartRow, 6);
    assert.deepEqual(payload.rows, [
      { row: 6, record: { id: 1001, name: "Alpha" } },
      { row: 7, record: { id: 1002, name: "Beta" } },
    ]);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("table command group supports composite key profiles", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "xlsx-ts-cli-test-"));

  try {
    const inputPath = await writeCompositeStructuredTableWorkbook(tempRoot);
    const profilesPath = join(tempRoot, "table-profiles.json");
    const upsertedPath = join(tempRoot, "composite-profile-upsert.xlsx");

    await writeFile(
      profilesPath,
      JSON.stringify(
        {
          profiles: {
            defineLike: {
              sheet: "Sheet1",
              headerRow: 2,
              dataStartRow: 7,
              keyFields: ["key1", "key2"],
            },
          },
        },
        null,
        2,
      ),
    );

    let result = await runCliCapture([
      "table",
      "get",
      inputPath,
      "--profile",
      "defineLike",
      "--profiles-file",
      profilesPath,
      "--key",
      '{"key1":"TASK_TYPE","key2":"MAIN"}',
    ]);
    assert.equal(result.exitCode, 0);
    assert.deepEqual(JSON.parse(result.stdout).row, {
      row: 7,
      record: {
        id: "-",
        comment: "任务类型",
        key1: "TASK_TYPE",
        key2: "MAIN",
        value_comment: "主线任务",
        value: 1,
        value_type: "int",
        enum: "TaskType",
        enum_option: "true",
      },
    });

    result = await runCliCapture([
      "table",
      "upsert",
      inputPath,
      "--profile",
      "defineLike",
      "--profiles-file",
      profilesPath,
      "--record",
      '{"id":"-","comment":"任务类型","key1":"TASK_TYPE","key2":"MAIN","value_comment":"主线任务-新","value":10,"value_type":"int","enum":"TaskType","enum_option":"true"}',
      "--output",
      upsertedPath,
    ]);
    assert.equal(result.exitCode, 0);

    const workbook = await Workbook.open(upsertedPath);
    assert.deepEqual(workbook.getSheet("Sheet1").getRecord(7, 2), {
      id: "-",
      comment: "任务类型",
      key1: "TASK_TYPE",
      key2: "MAIN",
      value_comment: "主线任务-新",
      value: 10,
      value_type: "int",
      enum: "TaskType",
      enum_option: "true",
    });
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("style commands update formatting and can copy styles through the CLI", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "xlsx-ts-cli-test-"));

  try {
    const inputPath = await writeFixtureWorkbook(tempRoot);
    const withValuePath = join(tempRoot, "with-value.xlsx");
    const formattedPath = join(tempRoot, "formatted.xlsx");

    let result = await runCliCapture([
      "set",
      inputPath,
      "--sheet",
      "Sheet1",
      "--cell",
      "B1",
      "--text",
      "Tail",
      "--output",
      withValuePath,
    ]);
    assert.equal(result.exitCode, 0);

    result = await runCliCapture([
      "set-background-color",
      withValuePath,
      "--sheet",
      "Sheet1",
      "--cell",
      "A1",
      "--color",
      "FFFF0000",
      "--in-place",
    ]);
    assert.equal(result.exitCode, 0);

    result = await runCliCapture([
      "set-number-format",
      withValuePath,
      "--sheet",
      "Sheet1",
      "--cell",
      "A1",
      "--format",
      "0.00%",
      "--in-place",
    ]);
    assert.equal(result.exitCode, 0);

    result = await runCliCapture([
      "copy-style",
      withValuePath,
      "--sheet",
      "Sheet1",
      "--from",
      "A1",
      "--to",
      "B1",
      "--output",
      formattedPath,
    ]);
    assert.equal(result.exitCode, 0);

    const payload = JSON.parse(result.stdout);
    assert.equal(payload.result.backgroundColor, "FFFF0000");
    assert.equal(payload.result.numberFormat, "0.00%");
    assert.equal(payload.result.value, "Tail");

    const workbook = await Workbook.open(formattedPath);
    const sheet = workbook.getSheet("Sheet1");
    assert.equal(sheet.getBackgroundColor("B1"), "FFFF0000");
    assert.equal(sheet.getNumberFormat("B1")?.code, "0.00%");
    assert.equal(sheet.getCell("B1"), "Tail");
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("apply executes structured workbook operations", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "xlsx-ts-cli-test-"));

  try {
    const inputPath = await writeFixtureWorkbook(tempRoot);
    const outputPath = join(tempRoot, "apply-output.xlsx");
    const opsPath = join(tempRoot, "ops.json");

    await writeFile(
      opsPath,
      JSON.stringify(
        {
          actions: [
            { type: "renameSheet", from: "Sheet1", to: "Config" },
            { type: "setCell", sheet: "Config", cell: "A1", value: "Updated" },
            { type: "setDefinedName", name: "Greeting", value: "Config!$A$1" },
            { type: "setActiveSheet", sheet: "Config" },
          ],
        },
        null,
        2,
      ),
    );

    const result = await runCliCapture([
      "apply",
      inputPath,
      "--ops",
      opsPath,
      "--output",
      outputPath,
    ]);

    assert.equal(result.exitCode, 0);

    const payload = JSON.parse(result.stdout);
    assert.equal(payload.actions, 4);
    assert.deepEqual(payload.sheets, ["Config"]);

    const workbook = await Workbook.open(outputPath);
    const sheet = workbook.getSheet("Config");
    assert.equal(sheet.getCell("A1"), "Updated");
    assert.equal(workbook.getDefinedName("Greeting"), "Config!$A$1");
    assert.equal(workbook.getActiveSheet().name, "Config");
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("apply supports worksheet and style operations", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "xlsx-ts-cli-test-"));

  try {
    const inputPath = await writeFixtureWorkbook(tempRoot);
    const outputPath = join(tempRoot, "apply-style-output.xlsx");
    const opsPath = join(tempRoot, "style-ops.json");

    await writeFile(
      opsPath,
      JSON.stringify(
        {
          actions: [
            { type: "addSheet", sheet: "Scratch" },
            { type: "renameSheet", from: "Sheet1", to: "Config" },
            { type: "setCell", sheet: "Config", cell: "B1", value: "Tail" },
            { type: "setBackgroundColor", sheet: "Config", cell: "A1", color: "FF00FF00" },
            { type: "setNumberFormat", sheet: "Config", cell: "A1", formatCode: "0.00%" },
            { type: "copyStyle", sheet: "Config", from: "A1", to: "B1" },
            { type: "deleteSheet", sheet: "Scratch" },
          ],
        },
        null,
        2,
      ),
    );

    const result = await runCliCapture([
      "apply",
      inputPath,
      "--ops",
      opsPath,
      "--output",
      outputPath,
    ]);

    assert.equal(result.exitCode, 0);

    const workbook = await Workbook.open(outputPath);
    assert.deepEqual(workbook.getSheets().map((sheet) => sheet.name), ["Config"]);
    assert.equal(workbook.getSheet("Config").getBackgroundColor("B1"), "FF00FF00");
    assert.equal(workbook.getSheet("Config").getNumberFormat("B1")?.code, "0.00%");
    assert.equal(workbook.getSheet("Config").getCell("B1"), "Tail");
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("apply supports record and header operations", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "xlsx-ts-cli-test-"));

  try {
    const inputPath = await writeFixtureWorkbook(tempRoot);
    const outputPath = join(tempRoot, "apply-records-output.xlsx");
    const opsPath = join(tempRoot, "records-ops.json");

    await writeFile(
      opsPath,
      JSON.stringify(
        {
          actions: [
            { type: "setHeaders", sheet: "Sheet1", headers: ["Key", "Value"] },
            { type: "addRecords", sheet: "Sheet1", records: [{ Key: "a", Value: "1" }, { Key: "b", Value: "2" }] },
            { type: "deleteRecord", sheet: "Sheet1", row: 2 },
          ],
        },
        null,
        2,
      ),
    );

    const result = await runCliCapture([
      "apply",
      inputPath,
      "--ops",
      opsPath,
      "--output",
      outputPath,
    ]);

    assert.equal(result.exitCode, 0);

    const workbook = await Workbook.open(outputPath);
    assert.deepEqual(workbook.getSheet("Sheet1").getHeaders(), ["Key", "Value"]);
    assert.deepEqual(workbook.getSheet("Sheet1").getRecords(), [{ Key: "b", Value: "2" }]);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

test("validate returns a successful roundtrip result for the fixture workbook", async () => {
  const tempRoot = await mkdtemp(join(tmpdir(), "xlsx-ts-cli-test-"));

  try {
    const inputPath = await writeFixtureWorkbook(tempRoot);
    const result = await runCliCapture(["validate", inputPath]);

    assert.equal(result.exitCode, 0);

    const payload = JSON.parse(result.stdout);
    assert.equal(payload.input, inputPath);
    assert.equal(payload.ok, true);
    assert.deepEqual(payload.diffs, []);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

async function runCliCapture(argv: string[]): Promise<{
  exitCode: number;
  stderr: string;
  stdout: string;
}> {
  let stdout = "";
  let stderr = "";
  const exitCode = await runCli(argv, {
    stderr: (chunk) => {
      stderr += chunk;
    },
    stdout: (chunk) => {
      stdout += chunk;
    },
  });

  return { exitCode, stderr, stdout };
}

async function writeFixtureWorkbook(rootDirectory: string): Promise<string> {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const entries = await loadFixtureEntries(fixtureDir);
  const workbook = Workbook.fromEntries(entries);
  const inputPath = join(rootDirectory, "input.xlsx");
  await workbook.save(inputPath);
  return inputPath;
}

async function writeStructuredTableWorkbook(rootDirectory: string): Promise<string> {
  const inputPath = await writeFixtureWorkbook(rootDirectory);
  const workbook = await Workbook.open(inputPath);
  const sheet = workbook.getSheet("Sheet1");

  sheet.setRow(1, ["id", "name"]);
  sheet.setRow(2, ["int", "string"]);
  sheet.setRow(3, [">>", "client"]);
  sheet.setRow(4, ["!!!", "x"]);
  sheet.setRow(5, ["###", "display"]);
  sheet.setRow(6, [1001, "Alpha"]);
  sheet.setRow(7, [1002, "Beta"]);

  await workbook.save(inputPath);
  return inputPath;
}

async function writeCompositeStructuredTableWorkbook(rootDirectory: string): Promise<string> {
  const inputPath = await writeFixtureWorkbook(rootDirectory);
  const workbook = await Workbook.open(inputPath);
  const sheet = workbook.getSheet("Sheet1");

  sheet.setRow(1, ["@define"]);
  sheet.setRow(2, [
    "id",
    "comment",
    "key1",
    "key2",
    "value_comment",
    "value",
    "value_type",
    "enum",
    "enum_option",
  ]);
  sheet.setRow(3, ["auto", "string?", "string", "string?", "string?", "@value_type", "string", "string?", "bool?"]);
  sheet.setRow(4, [">>", "client|server", null, null, null, null, null, null, null]);
  sheet.setRow(5, ["!!!", "x", "x", "x", "x", "x", "x", "x", "x"]);
  sheet.setRow(6, ["###", "注释", null, null, "注释", null, null, null, null]);
  sheet.setRow(7, ["-", "任务类型", "TASK_TYPE", "MAIN", "主线任务", 1, "int", "TaskType", "true"]);
  sheet.setRow(8, ["-", null, "TASK_TYPE", "BRANCH", "支线任务", 2, "int", null, null]);

  await workbook.save(inputPath);
  return inputPath;
}

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
