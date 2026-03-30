import { writeFile } from "node:fs/promises";
import { resolve } from "node:path";

import { Command, InvalidArgumentError } from "commander";

import {
  formatError,
  parseJsonCellRecord,
  parseJsonCellRecordArray,
  parseJsonStringArray,
  resolveMatchValue,
  resolveUpsertMatchValue,
  writeJson,
} from "./cli-json.js";
import type { Writer } from "./cli-json.js";
import {
  findConfigTableRow,
  findStructuredTableRow,
  generateTableProfiles,
  getConfigTableRecord,
  getConfigTableRows,
  getStructuredTableRecord,
  getStructuredTableRows,
  inspectTable,
  parseTableKey,
  pickKeyRecord,
  readConfigTableSyncInput,
  resolveConfigTableHeaders,
  resolveTableCommandContext,
  resolveTableKeyFields,
  writeStructuredTableRecord,
  writeStructuredTableRecords,
} from "./cli-table.js";
import type { TableCommandContext } from "./cli-table.js";
import { Workbook } from "./workbook.js";

type ConfigTableSyncMode = "replace" | "upsert";

export interface TableCommandIo {
  cwd: string;
  stdout: Writer;
}

export interface TableCommandHelpers {
  parsePositiveInteger: (value: string) => number;
  resolveOutputPath: (
    inputPath: string,
    options: {
      inPlace: boolean;
      output?: string;
    },
  ) => string;
}

export function registerTableCommands(
  program: Command,
  io: TableCommandIo,
  helpers: TableCommandHelpers,
): void {
  const { parsePositiveInteger, resolveOutputPath } = helpers;

  const configTable = program
    .command("config-table")
    .description("High-level workflow for header-based config sheets");

  configTable
    .command("init")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--headers <json>", "JSON array of header strings")
    .option("--header-row <row>", "target header row", parsePositiveInteger, 1)
    .option("--start-column <column>", "1-based start column", parsePositiveInteger, 1)
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          headerRow: number;
          headers: string;
          inPlace?: boolean;
          output?: string;
          sheet: string;
          startColumn: number;
        },
      ) => {
        const inputPath = resolveFrom(io.cwd, file);
        const outputPath = resolveOutputPath(inputPath, {
          inPlace: options.inPlace === true,
          output: options.output ? resolveFrom(io.cwd, options.output) : undefined,
        });
        const workbook = await Workbook.open(inputPath);
        const sheet = getOrCreateSheet(workbook, options.sheet);
        const headers = parseJsonStringArray(options.headers, "--headers");
        sheet.setHeaders(headers, options.headerRow, options.startColumn);
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "configTable.init",
          headers: sheet.getHeaders(options.headerRow),
          input: inputPath,
          output: outputPath,
          sheet: options.sheet,
        });
      },
    );

  configTable
    .command("list")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .option("--header-row <row>", "header row used for record mapping", parsePositiveInteger, 1)
    .action(async (file: string, options: { headerRow: number; sheet: string }) => {
      const result = await getConfigTableRows(resolveFrom(io.cwd, file), options.sheet, options.headerRow);
      writeJson(io.stdout, result);
    });

  configTable
    .command("get")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--field <name>", "header field used as the lookup key")
    .option("--value <json>", "JSON scalar value to match")
    .option("--text <value>", "plain string value to match")
    .option("--header-row <row>", "header row used for record mapping", parsePositiveInteger, 1)
    .action(
      async (
        file: string,
        options: {
          field: string;
          headerRow: number;
          sheet: string;
          text?: string;
          value?: string;
        },
      ) => {
        const matchValue = resolveMatchValue(options.value, options.text);
        const result = await getConfigTableRecord(
          resolveFrom(io.cwd, file),
          options.sheet,
          options.headerRow,
          options.field,
          matchValue,
        );
        writeJson(io.stdout, result);
      },
    );

  configTable
    .command("upsert")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--field <name>", "header field used as the match key")
    .requiredOption("--record <json>", "JSON object keyed by header names")
    .option("--match-value <json>", "JSON scalar override for the lookup value")
    .option("--match-text <value>", "plain string override for the lookup value")
    .option("--header-row <row>", "header row used for record mapping", parsePositiveInteger, 1)
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          field: string;
          headerRow: number;
          inPlace?: boolean;
          matchText?: string;
          matchValue?: string;
          output?: string;
          record: string;
          sheet: string;
        },
      ) => {
        const inputPath = resolveFrom(io.cwd, file);
        const outputPath = resolveOutputPath(inputPath, {
          inPlace: options.inPlace === true,
          output: options.output ? resolveFrom(io.cwd, options.output) : undefined,
        });
        const workbook = await Workbook.open(inputPath);
        const sheet = workbook.getSheet(options.sheet);
        const record = parseJsonCellRecord(options.record, "--record");
        const matchValue = resolveUpsertMatchValue(
          record,
          options.field,
          options.matchValue,
          options.matchText,
        );
        const matchedRow = findConfigTableRow(sheet, options.headerRow, options.field, matchValue)?.row ?? null;

        if (matchedRow === null) {
          sheet.addRecord(record, options.headerRow);
        } else {
          sheet.setRecord(matchedRow, record, options.headerRow);
        }

        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "configTable.upsert",
          input: inputPath,
          matchField: options.field,
          matchValue,
          output: outputPath,
          record,
          records: (await getConfigTableRows(outputPath, options.sheet, options.headerRow)).rows,
          row: findConfigTableRow(workbook.getSheet(options.sheet), options.headerRow, options.field, matchValue)?.row ?? null,
          sheet: options.sheet,
        });
      },
    );

  configTable
    .command("delete")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--field <name>", "header field used as the lookup key")
    .option("--value <json>", "JSON scalar value to match")
    .option("--text <value>", "plain string value to match")
    .option("--header-row <row>", "header row used for record mapping", parsePositiveInteger, 1)
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          field: string;
          headerRow: number;
          inPlace?: boolean;
          output?: string;
          sheet: string;
          text?: string;
          value?: string;
        },
      ) => {
        const matchValue = resolveMatchValue(options.value, options.text);
        const inputPath = resolveFrom(io.cwd, file);
        const outputPath = resolveOutputPath(inputPath, {
          inPlace: options.inPlace === true,
          output: options.output ? resolveFrom(io.cwd, options.output) : undefined,
        });
        const workbook = await Workbook.open(inputPath);
        const sheet = workbook.getSheet(options.sheet);
        const row = findConfigTableRow(sheet, options.headerRow, options.field, matchValue)?.row ?? null;

        if (row !== null) {
          sheet.deleteRecord(row, options.headerRow);
        }

        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "configTable.delete",
          deleted: row !== null,
          input: inputPath,
          matchField: options.field,
          matchValue,
          output: outputPath,
          records: (await getConfigTableRows(outputPath, options.sheet, options.headerRow)).rows,
          row,
          sheet: options.sheet,
        });
      },
    );

  configTable
    .command("replace")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--records <json>", "JSON array of record objects")
    .option("--header-row <row>", "header row used for record mapping", parsePositiveInteger, 1)
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          headerRow: number;
          inPlace?: boolean;
          output?: string;
          records: string;
          sheet: string;
        },
      ) => {
        const inputPath = resolveFrom(io.cwd, file);
        const outputPath = resolveOutputPath(inputPath, {
          inPlace: options.inPlace === true,
          output: options.output ? resolveFrom(io.cwd, options.output) : undefined,
        });
        const workbook = await Workbook.open(inputPath);
        const sheet = workbook.getSheet(options.sheet);
        const records = parseJsonCellRecordArray(options.records, "--records");
        sheet.setRecords(records, options.headerRow);
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "configTable.replace",
          input: inputPath,
          output: outputPath,
          records: (await getConfigTableRows(outputPath, options.sheet, options.headerRow)).rows,
          sheet: options.sheet,
        });
      },
    );

  configTable
    .command("sync")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--from-json <file>", "JSON file containing records or a config object")
    .option("--field <name>", "header field used as the key when normalizing config objects", "Key")
    .option("--value-field <name>", "header field used for scalar config object values", "Value")
    .option("--headers <json>", "JSON array of header strings")
    .option("--mode <mode>", "sync mode: replace or upsert", parseConfigTableSyncMode, "replace")
    .option("--header-row <row>", "header row used for record mapping", parsePositiveInteger, 1)
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          field: string;
          fromJson: string;
          headerRow: number;
          headers?: string;
          inPlace?: boolean;
          mode: ConfigTableSyncMode;
          output?: string;
          sheet: string;
          valueField: string;
        },
      ) => {
        const inputPath = resolveFrom(io.cwd, file);
        const jsonPath = resolveFrom(io.cwd, options.fromJson);
        const outputPath = resolveOutputPath(inputPath, {
          inPlace: options.inPlace === true,
          output: options.output ? resolveFrom(io.cwd, options.output) : undefined,
        });
        const workbook = await Workbook.open(inputPath);
        const sheet = getOrCreateSheet(workbook, options.sheet);
        const syncInput = await readConfigTableSyncInput(jsonPath, options.field, options.valueField);
        const explicitHeaders =
          options.headers === undefined ? undefined : parseJsonStringArray(options.headers, "--headers");
        const headers = resolveConfigTableHeaders(
          sheet,
          options.headerRow,
          explicitHeaders ?? syncInput.headers,
          syncInput.records,
        );

        sheet.setHeaders(headers, options.headerRow);

        if (options.mode === "replace") {
          sheet.setRecords(syncInput.records, options.headerRow);
        } else {
          for (const record of syncInput.records) {
            const matchValue = resolveUpsertMatchValue(record, options.field);
            const matchedRow = findConfigTableRow(sheet, options.headerRow, options.field, matchValue)?.row ?? null;

            if (matchedRow === null) {
              sheet.addRecord(record, options.headerRow);
            } else {
              sheet.setRecord(matchedRow, record, options.headerRow);
            }
          }
        }

        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "configTable.sync",
          headers,
          input: inputPath,
          mode: options.mode,
          output: outputPath,
          rows: (await getConfigTableRows(outputPath, options.sheet, options.headerRow)).rows,
          sheet: options.sheet,
          source: jsonPath,
        });
      },
    );

  const table = program
    .command("table")
    .description("Operate on structured sheets with explicit header and data row boundaries");

  table
    .command("generate-profiles")
    .argument("<files...>", "input xlsx files")
    .option("--sheet-filter <regex>", "regular expression used to select sheet names", parseRegex)
    .option("--output <file>", "write generated profiles JSON to a file")
    .action(
      async (
        files: string[],
        options: {
          output?: string;
          sheetFilter?: RegExp;
        },
      ) => {
        const inputPaths = files.map((file) => resolveFrom(io.cwd, file));
        const result = await generateTableProfiles(inputPaths, {
          sheetFilter: options.sheetFilter,
        });
        const outputPath = options.output ? resolveFrom(io.cwd, options.output) : null;

        if (outputPath) {
          await writeFile(outputPath, `${JSON.stringify({ profiles: result.profiles }, null, 2)}\n`);
        }

        writeJson(io.stdout, {
          ...result,
          output: outputPath,
        });
      },
    );

  table
    .command("inspect")
    .argument("<file>", "input xlsx file")
    .option("--sheet <name>", "sheet name")
    .option("--header-row <row>", "row number containing field names", parsePositiveInteger)
    .option("--data-start-row <row>", "first row containing actual data", parsePositiveInteger)
    .option("--profile <name>", "table profile name")
    .option("--profiles-file <file>", "JSON file containing table profiles")
    .action(
      async (
        file: string,
        options: {
          dataStartRow?: number;
          headerRow?: number;
          profile?: string;
          profilesFile?: string;
          sheet?: string;
        },
      ) => {
        const context = await resolveCliTableCommandContext(io.cwd, options);
        const result = await inspectTable(
          resolveFrom(io.cwd, file),
          context.sheet,
          context.headerRow,
          context.dataStartRow,
        );
        writeJson(io.stdout, result);
      },
    );

  table
    .command("list")
    .argument("<file>", "input xlsx file")
    .option("--sheet <name>", "sheet name")
    .option("--header-row <row>", "row number containing field names", parsePositiveInteger)
    .option("--data-start-row <row>", "first row containing actual data", parsePositiveInteger)
    .option("--profile <name>", "table profile name")
    .option("--profiles-file <file>", "JSON file containing table profiles")
    .action(
      async (
        file: string,
        options: {
          dataStartRow?: number;
          headerRow?: number;
          profile?: string;
          profilesFile?: string;
          sheet?: string;
        },
      ) => {
        const context = await resolveCliTableCommandContext(io.cwd, options);
        const result = await getStructuredTableRows(
          resolveFrom(io.cwd, file),
          context.sheet,
          context.headerRow,
          context.dataStartRow,
        );
        writeJson(io.stdout, result);
      },
    );

  table
    .command("get")
    .argument("<file>", "input xlsx file")
    .option("--sheet <name>", "sheet name")
    .option("--header-row <row>", "row number containing field names", parsePositiveInteger)
    .option("--data-start-row <row>", "first row containing actual data", parsePositiveInteger)
    .requiredOption("--key <json>", "JSON scalar or object used to locate the row")
    .option("--key-field <name>", "key field name", collectRepeatedStrings, [])
    .option("--profile <name>", "table profile name")
    .option("--profiles-file <file>", "JSON file containing table profiles")
    .action(
      async (
        file: string,
        options: {
          dataStartRow?: number;
          headerRow?: number;
          key: string;
          keyField: string[];
          profile?: string;
          profilesFile?: string;
          sheet?: string;
        },
      ) => {
        const context = await resolveCliTableCommandContext(io.cwd, options);
        const result = await getStructuredTableRecord(
          resolveFrom(io.cwd, file),
          context.sheet,
          context.headerRow,
          context.dataStartRow,
          context.keyFields,
          options.key,
        );
        writeJson(io.stdout, result);
      },
    );

  table
    .command("upsert")
    .argument("<file>", "input xlsx file")
    .option("--sheet <name>", "sheet name")
    .option("--header-row <row>", "row number containing field names", parsePositiveInteger)
    .option("--data-start-row <row>", "first row containing actual data", parsePositiveInteger)
    .requiredOption("--record <json>", "JSON object keyed by field names")
    .option("--key-field <name>", "key field name", collectRepeatedStrings, [])
    .option("--profile <name>", "table profile name")
    .option("--profiles-file <file>", "JSON file containing table profiles")
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          dataStartRow?: number;
          headerRow?: number;
          inPlace?: boolean;
          keyField: string[];
          output?: string;
          profile?: string;
          profilesFile?: string;
          record: string;
          sheet?: string;
        },
      ) => {
        const context = await resolveCliTableCommandContext(io.cwd, options);
        const inputPath = resolveFrom(io.cwd, file);
        const outputPath = resolveOutputPath(inputPath, {
          inPlace: options.inPlace === true,
          output: options.output ? resolveFrom(io.cwd, options.output) : undefined,
        });
        const workbook = await Workbook.open(inputPath);
        const sheet = workbook.getSheet(context.sheet);
        const record = parseJsonCellRecord(options.record, "--record");
        const keyFields = resolveTableKeyFields(sheet, context.headerRow, context.keyFields);
        const keyRecord = pickKeyRecord(record, keyFields);
        const matchedRow =
          findStructuredTableRow(sheet, context.headerRow, context.dataStartRow, keyFields, keyRecord)?.row ?? null;

        if (matchedRow === null) {
          sheet.addRecord(record, context.headerRow);
        } else {
          writeStructuredTableRecord(sheet, context.headerRow, matchedRow, record);
        }

        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "table.upsert",
          dataStartRow: context.dataStartRow,
          headerRow: context.headerRow,
          input: inputPath,
          keyFields,
          output: outputPath,
          record,
          rows: (await getStructuredTableRows(outputPath, context.sheet, context.headerRow, context.dataStartRow)).rows,
          sheet: context.sheet,
        });
      },
    );

  table
    .command("delete")
    .argument("<file>", "input xlsx file")
    .option("--sheet <name>", "sheet name")
    .option("--header-row <row>", "row number containing field names", parsePositiveInteger)
    .option("--data-start-row <row>", "first row containing actual data", parsePositiveInteger)
    .requiredOption("--key <json>", "JSON scalar or object used to locate the row")
    .option("--key-field <name>", "key field name", collectRepeatedStrings, [])
    .option("--profile <name>", "table profile name")
    .option("--profiles-file <file>", "JSON file containing table profiles")
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          dataStartRow?: number;
          headerRow?: number;
          inPlace?: boolean;
          key: string;
          keyField: string[];
          output?: string;
          profile?: string;
          profilesFile?: string;
          sheet?: string;
        },
      ) => {
        const context = await resolveCliTableCommandContext(io.cwd, options);
        const inputPath = resolveFrom(io.cwd, file);
        const outputPath = resolveOutputPath(inputPath, {
          inPlace: options.inPlace === true,
          output: options.output ? resolveFrom(io.cwd, options.output) : undefined,
        });
        const workbook = await Workbook.open(inputPath);
        const sheet = workbook.getSheet(context.sheet);
        const keyFields = resolveTableKeyFields(sheet, context.headerRow, context.keyFields);
        const keyRecord = parseTableKey(options.key, keyFields, "--key");
        const row =
          findStructuredTableRow(sheet, context.headerRow, context.dataStartRow, keyFields, keyRecord)?.row ?? null;

        if (row !== null) {
          sheet.deleteRecord(row, context.headerRow);
        }

        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "table.delete",
          dataStartRow: context.dataStartRow,
          deleted: row !== null,
          headerRow: context.headerRow,
          input: inputPath,
          key: keyRecord,
          keyFields,
          output: outputPath,
          row,
          rows: (await getStructuredTableRows(outputPath, context.sheet, context.headerRow, context.dataStartRow)).rows,
          sheet: context.sheet,
        });
      },
    );

  table
    .command("sync")
    .argument("<file>", "input xlsx file")
    .option("--sheet <name>", "sheet name")
    .option("--header-row <row>", "row number containing field names", parsePositiveInteger)
    .option("--data-start-row <row>", "first row containing actual data", parsePositiveInteger)
    .requiredOption("--from-json <file>", "JSON file containing records or a config object")
    .option("--key-field <name>", "key field name", collectRepeatedStrings, [])
    .option("--profile <name>", "table profile name")
    .option("--profiles-file <file>", "JSON file containing table profiles")
    .option("--value-field <name>", "field name used when normalizing scalar config objects", "Value")
    .option("--headers <json>", "JSON array of header strings")
    .option("--mode <mode>", "sync mode: replace or upsert", parseConfigTableSyncMode, "replace")
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          dataStartRow?: number;
          fromJson: string;
          headerRow?: number;
          headers?: string;
          inPlace?: boolean;
          keyField: string[];
          mode: ConfigTableSyncMode;
          output?: string;
          profile?: string;
          profilesFile?: string;
          sheet?: string;
          valueField: string;
        },
      ) => {
        const context = await resolveCliTableCommandContext(io.cwd, options);
        const inputPath = resolveFrom(io.cwd, file);
        const jsonPath = resolveFrom(io.cwd, options.fromJson);
        const outputPath = resolveOutputPath(inputPath, {
          inPlace: options.inPlace === true,
          output: options.output ? resolveFrom(io.cwd, options.output) : undefined,
        });
        const workbook = await Workbook.open(inputPath);
        const sheet = workbook.getSheet(context.sheet);
        const syncInput = await readConfigTableSyncInput(
          jsonPath,
          context.keyFields[0] ?? "Key",
          options.valueField,
        );
        const explicitHeaders =
          options.headers === undefined ? undefined : parseJsonStringArray(options.headers, "--headers");
        const headers = resolveConfigTableHeaders(
          sheet,
          context.headerRow,
          explicitHeaders ?? syncInput.headers,
          syncInput.records,
        );

        sheet.setHeaders(headers, context.headerRow);

        if (options.mode === "replace") {
          writeStructuredTableRecords(sheet, context.headerRow, context.dataStartRow, syncInput.records);
        } else {
          const keyFields = resolveTableKeyFields(sheet, context.headerRow, context.keyFields);
          for (const record of syncInput.records) {
            const keyRecord = pickKeyRecord(record, keyFields);
            const matchedRow =
              findStructuredTableRow(sheet, context.headerRow, context.dataStartRow, keyFields, keyRecord)?.row ??
              null;

            if (matchedRow === null) {
              sheet.addRecord(record, context.headerRow);
            } else {
              writeStructuredTableRecord(sheet, context.headerRow, matchedRow, record);
            }
          }
        }

        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "table.sync",
          dataStartRow: context.dataStartRow,
          headerRow: context.headerRow,
          input: inputPath,
          mode: options.mode,
          output: outputPath,
          rows: (await getStructuredTableRows(outputPath, context.sheet, context.headerRow, context.dataStartRow)).rows,
          sheet: context.sheet,
          source: jsonPath,
        });
      },
    );
}

function parseConfigTableSyncMode(value: string): ConfigTableSyncMode {
  if (value === "replace" || value === "upsert") {
    return value;
  }

  throw new InvalidArgumentError(`Expected replace or upsert, got: ${value}`);
}

function parseRegex(value: string): RegExp {
  try {
    return new RegExp(value);
  } catch (error) {
    throw new InvalidArgumentError(`Expected a valid regular expression, got: ${value}; ${formatError(error)}`);
  }
}

function getOrCreateSheet(workbook: Workbook, sheetName: string) {
  const existingSheet = workbook.getSheets().find((sheet) => sheet.name === sheetName);
  return existingSheet ?? workbook.addSheet(sheetName);
}

function collectRepeatedStrings(value: string, previous: string[] = []): string[] {
  return [...previous, value];
}

function resolveFrom(cwd: string, targetPath: string): string {
  return resolve(cwd, targetPath);
}

async function resolveCliTableCommandContext(
  cwd: string,
  options: {
    dataStartRow?: number;
    headerRow?: number;
    keyField?: string[];
    profile?: string;
    profilesFile?: string;
    sheet?: string;
  },
): Promise<TableCommandContext> {
  return resolveTableCommandContext(
    options,
    options.profile ? resolveFrom(cwd, options.profilesFile ?? "table-profiles.json") : undefined,
  );
}
