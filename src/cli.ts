#!/usr/bin/env node

import { readFile } from "node:fs/promises";
import { resolve } from "node:path";
import { fileURLToPath } from "node:url";

import { Command, CommanderError, InvalidArgumentError } from "commander";

import type { CellValue, DefinedName, SheetVisibility } from "./types.js";
import { validateRoundtripFile } from "./roundtrip.js";
import { Workbook } from "./workbook.js";

type Writer = (chunk: string) => void;

interface CliIo {
  cwd?: string;
  stderr?: Writer;
  stdout?: Writer;
}

interface InspectResult {
  activeSheet: string;
  definedNames: DefinedName[];
  file: string;
  sheets: Array<{
    columnCount: number;
    headers: string[];
    name: string;
    rowCount: number;
    usedRange: string | null;
    visibility: SheetVisibility;
  }>;
}

interface GetCellResult {
  backgroundColor: string | null;
  cell: string;
  exists: boolean;
  file: string;
  formula: string | null;
  numberFormat: string | null;
  rawType: string | null;
  sheet: string;
  styleId: number | null;
  type: string;
  value: CellValue;
}

type CellRecord = Record<string, CellValue>;
type ConfigTableRow = {
  record: CellRecord;
  row: number;
};
type ConfigTableSyncMode = "replace" | "upsert";

type WorkbookOperation =
  | {
      headerRow?: number;
      record: CellRecord;
      sheet: string;
      type: "addRecord";
    }
  | {
      headerRow?: number;
      records: CellRecord[];
      sheet: string;
      type: "addRecords";
    }
  | {
      cell: string;
      color: string | null;
      sheet: string;
      type: "setBackgroundColor";
    }
  | {
      from: string;
      sheet: string;
      to: string;
      type: "copyStyle";
    }
  | {
      cell: string;
      sheet: string;
      type: "clearCell";
    }
  | {
      headerRow?: number;
      row: number;
      sheet: string;
      type: "deleteRecord";
    }
  | {
      headerRow?: number;
      rows: number[];
      sheet: string;
      type: "deleteRecords";
    }
  | {
      name: string;
      scope?: string;
      type: "deleteDefinedName";
    }
  | {
      sheet: string;
      type: "addSheet";
    }
  | {
      sheet: string;
      type: "deleteSheet";
    }
  | {
      from: string;
      to: string;
      type: "renameSheet";
    }
  | {
      headerRow?: number;
      record: CellRecord;
      row: number;
      sheet: string;
      type: "setRecord";
    }
  | {
      headerRow?: number;
      records: CellRecord[];
      sheet: string;
      type: "setRecords";
    }
  | {
      cachedValue?: CellValue;
      cell: string;
      formula: string;
      sheet: string;
      type: "setFormula";
    }
  | {
      cell: string;
      formatCode: string;
      sheet: string;
      type: "setNumberFormat";
    }
  | {
      cell: string;
      sheet: string;
      type: "setCell";
      value: CellValue;
    }
  | {
      sheet: string;
      type: "setActiveSheet";
    }
  | {
      headerRow?: number;
      headers: string[];
      sheet: string;
      startColumn?: number;
      type: "setHeaders";
    }
  | {
      name: string;
      scope?: string;
      type: "setDefinedName";
      value: string;
    }
  | {
      sheet: string;
      type: "setSheetVisibility";
      visibility: SheetVisibility;
    };

interface OpsDocument {
  actions: WorkbookOperation[];
  output?: string;
}

class CliExitError extends Error {
  readonly exitCode: number;

  constructor(exitCode: number) {
    super(`CLI exited with code ${exitCode}`);
    this.exitCode = exitCode;
  }
}

export async function runCli(argv: string[], io: CliIo = {}): Promise<number> {
  const stdout = io.stdout ?? ((chunk: string) => process.stdout.write(chunk));
  const stderr = io.stderr ?? ((chunk: string) => process.stderr.write(chunk));
  const cwd = io.cwd ?? process.cwd();
  const program = createProgram({ cwd, stderr, stdout });

  try {
    await program.parseAsync(["node", "xlsx-ts", ...argv], { from: "node" });
    return 0;
  } catch (error) {
    if (error instanceof CliExitError) {
      return error.exitCode;
    }

    if (error instanceof CommanderError) {
      return error.exitCode;
    }

    stderr(`${formatError(error)}\n`);
    return 1;
  }
}

function createProgram(io: Required<CliIo>): Command {
  const program = new Command()
    .name("xlsx-ts")
    .description("Lossless-first XLSX inspection and editing CLI")
    .showHelpAfterError()
    .configureOutput({
      writeErr: io.stderr,
      writeOut: io.stdout,
    })
    .exitOverride();

  program
    .command("inspect")
    .argument("<file>", "input xlsx file")
    .option("--header-row <row>", "header row used for the sheet preview", parsePositiveInteger, 1)
    .action(async (file: string, options: { headerRow: number }) => {
      const result = await inspectWorkbook(resolveFrom(io.cwd, file), options.headerRow);
      writeJson(io.stdout, result);
    });

  program
    .command("get")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--cell <address>", "cell address, such as B2")
    .action(async (file: string, options: { cell: string; sheet: string }) => {
      const result = await getCell(resolveFrom(io.cwd, file), options.sheet, options.cell);
      writeJson(io.stdout, result);
    });

  program
    .command("records")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .option("--header-row <row>", "header row used for record mapping", parsePositiveInteger, 1)
    .action(async (file: string, options: { headerRow: number; sheet: string }) => {
      const result = await getRecords(resolveFrom(io.cwd, file), options.sheet, options.headerRow);
      writeJson(io.stdout, result);
    });

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

  program
    .command("add-sheet")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "new sheet name")
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          inPlace?: boolean;
          output?: string;
          sheet: string;
        },
      ) => {
        const inputPath = resolveFrom(io.cwd, file);
        const outputPath = resolveOutputPath(inputPath, {
          inPlace: options.inPlace === true,
          output: options.output ? resolveFrom(io.cwd, options.output) : undefined,
        });
        const workbook = await Workbook.open(inputPath);
        workbook.addSheet(options.sheet);
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "addSheet",
          input: inputPath,
          output: outputPath,
          sheets: workbook.getSheets().map((sheet) => sheet.name),
        });
      },
    );

  program
    .command("rename-sheet")
    .argument("<file>", "input xlsx file")
    .requiredOption("--from <name>", "current sheet name")
    .requiredOption("--to <name>", "new sheet name")
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          from: string;
          inPlace?: boolean;
          output?: string;
          to: string;
        },
      ) => {
        const inputPath = resolveFrom(io.cwd, file);
        const outputPath = resolveOutputPath(inputPath, {
          inPlace: options.inPlace === true,
          output: options.output ? resolveFrom(io.cwd, options.output) : undefined,
        });
        const workbook = await Workbook.open(inputPath);
        workbook.renameSheet(options.from, options.to);
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "renameSheet",
          input: inputPath,
          output: outputPath,
          sheets: workbook.getSheets().map((sheet) => sheet.name),
        });
      },
    );

  program
    .command("delete-sheet")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name to delete")
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          inPlace?: boolean;
          output?: string;
          sheet: string;
        },
      ) => {
        const inputPath = resolveFrom(io.cwd, file);
        const outputPath = resolveOutputPath(inputPath, {
          inPlace: options.inPlace === true,
          output: options.output ? resolveFrom(io.cwd, options.output) : undefined,
        });
        const workbook = await Workbook.open(inputPath);
        workbook.deleteSheet(options.sheet);
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "deleteSheet",
          input: inputPath,
          output: outputPath,
          sheets: workbook.getSheets().map((sheet) => sheet.name),
        });
      },
    );

  program
    .command("set-headers")
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
        const sheet = workbook.getSheet(options.sheet);
        const headers = parseJsonStringArray(options.headers, "--headers");
        sheet.setHeaders(headers, options.headerRow, options.startColumn);
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "setHeaders",
          headers: sheet.getHeaders(options.headerRow),
          input: inputPath,
          output: outputPath,
          sheet: options.sheet,
        });
      },
    );

  program
    .command("copy-style")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--from <cell>", "source cell address")
    .requiredOption("--to <cell>", "target cell address")
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          from: string;
          inPlace?: boolean;
          output?: string;
          sheet: string;
          to: string;
        },
      ) => {
        const inputPath = resolveFrom(io.cwd, file);
        const outputPath = resolveOutputPath(inputPath, {
          inPlace: options.inPlace === true,
          output: options.output ? resolveFrom(io.cwd, options.output) : undefined,
        });
        const workbook = await Workbook.open(inputPath);
        const sheet = workbook.getSheet(options.sheet);
        sheet.copyStyle(options.from, options.to);
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "copyStyle",
          input: inputPath,
          output: outputPath,
          result: await getCell(outputPath, options.sheet, options.to),
        });
      },
    );

  program
    .command("add-record")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--record <json>", "JSON object keyed by header names")
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
        sheet.addRecord(record, options.headerRow);
        await workbook.save(outputPath);
        const result = await getRecords(outputPath, options.sheet, options.headerRow);
        writeJson(io.stdout, {
          action: "addRecord",
          input: inputPath,
          output: outputPath,
          record,
          records: result.records,
          sheet: options.sheet,
        });
      },
    );

  program
    .command("set-number-format")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--cell <address>", "cell address, such as B2")
    .requiredOption("--format <code>", "number format code")
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          cell: string;
          format: string;
          inPlace?: boolean;
          output?: string;
          sheet: string;
        },
      ) => {
        const inputPath = resolveFrom(io.cwd, file);
        const outputPath = resolveOutputPath(inputPath, {
          inPlace: options.inPlace === true,
          output: options.output ? resolveFrom(io.cwd, options.output) : undefined,
        });
        const workbook = await Workbook.open(inputPath);
        workbook.getSheet(options.sheet).setNumberFormat(options.cell, options.format);
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "setNumberFormat",
          input: inputPath,
          output: outputPath,
          result: await getCell(outputPath, options.sheet, options.cell),
        });
      },
    );

  program
    .command("set-background-color")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--cell <address>", "cell address, such as B2")
    .option("--color <rgb>", "ARGB color, such as FFFF0000")
    .option("--clear", "remove the solid background fill")
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          cell: string;
          clear?: boolean;
          color?: string;
          inPlace?: boolean;
          output?: string;
          sheet: string;
        },
      ) => {
        const actionCount = Number(Boolean(options.clear)) + Number(options.color !== undefined);
        if (actionCount !== 1) {
          throw new Error("Exactly one of --color or --clear is required");
        }

        const inputPath = resolveFrom(io.cwd, file);
        const outputPath = resolveOutputPath(inputPath, {
          inPlace: options.inPlace === true,
          output: options.output ? resolveFrom(io.cwd, options.output) : undefined,
        });
        const workbook = await Workbook.open(inputPath);
        workbook.getSheet(options.sheet).setBackgroundColor(options.cell, options.clear ? null : options.color!);
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "setBackgroundColor",
          input: inputPath,
          output: outputPath,
          result: await getCell(outputPath, options.sheet, options.cell),
        });
      },
    );

  program
    .command("set-record")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--row <row>", "1-based row number", parsePositiveInteger)
    .requiredOption("--record <json>", "JSON object keyed by header names")
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
          record: string;
          row: number;
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
        sheet.setRecord(options.row, record, options.headerRow);
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "setRecord",
          input: inputPath,
          output: outputPath,
          record: await getRecord(outputPath, options.sheet, options.row, options.headerRow),
          row: options.row,
          sheet: options.sheet,
        });
      },
    );

  program
    .command("set-records")
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
          action: "setRecords",
          input: inputPath,
          output: outputPath,
          records: (await getRecords(outputPath, options.sheet, options.headerRow)).records,
          sheet: options.sheet,
        });
      },
    );

  program
    .command("delete-record")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--row <row>", "1-based row number", parsePositiveInteger)
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
          row: number;
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
        sheet.deleteRecord(options.row, options.headerRow);
        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: "deleteRecord",
          input: inputPath,
          output: outputPath,
          records: (await getRecords(outputPath, options.sheet, options.headerRow)).records,
          row: options.row,
          sheet: options.sheet,
        });
      },
    );

  program
    .command("set")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .requiredOption("--cell <address>", "cell address, such as B2")
    .option("--value <json>", "JSON literal for the next cell value")
    .option("--text <value>", "plain string value without JSON quoting")
    .option("--formula <formula>", "formula text without the leading equals sign")
    .option("--cached-value <json>", "JSON literal for the formula cached value")
    .option("--cached-text <value>", "plain string cached value for a formula")
    .option("--clear", "delete the cell instead of writing a value")
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          cachedText?: string;
          cachedValue?: string;
          cell: string;
          clear?: boolean;
          formula?: string;
          inPlace?: boolean;
          output?: string;
          sheet: string;
          text?: string;
          value?: string;
        },
      ) => {
        const inputPath = resolveFrom(io.cwd, file);
        const outputPath = resolveOutputPath(inputPath, {
          inPlace: options.inPlace === true,
          output: options.output ? resolveFrom(io.cwd, options.output) : undefined,
        });
        const workbook = await Workbook.open(inputPath);
        const sheet = workbook.getSheet(options.sheet);
        const actionCount =
          Number(Boolean(options.clear)) +
          Number(options.formula !== undefined) +
          Number(options.text !== undefined) +
          Number(options.value !== undefined);

        if (actionCount !== 1) {
          throw new Error("Exactly one of --value, --text, --formula, or --clear is required");
        }

        if (options.formula !== undefined) {
          const cachedValue =
            options.cachedText !== undefined
              ? options.cachedText
              : options.cachedValue !== undefined
                ? parseJsonCellValue(options.cachedValue, "--cached-value")
                : undefined;
          sheet.setFormula(options.cell, options.formula, cachedValue === undefined ? {} : { cachedValue });
        } else if (options.clear) {
          sheet.deleteCell(options.cell);
        } else {
          const value =
            options.text !== undefined
              ? options.text
              : parseJsonCellValue(options.value!, "--value");
          sheet.setCell(options.cell, value);
        }

        await workbook.save(outputPath);
        writeJson(io.stdout, {
          action: options.clear ? "clearCell" : options.formula !== undefined ? "setFormula" : "setCell",
          input: inputPath,
          output: outputPath,
          result: await getCell(outputPath, options.sheet, options.cell),
        });
      },
    );

  program
    .command("apply")
    .argument("<file>", "input xlsx file")
    .requiredOption("--ops <file>", "JSON document with workbook actions")
    .option("--output <file>", "output xlsx path")
    .option("--in-place", "overwrite the input workbook")
    .action(
      async (
        file: string,
        options: {
          inPlace?: boolean;
          ops: string;
          output?: string;
        },
      ) => {
        const inputPath = resolveFrom(io.cwd, file);
        const opsPath = resolveFrom(io.cwd, options.ops);
        const document = await readOpsDocument(opsPath);
        const configuredOutput = document.output ? resolveFrom(io.cwd, document.output) : undefined;
        const outputPath = resolveOutputPath(inputPath, {
          inPlace: options.inPlace === true,
          output: options.output ? resolveFrom(io.cwd, options.output) : configuredOutput,
        });
        const workbook = await Workbook.open(inputPath);

        for (const action of document.actions) {
          applyWorkbookOperation(workbook, action);
        }

        await workbook.save(outputPath);
        writeJson(io.stdout, {
          actions: document.actions.length,
          input: inputPath,
          ops: opsPath,
          output: outputPath,
          sheets: workbook.getSheets().map((sheet) => sheet.name),
        });
      },
    );

  program
    .command("validate")
    .argument("<file>", "input xlsx file")
    .option("--output <file>", "persist the roundtrip workbook to the given path")
    .action(async (file: string, options: { output?: string }) => {
      const result = await validateRoundtripFile(
        resolveFrom(io.cwd, file),
        options.output ? resolveFrom(io.cwd, options.output) : undefined,
      );
      writeJson(io.stdout, result);

      if (!result.ok) {
        throw new CliExitError(1);
      }
    });

  return program;
}

async function inspectWorkbook(filePath: string, headerRow: number): Promise<InspectResult> {
  const workbook = await Workbook.open(filePath);
  const sheets = workbook.getSheets().map((sheet) => ({
    columnCount: sheet.columnCount,
    headers: trimTrailingEmptyStrings(sheet.getHeaders(headerRow)),
    name: sheet.name,
    rowCount: sheet.rowCount,
    usedRange: sheet.getUsedRange(),
    visibility: workbook.getSheetVisibility(sheet.name),
  }));

  return {
    activeSheet: workbook.getActiveSheet().name,
    definedNames: workbook.getDefinedNames(),
    file: filePath,
    sheets,
  };
}

async function getCell(filePath: string, sheetName: string, cellAddress: string): Promise<GetCellResult> {
  const workbook = await Workbook.open(filePath);
  const sheet = workbook.getSheet(sheetName);
  const cell = sheet.cell(cellAddress);

  return {
    backgroundColor: sheet.getBackgroundColor(cellAddress),
    cell: cellAddress,
    exists: cell.exists,
    file: filePath,
    formula: cell.formula,
    numberFormat: cell.numberFormat?.code ?? null,
    rawType: cell.rawType,
    sheet: sheetName,
    styleId: cell.styleId,
    type: cell.type,
    value: cell.value,
  };
}

async function getRecords(
  filePath: string,
  sheetName: string,
  headerRow: number,
): Promise<{
  file: string;
  headerRow: number;
  records: CellRecord[];
  sheet: string;
}> {
  const workbook = await Workbook.open(filePath);
  const sheet = workbook.getSheet(sheetName);
  return {
    file: filePath,
    headerRow,
    records: sheet.getRecords(headerRow),
    sheet: sheetName,
  };
}

async function getRecord(
  filePath: string,
  sheetName: string,
  row: number,
  headerRow: number,
): Promise<CellRecord | null> {
  const workbook = await Workbook.open(filePath);
  return workbook.getSheet(sheetName).getRecord(row, headerRow);
}

async function getConfigTableRows(
  filePath: string,
  sheetName: string,
  headerRow: number,
): Promise<{
  file: string;
  headerRow: number;
  rows: ConfigTableRow[];
  sheet: string;
}> {
  const workbook = await Workbook.open(filePath);
  const sheet = workbook.getSheet(sheetName);
  return {
    file: filePath,
    headerRow,
    rows: listConfigTableRows(sheet, headerRow),
    sheet: sheetName,
  };
}

async function getConfigTableRecord(
  filePath: string,
  sheetName: string,
  headerRow: number,
  field: string,
  value: CellValue,
): Promise<{
  field: string;
  file: string;
  headerRow: number;
  record: ConfigTableRow | null;
  sheet: string;
  value: CellValue;
}> {
  const workbook = await Workbook.open(filePath);
  const sheet = workbook.getSheet(sheetName);
  return {
    field,
    file: filePath,
    headerRow,
    record: findConfigTableRow(sheet, headerRow, field, value) ?? null,
    sheet: sheetName,
    value,
  };
}

async function readConfigTableSyncInput(
  filePath: string,
  field: string,
  valueField: string,
): Promise<{
  headers?: string[];
  records: CellRecord[];
}> {
  const parsed = parseJsonDocument(await readFile(filePath, "utf8"), filePath);

  if (Array.isArray(parsed)) {
    return { records: assertCellRecordArray(parsed, `${filePath}`) };
  }

  const record = assertRecord(parsed, filePath);
  if (record.records !== undefined) {
    return {
      headers: record.headers === undefined ? undefined : assertStringArray(record.headers, `${filePath}.headers`),
      records: assertCellRecordArray(record.records, `${filePath}.records`),
    };
  }

  return {
    records: normalizeConfigObjectToRecords(record, filePath, field, valueField),
  };
}

async function readOpsDocument(filePath: string): Promise<OpsDocument> {
  const parsed = parseJsonDocument(await readFile(filePath, "utf8"), filePath);

  if (Array.isArray(parsed)) {
    return {
      actions: parsed.map((item, index) => parseWorkbookOperation(item, `${filePath}[${index}]`)),
    };
  }

  const record = assertRecord(parsed, filePath);
  const actions = assertArray(record.actions, `${filePath}.actions`);
  return {
    actions: actions.map((item, index) => parseWorkbookOperation(item, `${filePath}.actions[${index}]`)),
    output: record.output === undefined ? undefined : assertString(record.output, `${filePath}.output`),
  };
}

function parseWorkbookOperation(value: unknown, label: string): WorkbookOperation {
  const record = assertRecord(value, label);
  const type = assertString(record.type, `${label}.type`);

  switch (type) {
    case "addRecord":
      return {
        headerRow: optionalPositiveInteger(record.headerRow, `${label}.headerRow`),
        record: assertCellRecord(record.record, `${label}.record`),
        sheet: assertString(record.sheet, `${label}.sheet`),
        type,
      };
    case "addRecords":
      return {
        headerRow: optionalPositiveInteger(record.headerRow, `${label}.headerRow`),
        records: assertCellRecordArray(record.records, `${label}.records`),
        sheet: assertString(record.sheet, `${label}.sheet`),
        type,
      };
    case "addSheet":
      return {
        sheet: assertString(record.sheet, `${label}.sheet`),
        type,
      };
    case "copyStyle":
      return {
        from: assertString(record.from, `${label}.from`),
        sheet: assertString(record.sheet, `${label}.sheet`),
        to: assertString(record.to, `${label}.to`),
        type,
      };
    case "clearCell":
      return {
        cell: assertString(record.cell, `${label}.cell`),
        sheet: assertString(record.sheet, `${label}.sheet`),
        type,
      };
    case "deleteRecord":
      return {
        headerRow: optionalPositiveInteger(record.headerRow, `${label}.headerRow`),
        row: assertPositiveInteger(record.row, `${label}.row`),
        sheet: assertString(record.sheet, `${label}.sheet`),
        type,
      };
    case "deleteRecords":
      return {
        headerRow: optionalPositiveInteger(record.headerRow, `${label}.headerRow`),
        rows: assertPositiveIntegerArray(record.rows, `${label}.rows`),
        sheet: assertString(record.sheet, `${label}.sheet`),
        type,
      };
    case "deleteDefinedName":
      return {
        name: assertString(record.name, `${label}.name`),
        scope: optionalString(record.scope, `${label}.scope`),
        type,
      };
    case "deleteSheet":
      return {
        sheet: assertString(record.sheet, `${label}.sheet`),
        type,
      };
    case "renameSheet":
      return {
        from: assertString(record.from, `${label}.from`),
        to: assertString(record.to, `${label}.to`),
        type,
      };
    case "setBackgroundColor":
      return {
        cell: assertString(record.cell, `${label}.cell`),
        color: assertNullableString(record.color, `${label}.color`),
        sheet: assertString(record.sheet, `${label}.sheet`),
        type,
      };
    case "setHeaders":
      return {
        headerRow: optionalPositiveInteger(record.headerRow, `${label}.headerRow`),
        headers: assertStringArray(record.headers, `${label}.headers`),
        sheet: assertString(record.sheet, `${label}.sheet`),
        startColumn: optionalPositiveInteger(record.startColumn, `${label}.startColumn`),
        type,
      };
    case "setNumberFormat":
      return {
        cell: assertString(record.cell, `${label}.cell`),
        formatCode: assertString(record.formatCode, `${label}.formatCode`),
        sheet: assertString(record.sheet, `${label}.sheet`),
        type,
      };
    case "setRecord":
      return {
        headerRow: optionalPositiveInteger(record.headerRow, `${label}.headerRow`),
        record: assertCellRecord(record.record, `${label}.record`),
        row: assertPositiveInteger(record.row, `${label}.row`),
        sheet: assertString(record.sheet, `${label}.sheet`),
        type,
      };
    case "setRecords":
      return {
        headerRow: optionalPositiveInteger(record.headerRow, `${label}.headerRow`),
        records: assertCellRecordArray(record.records, `${label}.records`),
        sheet: assertString(record.sheet, `${label}.sheet`),
        type,
      };
    case "setActiveSheet":
      return {
        sheet: assertString(record.sheet, `${label}.sheet`),
        type,
      };
    case "setCell":
      return {
        cell: assertString(record.cell, `${label}.cell`),
        sheet: assertString(record.sheet, `${label}.sheet`),
        type,
        value: assertCellValue(record.value, `${label}.value`),
      };
    case "setDefinedName":
      return {
        name: assertString(record.name, `${label}.name`),
        scope: optionalString(record.scope, `${label}.scope`),
        type,
        value: assertString(record.value, `${label}.value`),
      };
    case "setFormula":
      return {
        cachedValue:
          record.cachedValue === undefined
            ? undefined
            : assertCellValue(record.cachedValue, `${label}.cachedValue`),
        cell: assertString(record.cell, `${label}.cell`),
        formula: assertString(record.formula, `${label}.formula`),
        sheet: assertString(record.sheet, `${label}.sheet`),
        type,
      };
    case "setSheetVisibility":
      return {
        sheet: assertString(record.sheet, `${label}.sheet`),
        type,
        visibility: assertSheetVisibility(record.visibility, `${label}.visibility`),
      };
    default:
      throw new Error(`Unsupported operation type at ${label}.type: ${type}`);
  }
}

function applyWorkbookOperation(workbook: Workbook, action: WorkbookOperation): void {
  switch (action.type) {
    case "addRecord":
      workbook.getSheet(action.sheet).addRecord(action.record, action.headerRow ?? 1);
      return;
    case "addRecords":
      workbook.getSheet(action.sheet).addRecords(action.records, action.headerRow ?? 1);
      return;
    case "addSheet":
      workbook.addSheet(action.sheet);
      return;
    case "copyStyle":
      workbook.getSheet(action.sheet).copyStyle(action.from, action.to);
      return;
    case "clearCell":
      workbook.getSheet(action.sheet).deleteCell(action.cell);
      return;
    case "deleteRecord":
      workbook.getSheet(action.sheet).deleteRecord(action.row, action.headerRow ?? 1);
      return;
    case "deleteRecords":
      workbook.getSheet(action.sheet).deleteRecords(action.rows, action.headerRow ?? 1);
      return;
    case "deleteDefinedName":
      workbook.deleteDefinedName(action.name, action.scope);
      return;
    case "deleteSheet":
      workbook.deleteSheet(action.sheet);
      return;
    case "renameSheet":
      workbook.renameSheet(action.from, action.to);
      return;
    case "setBackgroundColor":
      workbook.getSheet(action.sheet).setBackgroundColor(action.cell, action.color);
      return;
    case "setHeaders":
      workbook.getSheet(action.sheet).setHeaders(
        action.headers,
        action.headerRow ?? 1,
        action.startColumn ?? 1,
      );
      return;
    case "setRecord":
      workbook.getSheet(action.sheet).setRecord(action.row, action.record, action.headerRow ?? 1);
      return;
    case "setRecords":
      workbook.getSheet(action.sheet).setRecords(action.records, action.headerRow ?? 1);
      return;
    case "setActiveSheet":
      workbook.setActiveSheet(action.sheet);
      return;
    case "setCell":
      workbook.getSheet(action.sheet).setCell(action.cell, action.value);
      return;
    case "setDefinedName":
      workbook.setDefinedName(action.name, action.value, action.scope ? { scope: action.scope } : {});
      return;
    case "setFormula":
      workbook
        .getSheet(action.sheet)
        .setFormula(
          action.cell,
          action.formula,
          action.cachedValue === undefined ? {} : { cachedValue: action.cachedValue },
        );
      return;
    case "setNumberFormat":
      workbook.getSheet(action.sheet).setNumberFormat(action.cell, action.formatCode);
      return;
    case "setSheetVisibility":
      workbook.setSheetVisibility(action.sheet, action.visibility);
      return;
  }
}

function resolveOutputPath(
  inputPath: string,
  options: {
    inPlace: boolean;
    output?: string;
  },
): string {
  if (options.inPlace && options.output) {
    throw new Error("Use either --output or --in-place, not both");
  }

  if (options.inPlace) {
    return inputPath;
  }

  if (options.output) {
    return options.output;
  }

  throw new Error("An output path is required; pass --output or use --in-place");
}

function parsePositiveInteger(value: string): number {
  const parsed = Number(value);
  if (!Number.isInteger(parsed) || parsed <= 0) {
    throw new InvalidArgumentError(`Expected a positive integer, got: ${value}`);
  }

  return parsed;
}

function parseConfigTableSyncMode(value: string): ConfigTableSyncMode {
  if (value === "replace" || value === "upsert") {
    return value;
  }

  throw new InvalidArgumentError(`Expected replace or upsert, got: ${value}`);
}

function resolveMatchValue(value?: string, text?: string): CellValue {
  const actionCount = Number(value !== undefined) + Number(text !== undefined);
  if (actionCount !== 1) {
    throw new Error("Exactly one of --value or --text is required");
  }

  return text !== undefined ? text : parseJsonCellValue(value!, "--value");
}

function resolveUpsertMatchValue(
  record: CellRecord,
  field: string,
  matchValue?: string,
  matchText?: string,
): CellValue {
  const overrideCount = Number(matchValue !== undefined) + Number(matchText !== undefined);
  if (overrideCount > 1) {
    throw new Error("Use either --match-value or --match-text, not both");
  }

  if (matchText !== undefined) {
    return matchText;
  }

  if (matchValue !== undefined) {
    return parseJsonCellValue(matchValue, "--match-value");
  }

  if (!Object.hasOwn(record, field)) {
    throw new Error(`Record is missing the match field: ${field}`);
  }

  return record[field] ?? null;
}

function parseJsonCellValue(source: string, label: string): CellValue {
  return assertCellValue(parseJsonDocument(source, label), label);
}

function parseJsonDocument(source: string, label: string): unknown {
  try {
    return JSON.parse(source);
  } catch (error) {
    throw new Error(`Invalid JSON in ${label}: ${formatError(error)}`);
  }
}

function writeJson(writer: Writer, value: unknown): void {
  writer(`${JSON.stringify(value, null, 2)}\n`);
}

function assertRecord(value: unknown, label: string): Record<string, unknown> {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new Error(`Expected ${label} to be an object`);
  }

  return value as Record<string, unknown>;
}

function assertArray(value: unknown, label: string): unknown[] {
  if (!Array.isArray(value)) {
    throw new Error(`Expected ${label} to be an array`);
  }

  return value;
}

function assertString(value: unknown, label: string): string {
  if (typeof value !== "string" || value.length === 0) {
    throw new Error(`Expected ${label} to be a non-empty string`);
  }

  return value;
}

function optionalString(value: unknown, label: string): string | undefined {
  if (value === undefined) {
    return undefined;
  }

  return assertString(value, label);
}

function assertNullableString(value: unknown, label: string): string | null {
  if (value === null) {
    return null;
  }

  return assertString(value, label);
}

function optionalPositiveInteger(value: unknown, label: string): number | undefined {
  if (value === undefined) {
    return undefined;
  }

  return assertPositiveInteger(value, label);
}

function assertPositiveInteger(value: unknown, label: string): number {
  if (typeof value !== "number" || !Number.isInteger(value) || value <= 0) {
    throw new Error(`Expected ${label} to be a positive integer`);
  }

  return value;
}

function assertPositiveIntegerArray(value: unknown, label: string): number[] {
  const values = assertArray(value, label);
  return values.map((item, index) => assertPositiveInteger(item, `${label}[${index}]`));
}

function assertCellValue(value: unknown, label: string): CellValue {
  if (
    value === null ||
    typeof value === "string" ||
    typeof value === "number" ||
    typeof value === "boolean"
  ) {
    return value;
  }

  throw new Error(`Expected ${label} to be a string, number, boolean, or null`);
}

function assertCellRecord(value: unknown, label: string): CellRecord {
  const record = assertRecord(value, label);
  const next: CellRecord = {};

  for (const [key, item] of Object.entries(record)) {
    next[key] = assertCellValue(item, `${label}.${key}`);
  }

  return next;
}

function assertCellRecordArray(value: unknown, label: string): CellRecord[] {
  const values = assertArray(value, label);
  return values.map((item, index) => assertCellRecord(item, `${label}[${index}]`));
}

function assertStringArray(value: unknown, label: string): string[] {
  const values = assertArray(value, label);
  return values.map((item, index) => {
    if (typeof item !== "string") {
      throw new Error(`Expected ${label}[${index}] to be a string`);
    }

    return item;
  });
}

function assertSheetVisibility(value: unknown, label: string): SheetVisibility {
  if (value === "visible" || value === "hidden" || value === "veryHidden") {
    return value;
  }

  throw new Error(`Expected ${label} to be visible, hidden, or veryHidden`);
}

function trimTrailingEmptyStrings(values: string[]): string[] {
  let end = values.length;

  while (end > 0 && values[end - 1] === "") {
    end -= 1;
  }

  return values.slice(0, end);
}

function getOrCreateSheet(workbook: Workbook, sheetName: string) {
  const existingSheet = workbook.getSheets().find((sheet) => sheet.name === sheetName);
  return existingSheet ?? workbook.addSheet(sheetName);
}

function resolveConfigTableHeaders(
  sheet: ReturnType<Workbook["getSheet"]>,
  headerRow: number,
  explicitHeaders: string[] | undefined,
  records: CellRecord[],
): string[] {
  if (explicitHeaders && explicitHeaders.length > 0) {
    return explicitHeaders;
  }

  const inferredHeaders = inferHeadersFromRecords(records);
  if (inferredHeaders.length > 0) {
    return inferredHeaders;
  }

  const existingHeaders = trimTrailingEmptyStrings(sheet.getHeaders(headerRow));
  if (existingHeaders.length > 0) {
    return existingHeaders;
  }

  throw new Error("Unable to determine headers; provide --headers or include records with keys");
}

function inferHeadersFromRecords(records: CellRecord[]): string[] {
  const headers: string[] = [];
  const seen = new Set<string>();

  for (const record of records) {
    for (const key of Object.keys(record)) {
      if (!seen.has(key)) {
        seen.add(key);
        headers.push(key);
      }
    }
  }

  return headers;
}

function normalizeConfigObjectToRecords(
  value: Record<string, unknown>,
  label: string,
  field: string,
  valueField: string,
): CellRecord[] {
  const records: CellRecord[] = [];

  for (const [key, entry] of Object.entries(value)) {
    if (
      entry === null ||
      typeof entry === "string" ||
      typeof entry === "number" ||
      typeof entry === "boolean"
    ) {
      records.push({
        [field]: key,
        [valueField]: entry,
      });
      continue;
    }

    const nested = assertCellRecord(entry, `${label}.${key}`);
    records.push({
      [field]: key,
      ...nested,
    });
  }

  return records;
}

function listConfigTableRows(
  sheet: ReturnType<Workbook["getSheet"]>,
  headerRow: number,
): ConfigTableRow[] {
  const rows: ConfigTableRow[] = [];

  for (let row = headerRow + 1; row <= sheet.rowCount; row += 1) {
    const record = sheet.getRecord(row, headerRow);
    if (record !== null) {
      rows.push({ record, row });
    }
  }

  return rows;
}

function findConfigTableRow(
  sheet: ReturnType<Workbook["getSheet"]>,
  headerRow: number,
  field: string,
  value: CellValue,
): ConfigTableRow | null {
  return listConfigTableRows(sheet, headerRow).find((row) => row.record[field] === value) ?? null;
}

function resolveFrom(cwd: string, targetPath: string): string {
  return resolve(cwd, targetPath);
}

function parseJsonCellRecord(source: string, label: string): CellRecord {
  return assertCellRecord(parseJsonDocument(source, label), label);
}

function parseJsonCellRecordArray(source: string, label: string): CellRecord[] {
  return assertCellRecordArray(parseJsonDocument(source, label), label);
}

function parseJsonStringArray(source: string, label: string): string[] {
  return assertStringArray(parseJsonDocument(source, label), label);
}

function formatError(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

async function main(): Promise<void> {
  process.exitCode = await runCli(process.argv.slice(2));
}

if (process.argv[1] && fileURLToPath(import.meta.url) === resolve(process.argv[1])) {
  await main();
}
