#!/usr/bin/env node

import { realpathSync } from "node:fs";
import { readFile } from "node:fs/promises";
import { resolve } from "node:path";
import { fileURLToPath } from "node:url";

import { Command, CommanderError, InvalidArgumentError } from "commander";

import type { CellValue, DefinedName, SheetVisibility } from "./types.js";
import {
  assertArray,
  assertCellRecord,
  assertCellRecordArray,
  assertCellValue,
  assertNullableString,
  assertPositiveInteger,
  assertPositiveIntegerArray,
  assertRecord,
  assertSheetVisibility,
  assertString,
  assertStringArray,
  formatError,
  optionalPositiveInteger,
  optionalString,
  parseJsonCellValue,
  parseJsonDocument,
  writeJson,
} from "./cli-json.js";
import type { CellRecord, Writer } from "./cli-json.js";
import { registerRecordCommands } from "./cli-record-commands.js";
import { registerTableCommands } from "./cli-table-commands.js";
import { validateRoundtripFile } from "./roundtrip.js";
import { Workbook } from "./workbook.js";

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

  registerRecordCommands(
    program,
    {
      cwd: io.cwd,
      stdout: io.stdout,
    },
    {
      parsePositiveInteger,
      resolveOutputPath,
    },
  );

  registerTableCommands(
    program,
    {
      cwd: io.cwd,
      stdout: io.stdout,
    },
    {
      parsePositiveInteger,
      resolveOutputPath,
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

function trimTrailingEmptyStrings(values: string[]): string[] {
  let end = values.length;

  while (end > 0 && values[end - 1] === "") {
    end -= 1;
  }

  return values.slice(0, end);
}

function resolveFrom(cwd: string, targetPath: string): string {
  return resolve(cwd, targetPath);
}

async function main(): Promise<void> {
  process.exitCode = await runCli(process.argv.slice(2));
}

if (
  process.argv[1] &&
  realpathSync.native(resolve(process.argv[1])) === realpathSync.native(fileURLToPath(import.meta.url))
) {
  await main();
}
