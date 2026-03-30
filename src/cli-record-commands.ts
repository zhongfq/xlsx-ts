import { resolve } from "node:path";

import { Command } from "commander";

import {
  parseJsonCellRecord,
  parseJsonCellRecordArray,
  parseJsonStringArray,
  writeJson,
} from "./cli-json.js";
import type { CellRecord, Writer } from "./cli-json.js";
import { Workbook } from "./workbook.js";

export interface RecordCommandIo {
  cwd: string;
  stdout: Writer;
}

export interface RecordCommandHelpers {
  parsePositiveInteger: (value: string) => number;
  resolveOutputPath: (
    inputPath: string,
    options: {
      inPlace: boolean;
      output?: string;
    },
  ) => string;
}

export function registerRecordCommands(
  program: Command,
  io: RecordCommandIo,
  helpers: RecordCommandHelpers,
): void {
  const { parsePositiveInteger, resolveOutputPath } = helpers;

  program
    .command("records")
    .argument("<file>", "input xlsx file")
    .requiredOption("--sheet <name>", "sheet name")
    .option("--header-row <row>", "header row used for record mapping", parsePositiveInteger, 1)
    .action(async (file: string, options: { headerRow: number; sheet: string }) => {
      const result = await getRecords(resolveFrom(io.cwd, file), options.sheet, options.headerRow);
      writeJson(io.stdout, result);
    });

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

function resolveFrom(cwd: string, targetPath: string): string {
  return resolve(cwd, targetPath);
}
