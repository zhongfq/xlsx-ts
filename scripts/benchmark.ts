import { resolve } from "node:path";
import { fileURLToPath } from "node:url";

import xlsx from "xlsx";

import { Workbook } from "../src/index.js";

export interface BenchmarkResult {
  label: string;
  runs: number[];
  averageMs: number;
  nonNull: number;
}

export async function runCompareBenchmark(options: {
  filePath?: string;
  iterations?: number;
} = {}): Promise<{
  file: string;
  iterations: number;
  results: BenchmarkResult[];
}> {
  const filePath = options.filePath ?? resolve(process.cwd(), "res/monster.xlsx");
  const iterations = options.iterations ?? 3;
  const local = await benchmark("xlsx-ts", iterations, () => benchmarkLocalWorkbook(filePath));
  const dense = await benchmark("xlsx dense", iterations, () => benchmarkXlsxDense(filePath));

  return {
    file: filePath,
    iterations,
    results: [local, dense],
  };
}

async function benchmark(
  label: string,
  iterations: number,
  runOnce: () => Promise<number> | number,
): Promise<BenchmarkResult> {
  const runs: number[] = [];
  let nonNull = 0;

  for (let index = 0; index < iterations; index += 1) {
    const startedAt = performance.now();
    nonNull = await runOnce();
    runs.push(Number((performance.now() - startedAt).toFixed(1)));
  }

  return {
    label,
    runs,
    averageMs: Number((runs.reduce((sum, value) => sum + value, 0) / runs.length).toFixed(1)),
    nonNull,
  };
}

async function benchmarkLocalWorkbook(filePath: string): Promise<number> {
  const workbook = await Workbook.open(filePath);
  let nonNull = 0;

  for (const sheet of workbook.getSheets()) {
    const rowCount = sheet.rowCount;
    const columnCount = sheet.columnCount;

    for (let rowNumber = 1; rowNumber <= rowCount; rowNumber += 1) {
      for (let columnNumber = 1; columnNumber <= columnCount; columnNumber += 1) {
        const cell = sheet.getCell(rowNumber, columnNumber);
        if (cell !== null) {
          cell.toString();
          nonNull += 1;
        }
      }
    }
  }

  return nonNull;
}

async function benchmarkXlsxDense(filePath: string): Promise<number> {
  const workbook = xlsx.readFile(filePath, {
    dense: true,
    cellHTML: false,
    cellFormula: false,
    cellText: false,
    raw: true,
  });
  let nonNull = 0;

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const ref = sheet["!ref"];
    if (!ref) {
      continue;
    }

    const range = xlsx.utils.decode_range(ref);

    for (let rowIndex = range.s.r; rowIndex <= range.e.r; rowIndex += 1) {
      const row = sheet[rowIndex] ?? [];

      for (let columnIndex = range.s.c; columnIndex <= range.e.c; columnIndex += 1) {
        const cell = row[columnIndex];
        if (cell?.v !== undefined && cell.v !== null) {
          cell.v.toString();
          nonNull += 1;
        }
      }
    }
  }

  return nonNull;
}

async function main(): Promise<void> {
  const filePathArg = process.argv[2];
  const iterationsArg = process.argv[3];
  const result = await runCompareBenchmark({
    filePath: filePathArg ? resolve(process.cwd(), filePathArg) : undefined,
    iterations: iterationsArg ? Number(iterationsArg) : undefined,
  });

  console.log(JSON.stringify(result, null, 2));
}

if (process.argv[1] && fileURLToPath(import.meta.url) === resolve(process.argv[1])) {
  await main();
}
