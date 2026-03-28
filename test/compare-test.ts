import xlsx from "xlsx";

import { Workbook } from "../src/index.js";

const FILE_PATH = "/Users/codetypes/Desktop/Github/xlsx-ts/res/monster.xlsx";
const ITERATIONS = 3;

interface BenchmarkResult {
  label: string;
  runs: number[];
  averageMs: number;
  nonNull: number;
}

const local = await benchmark("xlsx-ts", benchmarkLocalWorkbook);
const dense = await benchmark("xlsx dense", benchmarkXlsxDense);

console.log(JSON.stringify({ file: FILE_PATH, iterations: ITERATIONS, results: [local, dense] }, null, 2));

async function benchmark(
  label: string,
  runOnce: () => Promise<number> | number,
): Promise<BenchmarkResult> {
  const runs: number[] = [];
  let nonNull = 0;

  for (let index = 0; index < ITERATIONS; index += 1) {
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

async function benchmarkLocalWorkbook(): Promise<number> {
  const workbook = await Workbook.open(FILE_PATH);
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

async function benchmarkXlsxDense(): Promise<number> {
  const workbook = xlsx.readFile(FILE_PATH, {
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
