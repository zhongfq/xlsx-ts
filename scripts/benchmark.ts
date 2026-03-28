import { readFile } from "node:fs/promises";
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

export interface BenchmarkComparison {
  absoluteGapMs: number;
  ratioVsDense: number;
  nonNullMatches: boolean;
}

export interface BenchmarkBaseline {
  expectedNonNull: number;
  maxAbsoluteGapMs?: number;
  maxRatioVsDense?: number;
  maxXlsxTsAverageMs?: number;
}

export async function runCompareBenchmark(options: {
  filePath?: string;
  iterations?: number;
} = {}): Promise<{
  file: string;
  iterations: number;
  results: BenchmarkResult[];
  comparison: BenchmarkComparison;
}> {
  const filePath = options.filePath ?? resolve(process.cwd(), "res/monster.xlsx");
  const iterations = options.iterations ?? 3;
  const local = await benchmark("xlsx-ts", iterations, () => benchmarkLocalWorkbook(filePath));
  const dense = await benchmark("xlsx dense", iterations, () => benchmarkXlsxDense(filePath));

  return {
    file: filePath,
    iterations,
    results: [local, dense],
    comparison: buildComparison(local, dense),
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
  const { filePathArg, iterationsArg, baselinePathArg } = parseCliArgs(process.argv.slice(2));
  const result = await runCompareBenchmark({
    filePath: filePathArg ? resolve(process.cwd(), filePathArg) : undefined,
    iterations: iterationsArg ? Number(iterationsArg) : undefined,
  });

  if (baselinePathArg) {
    const baselinePath = resolve(process.cwd(), baselinePathArg);
    const baseline = JSON.parse(await readFile(baselinePath, "utf8")) as BenchmarkBaseline;
    const failures = validateAgainstBaseline(result, baseline);
    const output = {
      ...result,
      check: {
        ok: failures.length === 0,
        baseline: baselinePath,
        failures,
      },
    };

    if (failures.length > 0) {
      console.error(JSON.stringify(output, null, 2));
      process.exitCode = 1;
      return;
    }

    console.log(JSON.stringify(output, null, 2));
    return;
  }

  console.log(JSON.stringify(result, null, 2));
}

if (process.argv[1] && fileURLToPath(import.meta.url) === resolve(process.argv[1])) {
  await main();
}

function buildComparison(local: BenchmarkResult, dense: BenchmarkResult): BenchmarkComparison {
  return {
    absoluteGapMs: Number((local.averageMs - dense.averageMs).toFixed(1)),
    ratioVsDense: Number((local.averageMs / dense.averageMs).toFixed(3)),
    nonNullMatches: local.nonNull === dense.nonNull,
  };
}

function parseCliArgs(args: string[]): {
  filePathArg?: string;
  iterationsArg?: string;
  baselinePathArg?: string;
} {
  const positional: string[] = [];
  let baselinePathArg: string | undefined;

  for (let index = 0; index < args.length; index += 1) {
    const argument = args[index];
    if (argument === "--check") {
      baselinePathArg = args[index + 1];
      if (!baselinePathArg) {
        throw new Error("Missing baseline path after --check");
      }
      index += 1;
      continue;
    }

    positional.push(argument);
  }

  return {
    filePathArg: positional[0],
    iterationsArg: positional[1],
    baselinePathArg,
  };
}

function validateAgainstBaseline(
  result: Awaited<ReturnType<typeof runCompareBenchmark>>,
  baseline: BenchmarkBaseline,
): string[] {
  const failures: string[] = [];
  const local = result.results.find((candidate) => candidate.label === "xlsx-ts");

  if (!local) {
    return ["Missing xlsx-ts benchmark result"];
  }

  if (local.nonNull !== baseline.expectedNonNull) {
    failures.push(`Expected nonNull=${baseline.expectedNonNull}, got ${local.nonNull}`);
  }

  if (!result.comparison.nonNullMatches) {
    failures.push("xlsx-ts and xlsx dense produced different nonNull counts");
  }

  if (
    baseline.maxAbsoluteGapMs !== undefined &&
    result.comparison.absoluteGapMs > baseline.maxAbsoluteGapMs
  ) {
    failures.push(
      `Absolute gap ${result.comparison.absoluteGapMs}ms exceeded ${baseline.maxAbsoluteGapMs}ms`,
    );
  }

  if (
    baseline.maxRatioVsDense !== undefined &&
    result.comparison.ratioVsDense > baseline.maxRatioVsDense
  ) {
    failures.push(
      `Ratio ${result.comparison.ratioVsDense} exceeded ${baseline.maxRatioVsDense}`,
    );
  }

  if (
    baseline.maxXlsxTsAverageMs !== undefined &&
    local.averageMs > baseline.maxXlsxTsAverageMs
  ) {
    failures.push(
      `xlsx-ts average ${local.averageMs}ms exceeded ${baseline.maxXlsxTsAverageMs}ms`,
    );
  }

  return failures;
}
