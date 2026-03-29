import { readFile } from "node:fs/promises";
import { resolve } from "node:path";
import { fileURLToPath } from "node:url";

import { Workbook } from "../src/index.js";

export interface BenchmarkResult {
  runs: number[];
  averageMs: number;
  nonNull: number;
}

export interface BenchmarkBaseline {
  expectedNonNull: number;
  maxAverageMs?: number;
}

export async function runBenchmark(options: {
  filePath?: string;
  iterations?: number;
} = {}): Promise<{
  file: string;
  iterations: number;
  result: BenchmarkResult;
}> {
  const filePath = options.filePath ?? resolve(process.cwd(), "res/monster.xlsx");
  const iterations = options.iterations ?? 3;
  const result = await benchmark(iterations, () => benchmarkLocalWorkbook(filePath));

  return {
    file: filePath,
    iterations,
    result,
  };
}

async function benchmark(
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

async function main(): Promise<void> {
  const { filePathArg, iterationsArg, baselinePathArg } = parseCliArgs(process.argv.slice(2));
  const result = await runBenchmark({
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
  result: Awaited<ReturnType<typeof runBenchmark>>,
  baseline: BenchmarkBaseline,
): string[] {
  const failures: string[] = [];
  const local = result.result;

  if (local.nonNull !== baseline.expectedNonNull) {
    failures.push(`Expected nonNull=${baseline.expectedNonNull}, got ${local.nonNull}`);
  }

  if (baseline.maxAverageMs !== undefined && local.averageMs > baseline.maxAverageMs) {
    failures.push(`Average ${local.averageMs}ms exceeded ${baseline.maxAverageMs}ms`);
  }

  return failures;
}
