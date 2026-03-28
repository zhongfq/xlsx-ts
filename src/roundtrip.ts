import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { basename, join, resolve } from "node:path";

import { Workbook } from "./workbook.js";

export interface RoundtripValidationResult {
  input: string;
  output: string | null;
  entries: number;
  diffs: string[];
  ok: boolean;
}

export async function validateRoundtripFile(
  inputPath: string,
  outputPath?: string,
): Promise<RoundtripValidationResult> {
  const resolvedInputPath = resolve(inputPath);
  const tempRoot = outputPath ? null : await mkdtemp(join(tmpdir(), "xlsx-ts-validate-"));
  const resolvedOutputPath = outputPath
    ? resolve(outputPath)
    : join(tempRoot!, `${basename(resolvedInputPath, ".xlsx")}.roundtrip.xlsx`);

  try {
    const document = await Workbook.open(resolvedInputPath);
    await document.save(resolvedOutputPath);

    const source = await Workbook.open(resolvedInputPath);
    const roundtrip = await Workbook.open(resolvedOutputPath);
    const sourceEntries = toEntryMap(source.toEntries());
    const roundtripEntries = toEntryMap(roundtrip.toEntries());
    const sourceKeys = [...sourceEntries.keys()].sort();
    const roundtripKeys = [...roundtripEntries.keys()].sort();
    const diffs: string[] = [];

    if (
      sourceKeys.length !== roundtripKeys.length ||
      sourceKeys.some((key, index) => key !== roundtripKeys[index])
    ) {
      diffs.push("__entry_list__");
    }

    for (const key of sourceKeys) {
      const left = sourceEntries.get(key);
      const right = roundtripEntries.get(key);

      if (!left || !right || Buffer.compare(left, right) !== 0) {
        diffs.push(key);
      }
    }

    return {
      input: resolvedInputPath,
      output: outputPath ? resolvedOutputPath : null,
      entries: sourceKeys.length,
      diffs,
      ok: diffs.length === 0,
    };
  } finally {
    if (tempRoot) {
      await rm(tempRoot, { recursive: true, force: true });
    }
  }
}

function toEntryMap(entries: Array<{ path: string; data: Uint8Array }>): Map<string, Uint8Array> {
  return new Map(entries.map((entry) => [entry.path, Buffer.from(entry.data)]));
}
