import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { basename, join, resolve } from "node:path";

import { Workbook } from "../dist/src/index.js";

const inputArg = process.argv[2];

if (!inputArg) {
  console.error("Usage: node scripts/validate-roundtrip.mjs <input.xlsx> [output.xlsx]");
  process.exit(1);
}

const inputPath = resolve(inputArg);
const tempRoot = await mkdtemp(join(tmpdir(), "xlsx-ts-validate-"));
const outputPath = process.argv[3]
  ? resolve(process.argv[3])
  : join(tempRoot, `${basename(inputPath, ".xlsx")}.roundtrip.xlsx`);

try {
  const document = await Workbook.open(inputPath);
  await document.save(outputPath);

  const source = await Workbook.open(inputPath);
  const roundtrip = await Workbook.open(outputPath);
  const sourceEntries = toEntryMap(source.toEntries());
  const roundtripEntries = toEntryMap(roundtrip.toEntries());
  const sourceKeys = [...sourceEntries.keys()].sort();
  const roundtripKeys = [...roundtripEntries.keys()].sort();

  assert.deepEqual(roundtripKeys, sourceKeys, "zip entry list changed after roundtrip");

  const diffs = [];

  for (const key of sourceKeys) {
    const left = sourceEntries.get(key);
    const right = roundtripEntries.get(key);

    if (!left || !right || Buffer.compare(left, right) !== 0) {
      diffs.push(key);
    }
  }

  if (diffs.length > 0) {
    console.error(
      JSON.stringify(
        {
          input: inputPath,
          output: outputPath,
          entries: sourceKeys.length,
          diffs,
        },
        null,
        2,
      ),
    );
    process.exit(2);
  }

  console.log(
    JSON.stringify(
      {
        input: inputPath,
        output: outputPath,
        entries: sourceKeys.length,
        diffs,
      },
      null,
      2,
    ),
  );
} finally {
  if (!process.argv[3]) {
    await rm(tempRoot, { recursive: true, force: true });
  }
}

function toEntryMap(entries) {
  return new Map(entries.map((entry) => [entry.path, Buffer.from(entry.data)]));
}
