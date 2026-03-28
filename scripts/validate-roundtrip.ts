import { resolve } from "node:path";

import { validateRoundtripFile } from "../src/index.ts";

const inputArg = process.argv[2];

if (!inputArg) {
  console.error("Usage: node scripts/validate-roundtrip.ts <input.xlsx> [output.xlsx]");
  process.exit(1);
}

const inputPath = resolve(inputArg);
const result = await validateRoundtripFile(inputPath, process.argv[3] ? resolve(process.argv[3]) : undefined);

if (!result.ok) {
  console.error(JSON.stringify(result, null, 2));
  process.exit(2);
}

console.log(JSON.stringify(result, null, 2));
