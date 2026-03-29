import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

test("package metadata exposes JS and type entrypoints for legacy TS resolution", async () => {
  const packageJson = JSON.parse(await readFile(new URL("../package.json", import.meta.url), "utf8")) as {
    bin?: Record<string, string>;
    exports?: {
      ".": {
        import?: string;
        types?: string;
      };
    };
    main?: string;
    types?: string;
  };

  assert.equal(packageJson.main, "./dist/src/index.js");
  assert.equal(packageJson.types, "./dist/src/index.d.ts");
  assert.equal(packageJson.bin?.["xlsx-ts"], "./dist/src/cli.js");
  assert.equal(packageJson.exports?.["."].import, "./dist/src/index.js");
  assert.equal(packageJson.exports?.["."].types, "./dist/src/index.d.ts");
});
