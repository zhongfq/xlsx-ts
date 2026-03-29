import assert from "node:assert/strict";
import { mkdtemp, readdir } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, posix, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { spawnSync } from "node:child_process";

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");

const expectedPackPaths = new Set([
  "LICENSE",
  "README.md",
  "README.zh.md",
  "package.json",
  ...(await getExpectedDistPaths()),
]);

const actualPackPaths = new Set(await getPackPaths());

const missing = [...expectedPackPaths].filter((path) => !actualPackPaths.has(path)).sort();
const unexpected = [...actualPackPaths].filter((path) => !expectedPackPaths.has(path)).sort();

if (missing.length > 0 || unexpected.length > 0) {
  if (missing.length > 0) {
    console.error(`Missing package files:\n${missing.join("\n")}`);
  }

  if (unexpected.length > 0) {
    console.error(`Unexpected package files:\n${unexpected.join("\n")}`);
  }

  process.exitCode = 1;
}

async function getExpectedDistPaths() {
  const sourceFiles = await listFiles(resolve(repoRoot, "src"));
  const expected = new Set();

  for (const sourceFile of sourceFiles) {
    if (sourceFile.endsWith(".d.ts")) {
      expected.add(`dist/${sourceFile}`);
      continue;
    }

    if (sourceFile.endsWith(".ts")) {
      const basePath = sourceFile.slice(0, -3);
      expected.add(`dist/${basePath}.d.ts`);
      expected.add(`dist/${basePath}.js`);
    }
  }

  return [...expected].sort();
}

async function listFiles(directoryPath, root = directoryPath) {
  const entries = await readdir(directoryPath, { withFileTypes: true });
  const files = [];

  for (const entry of entries.sort((left, right) => left.name.localeCompare(right.name))) {
    const absolutePath = resolve(directoryPath, entry.name);

    if (entry.isDirectory()) {
      files.push(...(await listFiles(absolutePath, root)));
      continue;
    }

    files.push(posix.join("src", absolutePath.slice(root.length + 1).replaceAll("\\", "/")));
  }

  return files;
}

async function getPackPaths() {
  const cacheDirectory = await mkdtemp(resolve(tmpdir(), "xlsx-ts-npm-cache-"));
  const npmCommand = process.platform === "win32" ? "npm.cmd" : "npm";
  const result = spawnSync(npmCommand, ["pack", "--json", "--dry-run"], {
    cwd: repoRoot,
    encoding: "utf8",
    env: {
      ...process.env,
      npm_config_cache: cacheDirectory,
    },
  });

  if (result.status !== 0) {
    process.stderr.write(result.stderr);
    process.exit(result.status ?? 1);
  }

  const payload = JSON.parse(result.stdout);
  assert.ok(Array.isArray(payload) && payload.length === 1, "Expected a single npm pack result");

  return payload[0].files.map((entry) => entry.path).sort();
}
