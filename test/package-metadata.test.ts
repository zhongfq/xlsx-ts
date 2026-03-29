import test from "node:test";
import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdir, mkdtemp, readFile, rm, stat } from "node:fs/promises";
import { join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

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

test("packed tarball exposes runnable entrypoints without stale build artifacts", async () => {
  const repoRoot = resolve(fileURLToPath(new URL("..", import.meta.url)));
  const tempRoot = await mkdtemp(join(repoRoot, ".pack-smoke-"));
  const npmCache = join(tempRoot, "npm-cache");

  try {
    runChecked(repoRoot, npmCommand(), ["run", "build"], {
      npm_config_cache: npmCache,
    });

    const packResult = runChecked(repoRoot, npmCommand(), [
      "pack",
      "--json",
      "--pack-destination",
      tempRoot,
    ], {
      npm_config_cache: npmCache,
    });
    const packPayload = JSON.parse(packResult.stdout) as Array<{ filename: string }>;
    assert.equal(packPayload.length, 1);

    const tarballPath = join(tempRoot, packPayload[0]!.filename);
    const extractRoot = join(tempRoot, "extract");
    await mkdir(extractRoot, { recursive: true });
    runChecked(repoRoot, "tar", ["-xzf", tarballPath, "-C", extractRoot]);

    const packageRoot = join(extractRoot, "package");
    await assertPathExists(join(packageRoot, "dist/src/index.js"));
    await assertPathExists(join(packageRoot, "dist/src/index.d.ts"));
    await assertPathExists(join(packageRoot, "dist/src/cli.js"));

    const cliResult = runChecked(packageRoot, process.execPath, ["dist/src/cli.js", "--help"]);
    assert.match(cliResult.stdout, /Usage: xlsx-ts \[options\] \[command\]/);

    runChecked(packageRoot, process.execPath, [
      "--input-type=module",
      "-e",
      'const mod = await import("./dist/src/index.js"); if (typeof mod.Workbook !== "function") process.exit(1);',
    ]);

    await assert.rejects(stat(join(packageRoot, "dist/src/rich-text.js")));
    await assert.rejects(stat(join(packageRoot, "dist/src/xlsx-document.js")));
    await assert.rejects(stat(join(packageRoot, "dist/src/zip-cli.js")));
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

function runChecked(
  cwd: string,
  command: string,
  args: string[],
  extraEnv: NodeJS.ProcessEnv = {},
): { stdout: string; stderr: string } {
  const result = spawnSync(command, args, {
    cwd,
    encoding: "utf8",
    env: {
      ...process.env,
      ...extraEnv,
    },
  });

  if (result.status !== 0) {
    throw new Error(
      [
        `Command failed: ${command} ${args.join(" ")}`,
        result.stdout.trim(),
        result.stderr.trim(),
      ]
        .filter((part) => part.length > 0)
        .join("\n"),
    );
  }

  return {
    stdout: result.stdout,
    stderr: result.stderr,
  };
}

function npmCommand(): string {
  return process.platform === "win32" ? "npm.cmd" : "npm";
}

async function assertPathExists(path: string): Promise<void> {
  await stat(path);
}
