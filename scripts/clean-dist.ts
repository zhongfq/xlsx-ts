import { rm } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const distPath = resolve(repoRoot, "dist");

await rm(distPath, { recursive: true, force: true });
