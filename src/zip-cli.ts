import { mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";

import type { ArchiveEntry } from "./types.js";
import { runCommand } from "./utils/exec.js";

export class CliZipAdapter {
  async readArchive(filePath: string): Promise<ArchiveEntry[]> {
    const listOutput = await runCommand("python3", [
      "-c",
      "import sys, zipfile; z = zipfile.ZipFile(sys.argv[1]); sys.stdout.write('\\n'.join(i.filename for i in z.infolist() if not i.is_dir()))",
      filePath,
    ]);
    const entryNames = new TextDecoder()
      .decode(listOutput)
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter((line) => line.length > 0);

    const entries: ArchiveEntry[] = [];

    for (const entryPath of entryNames) {
      const data = await runCommand("python3", [
        "-c",
        "import sys, zipfile; sys.stdout.buffer.write(zipfile.ZipFile(sys.argv[1]).read(sys.argv[2]))",
        filePath,
        entryPath,
      ]);
      entries.push({ path: entryPath, data });
    }

    return entries;
  }

  async writeArchive(filePath: string, entries: ArchiveEntry[]): Promise<void> {
    const destination = resolve(process.cwd(), filePath);
    const tempRoot = await mkdtemp(join(tmpdir(), "xlsx-ts-"));
    const stagingDir = join(tempRoot, "archive");

    await mkdir(stagingDir, { recursive: true });

    try {
      for (const entry of entries) {
        const absolutePath = join(stagingDir, entry.path);
        await mkdir(dirname(absolutePath), { recursive: true });
        await writeFile(absolutePath, entry.data);
      }

      await rm(destination, { force: true });
      await runCommand("zip", ["-X", "-q", "-r", destination, "."], { cwd: stagingDir });
    } finally {
      await rm(tempRoot, { recursive: true, force: true });
    }
  }

  async readDirectoryEntries(directoryPath: string, root = directoryPath): Promise<ArchiveEntry[]> {
    const collected: ArchiveEntry[] = [];

    const stack = [directoryPath];

    while (stack.length > 0) {
      const current = stack.pop();
      if (!current) {
        continue;
      }

      const children = await readDirNames(current);

      for (const child of children) {
        const absolutePath = join(current, child.name);

        if (child.isDirectory) {
          stack.push(absolutePath);
          continue;
        }

        const relativePath = absolutePath.slice(root.length + 1).replaceAll("\\", "/");
        collected.push({
          path: relativePath,
          data: await readFile(absolutePath),
        });
      }
    }

    collected.sort((left, right) => left.path.localeCompare(right.path));
    return collected;
  }
}

async function readDirNames(
  directoryPath: string,
): Promise<Array<{ name: string; isDirectory: boolean }>> {
  const names = await (await import("node:fs/promises")).readdir(directoryPath);
  const rows: Array<{ name: string; isDirectory: boolean }> = [];

  for (const name of names as string[]) {
    const absolutePath = join(directoryPath, name);
    const stats = await (await import("node:fs/promises")).stat(absolutePath);
    rows.push({ name, isDirectory: stats.isDirectory() });
  }

  rows.sort((left, right) => left.name.localeCompare(right.name));
  return rows;
}
