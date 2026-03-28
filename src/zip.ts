import { mkdir, readFile, writeFile } from "node:fs/promises";
import { dirname, join } from "node:path";

import { unzipSync, zipSync } from "fflate";

import type { ArchiveEntry } from "./types.js";

export class Zip {
  async readArchive(filePath: string): Promise<ArchiveEntry[]> {
    const archiveData = await readFile(filePath);
    const entriesByPath = unzipSync(new Uint8Array(archiveData));

    return Object.keys(entriesByPath)
      .sort((left, right) => left.localeCompare(right))
      .map((path) => ({
        path,
        data: entriesByPath[path],
      }));
  }

  async writeArchive(filePath: string, entries: ArchiveEntry[]): Promise<void> {
    const destinationDirectory = dirname(filePath);
    if (destinationDirectory !== ".") {
      await mkdir(destinationDirectory, { recursive: true });
    }

    const zipped = zipSync(
      Object.fromEntries(entries.map((entry) => [entry.path, new Uint8Array(entry.data)])),
      { level: 6 },
    );
    await writeFile(filePath, zipped);
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
