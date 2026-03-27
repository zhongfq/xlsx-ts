export function basenamePosix(path: string): string {
  const parts = path.split("/");
  return parts[parts.length - 1] ?? path;
}

export function dirnamePosix(path: string): string {
  const index = path.lastIndexOf("/");
  return index === -1 ? "" : path.slice(0, index);
}

export function resolvePosix(baseDir: string, target: string): string {
  const stack = baseDir ? baseDir.split("/") : [];

  for (const part of target.split("/")) {
    if (!part || part === ".") {
      continue;
    }

    if (part === "..") {
      stack.pop();
      continue;
    }

    stack.push(part);
  }

  return stack.join("/");
}
