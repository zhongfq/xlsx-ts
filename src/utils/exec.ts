import { spawn } from "node:child_process";

import { CommandError } from "../errors.js";

export async function runCommand(
  command: string,
  args: string[],
  options: {
    cwd?: string;
    stdin?: Uint8Array;
  } = {},
): Promise<Uint8Array> {
  const child = spawn(command, args, {
    cwd: options.cwd,
    stdio: ["pipe", "pipe", "pipe"],
  });

  const stdoutChunks: Uint8Array[] = [];
  const stderrChunks: Uint8Array[] = [];

  child.stdout.on("data", (chunk: Uint8Array) => {
    stdoutChunks.push(chunk);
  });

  child.stderr.on("data", (chunk: Uint8Array) => {
    stderrChunks.push(chunk);
  });

  child.stdin.end(options.stdin);

  const exitCode = await new Promise<number>((resolve, reject) => {
    child.on("error", reject);
    child.on("close", (code: number | null) => resolve(code ?? 0));
  });

  if (exitCode !== 0) {
    throw new CommandError(
      [command, ...args].join(" "),
      exitCode,
      new TextDecoder().decode(Buffer.concat(stderrChunks)).trim(),
    );
  }

  return Buffer.concat(stdoutChunks);
}
