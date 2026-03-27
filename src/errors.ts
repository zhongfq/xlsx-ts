export class XlsxError extends Error {
  constructor(message: string, options?: { cause?: unknown }) {
    super(message);
    this.name = "XlsxError";
    if (options && "cause" in options) {
      (this as Error & { cause?: unknown }).cause = options.cause;
    }
  }
}

export class CommandError extends XlsxError {
  constructor(command: string, code: number, stderr: string) {
    super(`Command failed: ${command} (exit ${code})${stderr ? `\n${stderr}` : ""}`);
    this.name = "CommandError";
  }
}
