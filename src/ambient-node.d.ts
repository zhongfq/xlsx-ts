declare module "node:child_process" {
  export function spawn(command: string, args?: string[], options?: Record<string, unknown>): {
    stdout: {
      on(event: string, listener: (chunk: Uint8Array) => void): void;
    };
    stderr: {
      on(event: string, listener: (chunk: Uint8Array) => void): void;
    };
    stdin: {
      end(chunk?: Uint8Array): void;
    };
    on(event: string, listener: (...args: any[]) => void): void;
  };
}

declare module "node:fs/promises" {
  export function mkdir(path: string, options?: Record<string, unknown>): Promise<void>;
  export function mkdtemp(prefix: string): Promise<string>;
  export function readFile(path: string): Promise<Uint8Array>;
  export function readFile(path: string, encoding: string): Promise<string>;
  export function readdir(path: string, options?: Record<string, unknown>): Promise<Array<any>>;
  export function rm(path: string, options?: Record<string, unknown>): Promise<void>;
  export function stat(path: string): Promise<{
    isDirectory(): boolean;
  }>;
  export function writeFile(path: string, data: string | Uint8Array): Promise<void>;
}

declare module "node:os" {
  export function tmpdir(): string;
}

declare module "node:path" {
  export function basename(path: string, suffix?: string): string;
  export function dirname(path: string): string;
  export function join(...parts: string[]): string;
  export function resolve(...parts: string[]): string;
}

declare module "node:url" {
  export function fileURLToPath(url: string): string;
}

declare const Buffer: {
  compare(left: Uint8Array, right: Uint8Array): number;
  concat(chunks: ReadonlyArray<Uint8Array>): Uint8Array;
  from(data: string | Uint8Array, encoding?: string): Uint8Array;
};

declare const process: {
  argv: string[];
  cwd(): string;
  exitCode?: number;
  stderr: {
    write(chunk: string): void;
  };
  stdout: {
    write(chunk: string): void;
  };
};
