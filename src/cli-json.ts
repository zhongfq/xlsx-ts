import type { CellValue, SheetVisibility } from "./types.js";

export type Writer = (chunk: string) => void;
export type CellRecord = Record<string, CellValue>;

export function resolveMatchValue(value?: string, text?: string): CellValue {
  const actionCount = Number(value !== undefined) + Number(text !== undefined);
  if (actionCount !== 1) {
    throw new Error("Exactly one of --value or --text is required");
  }

  return text !== undefined ? text : parseJsonCellValue(value!, "--value");
}

export function resolveUpsertMatchValue(
  record: CellRecord,
  field: string,
  matchValue?: string,
  matchText?: string,
): CellValue {
  const overrideCount = Number(matchValue !== undefined) + Number(matchText !== undefined);
  if (overrideCount > 1) {
    throw new Error("Use either --match-value or --match-text, not both");
  }

  if (matchText !== undefined) {
    return matchText;
  }

  if (matchValue !== undefined) {
    return parseJsonCellValue(matchValue, "--match-value");
  }

  if (!Object.hasOwn(record, field)) {
    throw new Error(`Record is missing the match field: ${field}`);
  }

  return record[field] ?? null;
}

export function parseJsonCellValue(source: string, label: string): CellValue {
  return assertCellValue(parseJsonDocument(source, label), label);
}

export function parseJsonDocument(source: string, label: string): unknown {
  try {
    return JSON.parse(source);
  } catch (error) {
    throw new Error(`Invalid JSON in ${label}: ${formatError(error)}`);
  }
}

export function writeJson(writer: Writer, value: unknown): void {
  writer(`${JSON.stringify(value, null, 2)}\n`);
}

export function assertRecord(value: unknown, label: string): Record<string, unknown> {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new Error(`Expected ${label} to be an object`);
  }

  return value as Record<string, unknown>;
}

export function assertArray(value: unknown, label: string): unknown[] {
  if (!Array.isArray(value)) {
    throw new Error(`Expected ${label} to be an array`);
  }

  return value;
}

export function assertString(value: unknown, label: string): string {
  if (typeof value !== "string" || value.length === 0) {
    throw new Error(`Expected ${label} to be a non-empty string`);
  }

  return value;
}

export function optionalString(value: unknown, label: string): string | undefined {
  if (value === undefined) {
    return undefined;
  }

  return assertString(value, label);
}

export function assertNullableString(value: unknown, label: string): string | null {
  if (value === null) {
    return null;
  }

  return assertString(value, label);
}

export function optionalPositiveInteger(value: unknown, label: string): number | undefined {
  if (value === undefined) {
    return undefined;
  }

  return assertPositiveInteger(value, label);
}

export function assertPositiveInteger(value: unknown, label: string): number {
  if (typeof value !== "number" || !Number.isInteger(value) || value <= 0) {
    throw new Error(`Expected ${label} to be a positive integer`);
  }

  return value;
}

export function assertPositiveIntegerArray(value: unknown, label: string): number[] {
  const values = assertArray(value, label);
  return values.map((item, index) => assertPositiveInteger(item, `${label}[${index}]`));
}

export function assertCellValue(value: unknown, label: string): CellValue {
  if (
    value === null ||
    typeof value === "string" ||
    typeof value === "number" ||
    typeof value === "boolean"
  ) {
    return value;
  }

  throw new Error(`Expected ${label} to be a string, number, boolean, or null`);
}

export function assertCellRecord(value: unknown, label: string): CellRecord {
  const record = assertRecord(value, label);
  const next: CellRecord = {};

  for (const [key, item] of Object.entries(record)) {
    next[key] = assertCellValue(item, `${label}.${key}`);
  }

  return next;
}

export function assertCellRecordArray(value: unknown, label: string): CellRecord[] {
  const values = assertArray(value, label);
  return values.map((item, index) => assertCellRecord(item, `${label}[${index}]`));
}

export function assertStringArray(value: unknown, label: string): string[] {
  const values = assertArray(value, label);
  return values.map((item, index) => {
    if (typeof item !== "string") {
      throw new Error(`Expected ${label}[${index}] to be a string`);
    }

    return item;
  });
}

export function assertSheetVisibility(value: unknown, label: string): SheetVisibility {
  if (value === "visible" || value === "hidden" || value === "veryHidden") {
    return value;
  }

  throw new Error(`Expected ${label} to be visible, hidden, or veryHidden`);
}

export function parseJsonCellRecord(source: string, label: string): CellRecord {
  return assertCellRecord(parseJsonDocument(source, label), label);
}

export function parseJsonCellRecordArray(source: string, label: string): CellRecord[] {
  return assertCellRecordArray(parseJsonDocument(source, label), label);
}

export function parseJsonStringArray(source: string, label: string): string[] {
  return assertStringArray(parseJsonDocument(source, label), label);
}

export function formatError(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
