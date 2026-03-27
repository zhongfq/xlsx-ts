export interface ArchiveEntry {
  path: string;
  data: Uint8Array;
}

export type CellValue = string | number | boolean | null;
export type CellType = "missing" | "blank" | "string" | "number" | "boolean" | "formula";

export interface SetFormulaOptions {
  cachedValue?: CellValue;
}
