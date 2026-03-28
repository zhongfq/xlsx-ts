export interface ArchiveEntry {
  path: string;
  data: Uint8Array;
}

export type CellValue = string | number | boolean | null;
export type CellType = "missing" | "blank" | "string" | "number" | "boolean" | "formula";

export interface DefinedName {
  hidden: boolean;
  name: string;
  scope: string | null;
  value: string;
}

export interface SetFormulaOptions {
  cachedValue?: CellValue;
}

export interface SetDefinedNameOptions {
  scope?: string;
}
