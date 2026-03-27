export interface ArchiveEntry {
  path: string;
  data: Uint8Array;
}

export type CellValue = string | number | boolean | null;

export interface SetFormulaOptions {
  cachedValue?: CellValue;
}
