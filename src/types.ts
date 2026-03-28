export interface ArchiveEntry {
  path: string;
  data: Uint8Array;
}

export type CellValue = string | number | boolean | null;
export type CellType = "missing" | "blank" | "string" | "number" | "boolean" | "formula";
export type SheetVisibility = "visible" | "hidden" | "veryHidden";

export interface CellSnapshot {
  exists: boolean;
  formula: string | null;
  rawType: string | null;
  styleId: number | null;
  type: CellType;
  value: CellValue;
}

export interface CellEntry extends CellSnapshot {
  address: string;
  rowNumber: number;
  columnNumber: number;
}

export interface DefinedName {
  hidden: boolean;
  name: string;
  scope: string | null;
  value: string;
}

export interface Hyperlink {
  address: string;
  target: string;
  tooltip: string | null;
  type: "external" | "internal";
}

export interface FreezePane {
  columnCount: number;
  rowCount: number;
  topLeftCell: string;
  activePane: "bottomLeft" | "topRight" | "bottomRight" | null;
}

export interface SheetSelection {
  activeCell: string | null;
  range: string | null;
  pane: "bottomLeft" | "topRight" | "bottomRight" | null;
}

export interface DataValidation {
  range: string;
  type: string | null;
  operator: string | null;
  allowBlank: boolean | null;
  showInputMessage: boolean | null;
  showErrorMessage: boolean | null;
  showDropDown: boolean | null;
  errorStyle: string | null;
  errorTitle: string | null;
  error: string | null;
  promptTitle: string | null;
  prompt: string | null;
  imeMode: string | null;
  formula1: string | null;
  formula2: string | null;
}

export interface SetFormulaOptions {
  cachedValue?: CellValue;
}

export interface SetDefinedNameOptions {
  scope?: string;
}

export interface SetHyperlinkOptions {
  text?: string;
  tooltip?: string;
}

export interface SetDataValidationOptions {
  type?: string;
  operator?: string;
  allowBlank?: boolean;
  showInputMessage?: boolean;
  showErrorMessage?: boolean;
  showDropDown?: boolean;
  errorStyle?: string;
  errorTitle?: string;
  error?: string;
  promptTitle?: string;
  prompt?: string;
  imeMode?: string;
  formula1?: string;
  formula2?: string;
}
