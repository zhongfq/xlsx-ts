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

export interface CellStyleAlignment {
  horizontal?: string;
  vertical?: string;
  textRotation?: number;
  wrapText?: boolean;
  shrinkToFit?: boolean;
  indent?: number;
  relativeIndent?: number;
  justifyLastLine?: boolean;
  readingOrder?: number;
}

export interface CellStyleAlignmentPatch {
  horizontal?: string | null;
  vertical?: string | null;
  textRotation?: number | null;
  wrapText?: boolean | null;
  shrinkToFit?: boolean | null;
  indent?: number | null;
  relativeIndent?: number | null;
  justifyLastLine?: boolean | null;
  readingOrder?: number | null;
}

export interface CellStyleDefinition {
  numFmtId: number;
  fontId: number;
  fillId: number;
  borderId: number;
  xfId: number | null;
  quotePrefix: boolean | null;
  pivotButton: boolean | null;
  applyNumberFormat: boolean | null;
  applyFont: boolean | null;
  applyFill: boolean | null;
  applyBorder: boolean | null;
  applyAlignment: boolean | null;
  applyProtection: boolean | null;
  alignment: CellStyleAlignment | null;
}

export interface CellStylePatch {
  numFmtId?: number;
  fontId?: number;
  fillId?: number;
  borderId?: number;
  xfId?: number | null;
  quotePrefix?: boolean | null;
  pivotButton?: boolean | null;
  applyNumberFormat?: boolean | null;
  applyFont?: boolean | null;
  applyFill?: boolean | null;
  applyBorder?: boolean | null;
  applyAlignment?: boolean | null;
  applyProtection?: boolean | null;
  alignment?: CellStyleAlignmentPatch | null;
}

export interface CellFontColor {
  rgb?: string;
  theme?: number;
  indexed?: number;
  auto?: boolean;
  tint?: number;
}

export interface CellFontColorPatch {
  rgb?: string | null;
  theme?: number | null;
  indexed?: number | null;
  auto?: boolean | null;
  tint?: number | null;
}

export interface CellFontDefinition {
  bold: boolean | null;
  italic: boolean | null;
  underline: string | null;
  strike: boolean | null;
  outline: boolean | null;
  shadow: boolean | null;
  condense: boolean | null;
  extend: boolean | null;
  size: number | null;
  name: string | null;
  family: number | null;
  charset: number | null;
  scheme: string | null;
  vertAlign: string | null;
  color: CellFontColor | null;
}

export interface CellFontPatch {
  bold?: boolean | null;
  italic?: boolean | null;
  underline?: string | null;
  strike?: boolean | null;
  outline?: boolean | null;
  shadow?: boolean | null;
  condense?: boolean | null;
  extend?: boolean | null;
  size?: number | null;
  name?: string | null;
  family?: number | null;
  charset?: number | null;
  scheme?: string | null;
  vertAlign?: string | null;
  color?: CellFontColorPatch | null;
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
