export interface ArchiveEntry {
  path: string;
  data: Uint8Array;
}

export type CellValue = string | number | boolean | null;
export type CellType = "missing" | "blank" | "string" | "number" | "boolean" | "formula";
export type SheetVisibility = "visible" | "hidden" | "veryHidden";
export type OpenXmlStringEnum = string & {};
export type CellStyleHorizontalAlignment =
  | "general"
  | "left"
  | "center"
  | "right"
  | "fill"
  | "justify"
  | "centerContinuous"
  | "distributed"
  | OpenXmlStringEnum;
export type CellStyleVerticalAlignment =
  | "top"
  | "center"
  | "bottom"
  | "justify"
  | "distributed"
  | OpenXmlStringEnum;
export type CellFontUnderline =
  | "single"
  | "double"
  | "singleAccounting"
  | "doubleAccounting"
  | OpenXmlStringEnum;
export type CellFontScheme = "major" | "minor" | OpenXmlStringEnum;
export type CellFontVerticalAlign = "baseline" | "superscript" | "subscript" | OpenXmlStringEnum;
export type CellFillPatternType =
  | "none"
  | "solid"
  | "mediumGray"
  | "darkGray"
  | "lightGray"
  | "darkHorizontal"
  | "darkVertical"
  | "darkDown"
  | "darkUp"
  | "darkGrid"
  | "darkTrellis"
  | "lightHorizontal"
  | "lightVertical"
  | "lightDown"
  | "lightUp"
  | "lightGrid"
  | "lightTrellis"
  | "gray125"
  | "gray0625"
  | OpenXmlStringEnum;
export type CellBorderStyle =
  | "thin"
  | "medium"
  | "dashed"
  | "dotted"
  | "thick"
  | "double"
  | "hair"
  | "mediumDashed"
  | "dashDot"
  | "mediumDashDot"
  | "dashDotDot"
  | "mediumDashDotDot"
  | "slantDashDot"
  | OpenXmlStringEnum;

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
  horizontal?: CellStyleHorizontalAlignment;
  vertical?: CellStyleVerticalAlignment;
  textRotation?: number;
  wrapText?: boolean;
  shrinkToFit?: boolean;
  indent?: number;
  relativeIndent?: number;
  justifyLastLine?: boolean;
  readingOrder?: number;
}

export interface CellStyleAlignmentPatch {
  horizontal?: CellStyleHorizontalAlignment | null;
  vertical?: CellStyleVerticalAlignment | null;
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

export interface CellNumberFormatDefinition {
  builtin: boolean;
  code: string | null;
  numFmtId: number;
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
  underline: CellFontUnderline | null;
  strike: boolean | null;
  outline: boolean | null;
  shadow: boolean | null;
  condense: boolean | null;
  extend: boolean | null;
  size: number | null;
  name: string | null;
  family: number | null;
  charset: number | null;
  scheme: CellFontScheme | null;
  vertAlign: CellFontVerticalAlign | null;
  color: CellFontColor | null;
}

export interface CellFontPatch {
  bold?: boolean | null;
  italic?: boolean | null;
  underline?: CellFontUnderline | null;
  strike?: boolean | null;
  outline?: boolean | null;
  shadow?: boolean | null;
  condense?: boolean | null;
  extend?: boolean | null;
  size?: number | null;
  name?: string | null;
  family?: number | null;
  charset?: number | null;
  scheme?: CellFontScheme | null;
  vertAlign?: CellFontVerticalAlign | null;
  color?: CellFontColorPatch | null;
}

export interface CellFillColor {
  rgb?: string;
  theme?: number;
  indexed?: number;
  auto?: boolean;
  tint?: number;
}

export interface CellFillColorPatch {
  rgb?: string | null;
  theme?: number | null;
  indexed?: number | null;
  auto?: boolean | null;
  tint?: number | null;
}

export interface CellFillDefinition {
  patternType: CellFillPatternType | null;
  fgColor: CellFillColor | null;
  bgColor: CellFillColor | null;
}

export interface CellFillPatch {
  patternType?: CellFillPatternType | null;
  fgColor?: CellFillColorPatch | null;
  bgColor?: CellFillColorPatch | null;
}

export interface CellBorderColor {
  rgb?: string;
  theme?: number;
  indexed?: number;
  auto?: boolean;
  tint?: number;
}

export interface CellBorderColorPatch {
  rgb?: string | null;
  theme?: number | null;
  indexed?: number | null;
  auto?: boolean | null;
  tint?: number | null;
}

export interface CellBorderSideDefinition {
  style: CellBorderStyle | null;
  color: CellBorderColor | null;
}

export interface CellBorderSidePatch {
  style?: CellBorderStyle | null;
  color?: CellBorderColorPatch | null;
}

export interface CellBorderDefinition {
  left: CellBorderSideDefinition | null;
  right: CellBorderSideDefinition | null;
  top: CellBorderSideDefinition | null;
  bottom: CellBorderSideDefinition | null;
  diagonal: CellBorderSideDefinition | null;
  vertical: CellBorderSideDefinition | null;
  horizontal: CellBorderSideDefinition | null;
  diagonalUp: boolean | null;
  diagonalDown: boolean | null;
  outline: boolean | null;
}

export interface CellBorderPatch {
  left?: CellBorderSidePatch | null;
  right?: CellBorderSidePatch | null;
  top?: CellBorderSidePatch | null;
  bottom?: CellBorderSidePatch | null;
  diagonal?: CellBorderSidePatch | null;
  vertical?: CellBorderSidePatch | null;
  horizontal?: CellBorderSidePatch | null;
  diagonalUp?: boolean | null;
  diagonalDown?: boolean | null;
  outline?: boolean | null;
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
