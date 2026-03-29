import { XlsxError } from "./errors.js";
import { parseStringItemText } from "./shared-strings.js";
import type { CellSnapshot, CellType, CellValue } from "./types.js";
import type { Workbook } from "./workbook.js";
import { decodeXmlText } from "./utils/xml.js";

export interface LocatedCell {
  address: string;
  start: number;
  end: number;
  attributesSource: string;
  snapshot: CellSnapshot;
  rowNumber: number;
  columnNumber: number;
}

export interface LocatedRow {
  start: number;
  end: number;
  attributesSource: string;
  selfClosing: boolean;
  rowNumber: number;
  innerStart: number;
  innerEnd: number;
  cells: LocatedCell[];
  cellsByColumn: Array<LocatedCell | undefined>;
  maxColumnNumber: number;
}

interface UsedRangeBounds {
  minRow: number;
  maxRow: number;
  minColumn: number;
  maxColumn: number;
}

export interface SheetIndex {
  xml: string;
  cells: Map<string, LocatedCell>;
  rows: Map<number, LocatedRow>;
  rowNumbers: number[];
  usedBounds: UsedRangeBounds | null;
  sheetDataInnerStart: number;
  sheetDataInnerEnd: number;
}

export function parseCellSnapshot(cell: LocatedCell | undefined): CellSnapshot {
  if (!cell) {
    return {
      exists: false,
      formula: null,
      rawType: null,
      styleId: null,
      type: "missing",
      value: null,
    };
  }

  return cell.snapshot;
}

export function buildSheetIndex(workbook: Workbook, sheetXml: string): SheetIndex {
  const sheetDataStart = sheetXml.indexOf("<sheetData");
  if (sheetDataStart === -1) {
    throw new XlsxError("Worksheet is missing <sheetData>");
  }

  const sheetDataOpenTagEnd = sheetXml.indexOf(">", sheetDataStart);
  const sheetDataCloseTagStart = sheetXml.indexOf("</sheetData>", sheetDataOpenTagEnd + 1);
  if (sheetDataOpenTagEnd === -1 || sheetDataCloseTagStart === -1) {
    throw new XlsxError("Worksheet is missing <sheetData>");
  }

  const sheetDataInnerStart = sheetDataOpenTagEnd + 1;
  const sheetDataInnerEnd = sheetDataCloseTagStart;
  const rows = new Map<number, LocatedRow>();
  const cells = new Map<string, LocatedCell>();
  const rowNumbers: number[] = [];
  let minRow = Number.POSITIVE_INFINITY;
  let maxRow = 0;
  let minColumn = Number.POSITIVE_INFINITY;
  let maxColumn = 0;
  let hasCells = false;
  let previousRowNumber = 0;
  let rowsAreSorted = true;
  let cursor = sheetDataInnerStart;

  while (cursor < sheetDataInnerEnd) {
    const rowStart = sheetXml.indexOf("<row", cursor);
    if (rowStart === -1 || rowStart >= sheetDataInnerEnd) {
      break;
    }

    const rowOpenTagEnd = sheetXml.indexOf(">", rowStart + 4);
    if (rowOpenTagEnd === -1 || rowOpenTagEnd >= sheetDataInnerEnd) {
      break;
    }

    const rowTagSource = sheetXml.slice(rowStart + 4, rowOpenTagEnd);
    const selfClosing = isSelfClosingTagSource(rowTagSource);
    const attributesSource = cleanTagAttributesSource(rowTagSource);
    const rowNumberText = readXmlAttrFast(attributesSource, "r");
    const rowEnd = selfClosing
      ? rowOpenTagEnd + 1
      : sheetXml.indexOf("</row>", rowOpenTagEnd + 1);

    if (!rowNumberText || rowEnd === -1) {
      cursor = rowOpenTagEnd + 1;
      continue;
    }

    const rowNumber = Number(rowNumberText);
    const innerStart = selfClosing ? rowEnd : rowOpenTagEnd + 1;
    const innerEnd = selfClosing ? rowEnd : rowEnd;
    const row: LocatedRow = {
      start: rowStart,
      end: selfClosing ? rowEnd : rowEnd + ROW_CLOSE_TAG.length,
      attributesSource,
      selfClosing,
      rowNumber,
      innerStart,
      innerEnd,
      cells: [],
      cellsByColumn: [],
      maxColumnNumber: 0,
    };

    if (!selfClosing) {
      let cellCursor = innerStart;
      let previousColumnNumber = 0;
      let cellsAreSorted = true;

      while (cellCursor < innerEnd) {
        const cellStart = sheetXml.indexOf("<c", cellCursor);
        if (cellStart === -1 || cellStart >= innerEnd) {
          break;
        }

        const cellOpenTagEnd = sheetXml.indexOf(">", cellStart + 2);
        if (cellOpenTagEnd === -1 || cellOpenTagEnd > innerEnd) {
          break;
        }

        const cellTagSource = sheetXml.slice(cellStart + 2, cellOpenTagEnd);
        const cellSelfClosing = isSelfClosingTagSource(cellTagSource);
        const cellAttributesSource = cleanTagAttributesSource(cellTagSource);
        const addressSource = readXmlAttrFast(cellAttributesSource, "r");
        const cellEnd = cellSelfClosing
          ? cellOpenTagEnd + 1
          : sheetXml.indexOf(CELL_CLOSE_TAG, cellOpenTagEnd + 1);

        if (!addressSource || cellEnd === -1) {
          cellCursor = cellOpenTagEnd + 1;
          continue;
        }

        const address = addressSource.toUpperCase();
        const columnNumber = columnLabelToNumberFromAddress(address);
        const innerXml = cellSelfClosing ? "" : sheetXml.slice(cellOpenTagEnd + 1, cellEnd);
        const rawType = extractCellTypeAttr(cellAttributesSource);
        const styleIdText = extractCellStyleAttr(cellAttributesSource);
        const styleId = styleIdText === undefined ? null : Number(styleIdText);
        const cell: LocatedCell = {
          address,
          start: cellStart,
          end: cellSelfClosing ? cellEnd : cellEnd + CELL_CLOSE_TAG.length,
          attributesSource: cellAttributesSource,
          snapshot: buildCellSnapshot(workbook, rawType, styleId, innerXml),
          rowNumber,
          columnNumber,
        };

        row.cells.push(cell);
        row.cellsByColumn[columnNumber] = cell;
        row.maxColumnNumber = Math.max(row.maxColumnNumber, columnNumber);
        cells.set(address, cell);
        hasCells = true;
        minRow = Math.min(minRow, rowNumber);
        maxRow = Math.max(maxRow, rowNumber);
        minColumn = Math.min(minColumn, columnNumber);
        maxColumn = Math.max(maxColumn, columnNumber);

        if (columnNumber < previousColumnNumber) {
          cellsAreSorted = false;
        }

        previousColumnNumber = columnNumber;
        cellCursor = cell.end;
      }

      if (!cellsAreSorted) {
        row.cells.sort((left, right) => left.columnNumber - right.columnNumber);
      }
    }

    rows.set(rowNumber, row);
    rowNumbers.push(rowNumber);
    if (rowNumber < previousRowNumber) {
      rowsAreSorted = false;
    }

    previousRowNumber = rowNumber;
    cursor = row.end;
  }

  if (!rowsAreSorted) {
    rowNumbers.sort((left, right) => left - right);
  }

  return {
    xml: sheetXml,
    cells,
    rows,
    rowNumbers,
    usedBounds: hasCells ? { minRow, maxRow, minColumn, maxColumn } : null,
    sheetDataInnerStart,
    sheetDataInnerEnd,
  };
}

export function parseCellAddressFast(address: string): { rowNumber: number; columnNumber: number } {
  let columnNumber = 0;
  let rowNumber = 0;
  let index = 0;

  while (index < address.length) {
    let characterCode = address.charCodeAt(index);
    if (characterCode === 36) {
      index += 1;
      continue;
    }

    if (characterCode >= 97 && characterCode <= 122) {
      characterCode -= 32;
    }

    if (characterCode < 65 || characterCode > 90) {
      break;
    }

    columnNumber = columnNumber * 26 + (characterCode - 64);
    index += 1;
  }

  while (index < address.length) {
    const characterCode = address.charCodeAt(index);
    if (characterCode === 36) {
      index += 1;
      continue;
    }

    if (characterCode < 48 || characterCode > 57) {
      throw new XlsxError(`Invalid cell address: ${address}`);
    }

    rowNumber = rowNumber * 10 + (characterCode - 48);
    index += 1;
  }

  if (columnNumber === 0 || rowNumber === 0) {
    throw new XlsxError(`Invalid cell address: ${address}`);
  }

  return { rowNumber, columnNumber };
}

function parseCellValue(
  workbook: Workbook,
  rawType: string | null,
  valueSource: string | undefined,
  inlineText: string | null,
  hasSelfClosingValue: boolean,
): CellValue {
  if (rawType === "inlineStr") {
    return inlineText ?? "";
  }

  if (rawType === "str") {
    if (valueSource !== undefined) {
      return decodeXmlText(valueSource);
    }

    return hasSelfClosingValue ? "" : null;
  }

  if (rawType === "s") {
    const indexText = valueSource;
    if (!indexText) {
      return null;
    }

    return workbook.getSharedString(Number(indexText));
  }

  if (rawType === "b") {
    return valueSource === "1";
  }

  if (valueSource === undefined) {
    return null;
  }

  const numericValue = Number(valueSource);
  return Number.isFinite(numericValue) ? numericValue : decodeXmlText(valueSource);
}

function buildCellSnapshot(
  workbook: Workbook,
  rawType: string | null,
  styleId: number | null,
  innerXml: string,
): CellSnapshot {
  const formulaSource = extractCellFormulaText(innerXml);
  const valueSource = extractCellValueText(innerXml);
  const inlineText = rawType === "inlineStr" ? parseStringItemText(innerXml) : null;
  const hasSelfClosingValue = valueSource === undefined && hasSelfClosingValueTag(innerXml);
  const formula = formulaSource === null ? null : decodeXmlText(formulaSource);
  const value = parseCellValue(workbook, rawType, valueSource, inlineText, hasSelfClosingValue);

  if (formula !== null) {
    return {
      exists: true,
      formula,
      rawType,
      styleId,
      type: "formula",
      value,
    };
  }

  const type: CellType =
    value === null
      ? "blank"
      : typeof value === "string"
        ? "string"
        : typeof value === "number"
          ? "number"
          : "boolean";

  return {
    exists: true,
    formula: null,
    rawType,
    styleId,
    type,
    value,
  };
}

function extractCellFormulaText(innerXml: string): string | null {
  return extractTagTextFast(innerXml, "f");
}

function extractCellValueText(innerXml: string): string | undefined {
  return extractTagTextFast(innerXml, "v") ?? undefined;
}

function hasSelfClosingValueTag(innerXml: string): boolean {
  return hasSelfClosingTagFast(innerXml, "v");
}

function extractCellTypeAttr(attributesSource: string): string | null {
  return readXmlAttrFast(attributesSource, "t") ?? null;
}

function extractCellStyleAttr(attributesSource: string): string | undefined {
  return readXmlAttrFast(attributesSource, "s");
}

function cleanTagAttributesSource(source: string): string {
  let end = source.length;

  while (end > 0 && isXmlWhitespaceCode(source.charCodeAt(end - 1))) {
    end -= 1;
  }

  if (source.charCodeAt(end - 1) === 47) {
    end -= 1;
  }

  while (end > 0 && isXmlWhitespaceCode(source.charCodeAt(end - 1))) {
    end -= 1;
  }

  let start = 0;
  while (start < end && isXmlWhitespaceCode(source.charCodeAt(start))) {
    start += 1;
  }

  return source.slice(start, end);
}

function isSelfClosingTagSource(source: string): boolean {
  let index = source.length - 1;

  while (index >= 0 && isXmlWhitespaceCode(source.charCodeAt(index))) {
    index -= 1;
  }

  return index >= 0 && source.charCodeAt(index) === 47;
}

function readXmlAttrFast(source: string, attributeName: string): string | undefined {
  const pattern = `${attributeName}="`;
  let searchStart = 0;

  while (searchStart < source.length) {
    const attributeStart = source.indexOf(pattern, searchStart);
    if (attributeStart === -1) {
      return undefined;
    }

    const previousCode = attributeStart === 0 ? 32 : source.charCodeAt(attributeStart - 1);
    if (isXmlAttributeBoundaryCode(previousCode)) {
      const valueStart = attributeStart + pattern.length;
      const valueEnd = source.indexOf("\"", valueStart);
      return valueEnd === -1 ? undefined : source.slice(valueStart, valueEnd);
    }

    searchStart = attributeStart + pattern.length;
  }

  return undefined;
}

function extractTagTextFast(xml: string, tagName: string): string | null {
  const tagStart = findTagStartFast(xml, tagName, 0);
  if (tagStart === -1) {
    return null;
  }

  const tagOpenEnd = xml.indexOf(">", tagStart + tagName.length + 1);
  if (tagOpenEnd === -1 || isSelfClosingTagSource(xml.slice(tagStart + tagName.length + 1, tagOpenEnd))) {
    return null;
  }

  const closeStart = xml.indexOf(`</${tagName}>`, tagOpenEnd + 1);
  return closeStart === -1 ? null : xml.slice(tagOpenEnd + 1, closeStart);
}

function hasSelfClosingTagFast(xml: string, tagName: string): boolean {
  const tagStart = findTagStartFast(xml, tagName, 0);
  if (tagStart === -1) {
    return false;
  }

  const tagOpenEnd = xml.indexOf(">", tagStart + tagName.length + 1);
  return tagOpenEnd !== -1 && isSelfClosingTagSource(xml.slice(tagStart + tagName.length + 1, tagOpenEnd));
}

function findTagStartFast(xml: string, tagName: string, fromIndex: number): number {
  const pattern = `<${tagName}`;
  let searchStart = fromIndex;

  while (searchStart < xml.length) {
    const tagStart = xml.indexOf(pattern, searchStart);
    if (tagStart === -1) {
      return -1;
    }

    const nextCode = xml.charCodeAt(tagStart + pattern.length);
    if (Number.isNaN(nextCode) || isXmlTagBoundaryCode(nextCode)) {
      return tagStart;
    }

    searchStart = tagStart + pattern.length;
  }

  return -1;
}

function columnLabelToNumberFromAddress(address: string): number {
  let value = 0;
  let index = 0;

  while (index < address.length) {
    let characterCode = address.charCodeAt(index);
    if (characterCode === 36) {
      index += 1;
      continue;
    }

    if (characterCode >= 97 && characterCode <= 122) {
      characterCode -= 32;
    }

    if (characterCode < 65 || characterCode > 90) {
      break;
    }

    value = value * 26 + (characterCode - 64);
    index += 1;
  }

  if (value === 0) {
    throw new XlsxError(`Invalid cell address: ${address}`);
  }

  return value;
}

function isXmlWhitespaceCode(code: number): boolean {
  return code === 9 || code === 10 || code === 13 || code === 32;
}

function isXmlAttributeBoundaryCode(code: number): boolean {
  return code === 47 || isXmlWhitespaceCode(code);
}

function isXmlTagBoundaryCode(code: number): boolean {
  return code === 47 || code === 62 || isXmlWhitespaceCode(code);
}

const ROW_CLOSE_TAG = "</row>";
const CELL_CLOSE_TAG = "</c>";
