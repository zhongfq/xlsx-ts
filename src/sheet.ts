import type { CellValue, SetFormulaOptions } from "./types.js";
import { XlsxError } from "./errors.js";
import type { Workbook } from "./workbook.js";
import {
  decodeXmlText,
  escapeRegex,
  escapeXmlText,
  extractAllTagTexts,
  extractTagText,
  getXmlAttr,
  parseAttributes,
  serializeAttributes,
} from "./utils/xml.js";

interface LocatedCell {
  start: number;
  end: number;
  attributesSource: string;
  innerXml: string;
  rowNumber: number;
  columnNumber: number;
}

interface LocatedRow {
  start: number;
  end: number;
  attributesSource: string;
  innerXml: string;
  selfClosing: boolean;
  rowNumber: number;
  innerStart: number;
  innerEnd: number;
  cells: LocatedCell[];
}

interface SheetIndex {
  xml: string;
  cells: Map<string, LocatedCell>;
  rows: Map<number, LocatedRow>;
  rowNumbers: number[];
  sheetDataInnerStart: number;
  sheetDataInnerEnd: number;
}

export class Sheet {
  readonly name: string;
  readonly path: string;
  readonly relationshipId: string;

  private readonly workbook: Workbook;
  private sheetIndex?: SheetIndex;

  constructor(
    workbook: Workbook,
    options: {
      name: string;
      path: string;
      relationshipId: string;
    },
  ) {
    this.workbook = workbook;
    this.name = options.name;
    this.path = options.path;
    this.relationshipId = options.relationshipId;
  }

  getCell(address: string): CellValue {
    const cell = this.getSheetIndex().cells.get(normalizeCellAddress(address));

    if (!cell) {
      return null;
    }

    const type = getXmlAttr(cell.attributesSource, "t");

    if (type === "inlineStr") {
      return extractAllTagTexts(cell.innerXml, "t").map(decodeXmlText).join("");
    }

    if (type === "str") {
      const rawString = extractTagText(cell.innerXml, "v");
      return rawString === undefined ? null : decodeXmlText(rawString);
    }

    if (type === "s") {
      const indexText = extractTagText(cell.innerXml, "v");
      if (!indexText) {
        return null;
      }

      const value = this.workbook.readSharedStrings()[Number(indexText)];
      return value ?? null;
    }

    if (type === "b") {
      return extractTagText(cell.innerXml, "v") === "1";
    }

    const rawValue = extractTagText(cell.innerXml, "v");
    if (rawValue === undefined) {
      return null;
    }

    const numericValue = Number(rawValue);
    return Number.isFinite(numericValue) ? numericValue : decodeXmlText(rawValue);
  }

  getFormula(address: string): string | null {
    const cell = this.getSheetIndex().cells.get(normalizeCellAddress(address));
    if (!cell) {
      return null;
    }

    const formula = extractTagText(cell.innerXml, "f");
    return formula === undefined ? null : decodeXmlText(formula);
  }

  setCell(address: string, value: CellValue): void {
    const normalizedAddress = normalizeCellAddress(address);
    const existingCell = this.getSheetIndex().cells.get(normalizedAddress);
    this.writeCellXml(
      normalizedAddress,
      buildValueCellXml(normalizedAddress, value, existingCell?.attributesSource),
    );
  }

  setFormula(address: string, formula: string, options: SetFormulaOptions = {}): void {
    const normalizedAddress = normalizeCellAddress(address);
    const existingCell = this.getSheetIndex().cells.get(normalizedAddress);
    this.writeCellXml(
      normalizedAddress,
      buildFormulaCellXml(
        normalizedAddress,
        formula,
        options.cachedValue ?? null,
        existingCell?.attributesSource,
      ),
    );
  }

  private getSheetIndex(): SheetIndex {
    if (this.sheetIndex) {
      return this.sheetIndex;
    }

    this.sheetIndex = buildSheetIndex(this.workbook.readEntryText(this.path));
    return this.sheetIndex;
  }

  private writeCellXml(address: string, cellXml: string): void {
    const index = this.getSheetIndex();
    const existingCell = index.cells.get(address);
    const nextSheetXml = existingCell
      ? index.xml.slice(0, existingCell.start) + cellXml + index.xml.slice(existingCell.end)
      : insertCell(index, address, cellXml);

    this.workbook.writeEntryText(this.path, nextSheetXml);
    this.sheetIndex = buildSheetIndex(nextSheetXml);
  }
}

function buildValueCellXml(address: string, value: CellValue, existingAttributesSource?: string): string {
  const attributes = parseAttributes(existingAttributesSource ?? "");
  const preserved = attributes.filter(([name]) => name !== "r" && name !== "t");
  const nextAttributes: Array<[string, string]> = [["r", address]];

  if (typeof value === "string") {
    nextAttributes.push(["t", "inlineStr"]);
  } else if (typeof value === "boolean") {
    nextAttributes.push(["t", "b"]);
  }

  nextAttributes.push(...preserved);

  const serializedAttributes = serializeAttributes(nextAttributes);

  if (value === null) {
    return `<c ${serializedAttributes}/>`;
  }

  if (typeof value === "string") {
    const needsSpace = value.trim() !== value;
    const space = needsSpace ? ' xml:space="preserve"' : "";
    return `<c ${serializedAttributes}><is><t${space}>${escapeXmlText(value)}</t></is></c>`;
  }

  if (typeof value === "boolean") {
    return `<c ${serializedAttributes}><v>${value ? "1" : "0"}</v></c>`;
  }

  return `<c ${serializedAttributes}><v>${String(value)}</v></c>`;
}

function buildFormulaCellXml(
  address: string,
  formula: string,
  cachedValue: CellValue,
  existingAttributesSource?: string,
): string {
  const attributes = parseAttributes(existingAttributesSource ?? "");
  const preserved = attributes.filter(([name]) => name !== "r" && name !== "t");
  const nextAttributes: Array<[string, string]> = [["r", address]];

  if (typeof cachedValue === "string") {
    nextAttributes.push(["t", "str"]);
  } else if (typeof cachedValue === "boolean") {
    nextAttributes.push(["t", "b"]);
  }

  nextAttributes.push(...preserved);

  const serializedAttributes = serializeAttributes(nextAttributes);
  const valueXml = buildFormulaValueXml(cachedValue);

  return `<c ${serializedAttributes}><f>${escapeXmlText(formula)}</f>${valueXml}</c>`;
}

function buildFormulaValueXml(value: CellValue): string {
  if (value === null) {
    return "";
  }

  if (typeof value === "string") {
    return `<v>${escapeXmlText(value)}</v>`;
  }

  if (typeof value === "boolean") {
    return `<v>${value ? "1" : "0"}</v>`;
  }

  return `<v>${String(value)}</v>`;
}

function insertCell(sheetIndex: SheetIndex, address: string, cellXml: string): string {
  const { rowNumber, columnNumber } = splitCellAddress(address);
  const row = sheetIndex.rows.get(rowNumber);

  if (row) {
    if (row.selfClosing) {
      const nextRowXml = `<row ${row.attributesSource}>${cellXml}</row>`;
      return sheetIndex.xml.slice(0, row.start) + nextRowXml + sheetIndex.xml.slice(row.end);
    }

    const insertionIndex = findCellInsertionIndex(row, columnNumber);
    return (
      sheetIndex.xml.slice(0, insertionIndex) +
      cellXml +
      sheetIndex.xml.slice(insertionIndex)
    );
  }

  const rowXml = `<row r="${rowNumber}">${cellXml}</row>`;
  const insertionIndex = findRowInsertionIndex(sheetIndex, rowNumber);

  return (
    sheetIndex.xml.slice(0, insertionIndex) +
    rowXml +
    sheetIndex.xml.slice(insertionIndex)
  );
}

function findCellInsertionIndex(row: LocatedRow, columnNumber: number): number {
  for (const cell of row.cells) {
    if (cell.columnNumber > columnNumber) {
      return cell.start;
    }
  }

  return row.innerEnd;
}

function findRowInsertionIndex(sheetIndex: SheetIndex, rowNumber: number): number {
  for (const candidateRow of sheetIndex.rowNumbers) {
    if (candidateRow > rowNumber) {
      const row = sheetIndex.rows.get(candidateRow);
      if (!row) {
        break;
      }

      return row.start;
    }
  }

  return sheetIndex.sheetDataInnerEnd;
}

function buildSheetIndex(sheetXml: string): SheetIndex {
  const sheetDataMatch = sheetXml.match(/<sheetData\b[^>]*>([\s\S]*?)<\/sheetData>/);
  if (!sheetDataMatch || sheetDataMatch.index === undefined) {
    throw new XlsxError("Worksheet is missing <sheetData>");
  }

  const sheetDataOpenTagEnd = sheetDataMatch[0].indexOf(">") + 1;
  const sheetDataCloseTagStart = sheetDataMatch[0].lastIndexOf("</sheetData>");
  const sheetDataInnerStart = sheetDataMatch.index + sheetDataOpenTagEnd;
  const sheetDataInnerEnd = sheetDataMatch.index + sheetDataCloseTagStart;
  const rows = new Map<number, LocatedRow>();
  const cells = new Map<string, LocatedCell>();
  const rowRegex = /<row\b([^>]*\br="(\d+)"[^>]*)\s*(?:>([\s\S]*?)<\/row>|\/>)/g;

  for (const match of sheetXml.matchAll(rowRegex)) {
    if (match.index === undefined) {
      continue;
    }

    const rowNumber = Number(match[2]);
    const fullMatch = match[0];
    const selfClosing = !fullMatch.includes("</row>");
    const innerXml = match[3] ?? "";
    const rowStart = match.index;
    const rowEnd = rowStart + fullMatch.length;
    const innerStart = selfClosing ? rowEnd : rowStart + fullMatch.indexOf(">") + 1;
    const innerEnd = selfClosing ? rowEnd : rowStart + fullMatch.lastIndexOf("</row>");
    const row: LocatedRow = {
      start: rowStart,
      end: rowEnd,
      attributesSource: match[1].trim(),
      innerXml,
      selfClosing,
      rowNumber,
      innerStart,
      innerEnd,
      cells: [],
    };

    if (!selfClosing) {
      const cellRegex = /<c\b([^>]*\br="([A-Z]+)(\d+)"[^>]*)\s*(?:>([\s\S]*?)<\/c>|\/>)/gi;

      for (const cellMatch of innerXml.matchAll(cellRegex)) {
        if (cellMatch.index === undefined) {
          continue;
        }

        const fullCellMatch = cellMatch[0];
        const address = `${cellMatch[2].toUpperCase()}${cellMatch[3]}`;
        const cellStart = innerStart + cellMatch.index;
        const cell: LocatedCell = {
          start: cellStart,
          end: cellStart + fullCellMatch.length,
          attributesSource: cellMatch[1].trim(),
          innerXml: cellMatch[4] ?? "",
          rowNumber,
          columnNumber: columnLabelToNumber(cellMatch[2].toUpperCase()),
        };

        row.cells.push(cell);
        cells.set(address, cell);
      }
    }

    rows.set(rowNumber, row);
  }

  const rowNumbers = [...rows.keys()].sort((left, right) => left - right);

  return {
    xml: sheetXml,
    cells,
    rows,
    rowNumbers,
    sheetDataInnerStart,
    sheetDataInnerEnd,
  };
}

function splitCellAddress(address: string): { rowNumber: number; columnNumber: number } {
  const match = address.match(/^([A-Z]+)(\d+)$/i);
  if (!match) {
    throw new XlsxError(`Invalid cell address: ${address}`);
  }

  return {
    columnNumber: columnLabelToNumber(match[1].toUpperCase()),
    rowNumber: Number(match[2]),
  };
}

function columnLabelToNumber(label: string): number {
  let value = 0;

  for (const character of label.toUpperCase()) {
    value = value * 26 + (character.charCodeAt(0) - 64);
  }

  return value;
}

function assertCellAddress(address: string): void {
  if (!/^[A-Z]+[1-9]\d*$/i.test(address)) {
    throw new XlsxError(`Invalid cell address: ${address}`);
  }
}

function normalizeCellAddress(address: string): string {
  assertCellAddress(address);
  return address.toUpperCase();
}
