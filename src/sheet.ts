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
  address: string;
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

  getRow(rowNumber: number): CellValue[] {
    assertRowNumber(rowNumber);

    const row = this.getSheetIndex().rows.get(rowNumber);
    if (!row || row.cells.length === 0) {
      return [];
    }

    const values: CellValue[] = [];
    const maxColumn = Math.max(...row.cells.map((cell) => cell.columnNumber));

    for (let columnNumber = 1; columnNumber <= maxColumn; columnNumber += 1) {
      values.push(this.getCell(makeCellAddress(rowNumber, columnNumber)));
    }

    return values;
  }

  getColumn(column: number | string): CellValue[] {
    const columnNumber = normalizeColumnNumber(column);
    const cells = [...this.getSheetIndex().cells.values()]
      .filter((cell) => cell.columnNumber === columnNumber)
      .sort((left, right) => left.rowNumber - right.rowNumber);

    if (cells.length === 0) {
      return [];
    }

    const values: CellValue[] = [];
    const maxRow = cells[cells.length - 1].rowNumber;

    for (let rowNumber = 1; rowNumber <= maxRow; rowNumber += 1) {
      values.push(this.getCell(makeCellAddress(rowNumber, columnNumber)));
    }

    return values;
  }

  getRecords(headerRowNumber = 1): Array<Record<string, CellValue>> {
    const headers = this.getRow(headerRowNumber);
    let lastHeaderColumn = 0;

    for (let columnIndex = 0; columnIndex < headers.length; columnIndex += 1) {
      const value = headers[columnIndex];
      if (value !== null) {
        lastHeaderColumn = columnIndex + 1;
      }
    }

    if (lastHeaderColumn === 0) {
      return [];
    }

    const records: Array<Record<string, CellValue>> = [];
    const maxRow = this.getSheetIndex().rowNumbers.at(-1) ?? headerRowNumber;

    for (let rowNumber = headerRowNumber + 1; rowNumber <= maxRow; rowNumber += 1) {
      const row = this.getRow(rowNumber);
      const hasAnyValue = row.some((value) => value !== null);
      if (!hasAnyValue) {
        continue;
      }

      const record: Record<string, CellValue> = {};

      for (let columnIndex = 0; columnIndex < lastHeaderColumn; columnIndex += 1) {
        const header = headers[columnIndex];
        if (typeof header !== "string" || header.length === 0) {
          continue;
        }

        record[header] = row[columnIndex] ?? null;
      }

      records.push(record);
    }

    return records;
  }

  getRange(range: string): CellValue[][] {
    const { startRow, endRow, startColumn, endColumn } = parseRangeRef(range);
    const values: CellValue[][] = [];

    for (let rowNumber = startRow; rowNumber <= endRow; rowNumber += 1) {
      const rowValues: CellValue[] = [];

      for (let columnNumber = startColumn; columnNumber <= endColumn; columnNumber += 1) {
        rowValues.push(this.getCell(makeCellAddress(rowNumber, columnNumber)));
      }

      values.push(rowValues);
    }

    return values;
  }

  getUsedRange(): string | null {
    const cells = [...this.getSheetIndex().cells.values()];
    if (cells.length === 0) {
      return null;
    }

    let minRow = Number.POSITIVE_INFINITY;
    let maxRow = 0;
    let minColumn = Number.POSITIVE_INFINITY;
    let maxColumn = 0;

    for (const cell of cells) {
      minRow = Math.min(minRow, cell.rowNumber);
      maxRow = Math.max(maxRow, cell.rowNumber);
      minColumn = Math.min(minColumn, cell.columnNumber);
      maxColumn = Math.max(maxColumn, cell.columnNumber);
    }

    return formatRangeRef(minRow, minColumn, maxRow, maxColumn);
  }

  getMergedRanges(): string[] {
    return parseMergedRanges(this.getSheetIndex().xml);
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

  addMergedRange(range: string): void {
    const normalizedRange = normalizeRangeRef(range);
    const ranges = this.getMergedRanges();
    if (ranges.includes(normalizedRange)) {
      return;
    }

    this.writeSheetXml(updateMergedRanges(this.getSheetIndex().xml, [...ranges, normalizedRange]));
  }

  removeMergedRange(range: string): void {
    const normalizedRange = normalizeRangeRef(range);
    const ranges = this.getMergedRanges().filter((candidate) => candidate !== normalizedRange);
    this.writeSheetXml(updateMergedRanges(this.getSheetIndex().xml, ranges));
  }

  setRow(rowNumber: number, values: CellValue[], startColumn = 1): void {
    assertRowNumber(rowNumber);
    assertColumnNumber(startColumn);

    for (let columnOffset = 0; columnOffset < values.length; columnOffset += 1) {
      this.setCell(makeCellAddress(rowNumber, startColumn + columnOffset), values[columnOffset]);
    }
  }

  setColumn(column: number | string, values: CellValue[], startRow = 1): void {
    const columnNumber = normalizeColumnNumber(column);
    assertRowNumber(startRow);

    for (let rowOffset = 0; rowOffset < values.length; rowOffset += 1) {
      this.setCell(makeCellAddress(startRow + rowOffset, columnNumber), values[rowOffset]);
    }
  }

  addRecord(record: Record<string, CellValue>, headerRowNumber = 1): void {
    const headerMap = this.getHeaderMap(headerRowNumber);
    const keys = Object.keys(record);
    if (keys.length === 0) {
      return;
    }

    for (const key of keys) {
      if (!headerMap.has(key)) {
        throw new XlsxError(`Header not found: ${key}`);
      }
    }

    const nextRowNumber = Math.max(headerRowNumber + 1, (this.getSheetIndex().rowNumbers.at(-1) ?? headerRowNumber) + 1);

    for (const key of keys) {
      const columnNumber = headerMap.get(key);
      if (!columnNumber) {
        continue;
      }

      this.setCell(makeCellAddress(nextRowNumber, columnNumber), record[key] ?? null);
    }
  }

  addRecords(records: Array<Record<string, CellValue>>, headerRowNumber = 1): void {
    if (records.length === 0) {
      return;
    }

    const headerMap = this.getHeaderMap(headerRowNumber);
    let nextRowNumber = Math.max(headerRowNumber + 1, (this.getSheetIndex().rowNumbers.at(-1) ?? headerRowNumber) + 1);

    for (const record of records) {
      const keys = Object.keys(record);
      if (keys.length === 0) {
        nextRowNumber += 1;
        continue;
      }

      for (const key of keys) {
        if (!headerMap.has(key)) {
          throw new XlsxError(`Header not found: ${key}`);
        }
      }

      for (const key of keys) {
        const columnNumber = headerMap.get(key);
        if (!columnNumber) {
          continue;
        }

        this.setCell(makeCellAddress(nextRowNumber, columnNumber), record[key] ?? null);
      }

      nextRowNumber += 1;
    }
  }

  private getHeaderMap(headerRowNumber: number): Map<string, number> {
    assertRowNumber(headerRowNumber);

    const headers = this.getRow(headerRowNumber);
    const headerMap = new Map<string, number>();

    headers.forEach((value, index) => {
      if (typeof value === "string" && value.length > 0 && !headerMap.has(value)) {
        headerMap.set(value, index + 1);
      }
    });

    return headerMap;
  }

  setRange(startAddress: string, values: CellValue[][]): void {
    const normalizedStartAddress = normalizeCellAddress(startAddress);
    if (values.length === 0) {
      return;
    }

    const expectedWidth = values[0]?.length ?? 0;
    if (expectedWidth === 0) {
      throw new XlsxError("Range values must contain at least one column");
    }

    for (const row of values) {
      if (row.length !== expectedWidth) {
        throw new XlsxError("Range values must be rectangular");
      }
    }

    const { rowNumber: startRow, columnNumber: startColumn } = splitCellAddress(normalizedStartAddress);

    for (let rowOffset = 0; rowOffset < values.length; rowOffset += 1) {
      const row = values[rowOffset];

      for (let columnOffset = 0; columnOffset < row.length; columnOffset += 1) {
        this.setCell(makeCellAddress(startRow + rowOffset, startColumn + columnOffset), row[columnOffset]);
      }
    }
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

    this.writeSheetXml(nextSheetXml);
  }

  private writeSheetXml(nextSheetXml: string): void {
    const indexedSheet = buildSheetIndex(nextSheetXml);
    const normalizedSheetXml = updateDimensionRef(indexedSheet);

    this.workbook.writeEntryText(this.path, normalizedSheetXml);
    this.sheetIndex = buildSheetIndex(normalizedSheetXml);
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
  const rowRegex = /<row\b([^>]*?\br="(\d+)"[^>]*?)\s*(?:>([\s\S]*?)<\/row>|\/>)/g;

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
      const cellRegex = /<c\b([^>]*?\br="([A-Z]+)(\d+)"[^>]*?)\s*(?:>([\s\S]*?)<\/c>|\/>)/gi;

      for (const cellMatch of innerXml.matchAll(cellRegex)) {
        if (cellMatch.index === undefined) {
          continue;
        }

        const fullCellMatch = cellMatch[0];
        const address = `${cellMatch[2].toUpperCase()}${cellMatch[3]}`;
        const cellStart = innerStart + cellMatch.index;
        const cell: LocatedCell = {
          address,
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

      row.cells.sort((left, right) => left.columnNumber - right.columnNumber);
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

function normalizeColumnNumber(column: number | string): number {
  if (typeof column === "number") {
    assertColumnNumber(column);
    return column;
  }

  if (!/^[A-Z]+$/i.test(column)) {
    throw new XlsxError(`Invalid column label: ${column}`);
  }

  return columnLabelToNumber(column.toUpperCase());
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

function normalizeRangeRef(range: string): string {
  const { startRow, endRow, startColumn, endColumn } = parseRangeRef(range);
  return formatRangeRef(startRow, startColumn, endRow, endColumn);
}

function parseRangeRef(range: string): {
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
} {
  const normalizedRange = range.toUpperCase();
  const [startAddress, endAddress = normalizedRange] = normalizedRange.split(":");

  if (!startAddress || !endAddress) {
    throw new XlsxError(`Invalid range reference: ${range}`);
  }

  const start = splitCellAddress(startAddress);
  const end = splitCellAddress(endAddress);

  return {
    startRow: Math.min(start.rowNumber, end.rowNumber),
    endRow: Math.max(start.rowNumber, end.rowNumber),
    startColumn: Math.min(start.columnNumber, end.columnNumber),
    endColumn: Math.max(start.columnNumber, end.columnNumber),
  };
}

function makeCellAddress(rowNumber: number, columnNumber: number): string {
  return `${numberToColumnLabel(columnNumber)}${rowNumber}`;
}

function formatRangeRef(
  startRow: number,
  startColumn: number,
  endRow: number,
  endColumn: number,
): string {
  const startAddress = makeCellAddress(startRow, startColumn);
  const endAddress = makeCellAddress(endRow, endColumn);
  return startAddress === endAddress ? startAddress : `${startAddress}:${endAddress}`;
}

function numberToColumnLabel(columnNumber: number): string {
  assertColumnNumber(columnNumber);

  let remaining = columnNumber;
  let label = "";

  while (remaining > 0) {
    const offset = (remaining - 1) % 26;
    label = String.fromCharCode(65 + offset) + label;
    remaining = Math.floor((remaining - 1) / 26);
  }

  return label;
}

function parseMergedRanges(sheetXml: string): string[] {
  const mergeCellsMatch = sheetXml.match(/<mergeCells\b[^>]*>([\s\S]*?)<\/mergeCells>/);
  if (!mergeCellsMatch) {
    return [];
  }

  return Array.from(
    mergeCellsMatch[1].matchAll(/<mergeCell\b[^>]*\bref="([^"]+)"[^>]*\/>/g),
    (match) => normalizeRangeRef(match[1]),
  );
}

function updateMergedRanges(sheetXml: string, ranges: string[]): string {
  const normalizedRanges = [...new Set(ranges.map(normalizeRangeRef))].sort(compareRangeRefs);
  const existingMergeCellsMatch = sheetXml.match(/<mergeCells\b[^>]*>[\s\S]*?<\/mergeCells>/);

  if (normalizedRanges.length === 0) {
    if (!existingMergeCellsMatch || existingMergeCellsMatch.index === undefined) {
      return sheetXml;
    }

    return (
      sheetXml.slice(0, existingMergeCellsMatch.index) +
      sheetXml.slice(existingMergeCellsMatch.index + existingMergeCellsMatch[0].length)
    );
  }

  const mergeCellsXml =
    `<mergeCells count="${normalizedRanges.length}">` +
    normalizedRanges.map((range) => `<mergeCell ref="${range}"/>`).join("") +
    `</mergeCells>`;

  if (existingMergeCellsMatch && existingMergeCellsMatch.index !== undefined) {
    return (
      sheetXml.slice(0, existingMergeCellsMatch.index) +
      mergeCellsXml +
      sheetXml.slice(existingMergeCellsMatch.index + existingMergeCellsMatch[0].length)
    );
  }

  const sheetDataCloseTag = "</sheetData>";
  const insertionIndex = sheetXml.indexOf(sheetDataCloseTag);
  if (insertionIndex === -1) {
    throw new XlsxError("Worksheet is missing </sheetData>");
  }

  const anchorIndex = insertionIndex + sheetDataCloseTag.length;
  return sheetXml.slice(0, anchorIndex) + mergeCellsXml + sheetXml.slice(anchorIndex);
}

function compareRangeRefs(left: string, right: string): number {
  const leftRange = parseRangeRef(left);
  const rightRange = parseRangeRef(right);

  return (
    leftRange.startRow - rightRange.startRow ||
    leftRange.startColumn - rightRange.startColumn ||
    leftRange.endRow - rightRange.endRow ||
    leftRange.endColumn - rightRange.endColumn
  );
}

function updateDimensionRef(sheetIndex: SheetIndex): string {
  const usedRange = getUsedRangeFromCells(sheetIndex.cells.values());
  const dimensionMatch = sheetIndex.xml.match(/<dimension\b([^>]*?)\/>/);

  if (!usedRange) {
    if (!dimensionMatch || dimensionMatch.index === undefined) {
      return sheetIndex.xml;
    }

    return (
      sheetIndex.xml.slice(0, dimensionMatch.index) +
      sheetIndex.xml.slice(dimensionMatch.index + dimensionMatch[0].length)
    );
  }

  const dimensionXml = `<dimension ref="${usedRange}"/>`;

  if (dimensionMatch && dimensionMatch.index !== undefined) {
    return (
      sheetIndex.xml.slice(0, dimensionMatch.index) +
      dimensionXml +
      sheetIndex.xml.slice(dimensionMatch.index + dimensionMatch[0].length)
    );
  }

  const worksheetOpenTagMatch = sheetIndex.xml.match(/<worksheet\b[^>]*>/);
  if (!worksheetOpenTagMatch || worksheetOpenTagMatch.index === undefined) {
    throw new XlsxError("Worksheet is missing opening tag");
  }

  return (
    sheetIndex.xml.slice(0, worksheetOpenTagMatch.index + worksheetOpenTagMatch[0].length) +
    dimensionXml +
    sheetIndex.xml.slice(worksheetOpenTagMatch.index + worksheetOpenTagMatch[0].length)
  );
}

function getUsedRangeFromCells(cells: Iterable<LocatedCell>): string | null {
  let minRow = Number.POSITIVE_INFINITY;
  let maxRow = 0;
  let minColumn = Number.POSITIVE_INFINITY;
  let maxColumn = 0;
  let hasCells = false;

  for (const cell of cells) {
    hasCells = true;
    minRow = Math.min(minRow, cell.rowNumber);
    maxRow = Math.max(maxRow, cell.rowNumber);
    minColumn = Math.min(minColumn, cell.columnNumber);
    maxColumn = Math.max(maxColumn, cell.columnNumber);
  }

  return hasCells ? formatRangeRef(minRow, minColumn, maxRow, maxColumn) : null;
}

function assertRowNumber(rowNumber: number): void {
  if (!Number.isInteger(rowNumber) || rowNumber < 1) {
    throw new XlsxError(`Invalid row number: ${rowNumber}`);
  }
}

function assertColumnNumber(columnNumber: number): void {
  if (!Number.isInteger(columnNumber) || columnNumber < 1) {
    throw new XlsxError(`Invalid column number: ${columnNumber}`);
  }
}
