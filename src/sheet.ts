import { Cell } from "./cell.js";
import type { CellType, CellValue, SetFormulaOptions } from "./types.js";
import { XlsxError } from "./errors.js";
import type { Workbook } from "./workbook.js";
import { basenamePosix, dirnamePosix, resolvePosix } from "./utils/path.js";
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

interface TableReference {
  relationshipId: string;
  path: string;
}

export class Sheet {
  name: string;
  readonly path: string;
  readonly relationshipId: string;

  private readonly cellHandles = new Map<string, Cell>();
  private revision = 0;
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

  cell(address: string): Cell {
    const normalizedAddress = normalizeCellAddress(address);
    let cell = this.cellHandles.get(normalizedAddress);

    if (!cell) {
      cell = new Cell(this, normalizedAddress);
      this.cellHandles.set(normalizedAddress, cell);
    }

    return cell;
  }

  getCell(address: string): CellValue {
    return this.cell(address).value;
  }

  rename(name: string): void {
    this.workbook.renameSheet(this.name, name);
  }

  getFormula(address: string): string | null {
    return this.cell(address).formula;
  }

  getHeaders(headerRowNumber = 1): string[] {
    assertRowNumber(headerRowNumber);
    return this.getRow(headerRowNumber).map((value) => (typeof value === "string" ? value : ""));
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

  getRecord(rowNumber: number, headerRowNumber = 1): Record<string, CellValue> | null {
    assertRowNumber(rowNumber);

    const row = this.getRow(rowNumber);
    if (row.length === 0 || row.every((value) => value === null)) {
      return null;
    }

    const headers = this.getRow(headerRowNumber);
    const record: Record<string, CellValue> = {};

    for (let columnIndex = 0; columnIndex < headers.length; columnIndex += 1) {
      const header = headers[columnIndex];
      if (typeof header !== "string" || header.length === 0) {
        continue;
      }

      record[header] = row[columnIndex] ?? null;
    }

    return record;
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

  getTables(): Array<{ name: string; displayName: string; range: string; path: string }> {
    const tables: Array<{ name: string; displayName: string; range: string; path: string }> = [];

    for (const table of this.getTableReferences()) {
      const tableXml = this.workbook.readEntryText(table.path);
      const tableTagMatch = tableXml.match(/<table\b([^>]*?)>/);
      if (!tableTagMatch) {
        continue;
      }

      const attributesSource = tableTagMatch[1];
      const name = getXmlAttr(attributesSource, "name");
      const displayName = getXmlAttr(attributesSource, "displayName");
      const range = getXmlAttr(attributesSource, "ref");

      if (!name || !displayName || !range) {
        continue;
      }

      tables.push({ name, displayName, range: normalizeRangeRef(range), path: table.path });
    }

    return tables;
  }

  insertRow(rowNumber: number, count = 1): void {
    assertRowNumber(rowNumber);
    assertInsertCount(count);

    const index = this.getSheetIndex();
    let nextSheetXml = index.xml;
    const nextMergedRanges = this.getMergedRanges().map((range) =>
      shiftRangeRefRows(range, rowNumber, count),
    );

    for (const sourceRowNumber of [...index.rowNumbers].sort((left, right) => right - left)) {
      const row = index.rows.get(sourceRowNumber);
      if (!row) {
        continue;
      }

      const nextRowXml = transformRowXml(
        index.xml,
        row,
        this.name,
        0,
        0,
        rowNumber,
        count,
      );
      nextSheetXml = nextSheetXml.slice(0, row.start) + nextRowXml + nextSheetXml.slice(row.end);
    }

    nextSheetXml = updateMergedRanges(nextSheetXml, nextMergedRanges);
    nextSheetXml = transformWorksheetStructureReferences(nextSheetXml, 0, 0, rowNumber, count, "shift");
    this.writeSheetXml(nextSheetXml);
    this.syncReferencedFormulasInOtherSheets((formula) =>
      shiftFormulaReferences(formula, this.name, 0, 0, rowNumber, count, false),
    );
    this.workbook.rewriteDefinedNamesForSheetStructure(this.name, 0, 0, rowNumber, count, "shift");
    this.updateTableReferences(0, 0, rowNumber, count, "shift");
  }

  insertColumn(column: number | string, count = 1): void {
    const columnNumber = normalizeColumnNumber(column);
    assertInsertCount(count);

    const index = this.getSheetIndex();
    let nextSheetXml = index.xml;
    const nextMergedRanges = this.getMergedRanges().map((range) =>
      shiftRangeRefColumns(range, columnNumber, count),
    );

    for (const rowNumber of [...index.rowNumbers].sort((left, right) => right - left)) {
      const row = index.rows.get(rowNumber);
      if (!row || row.selfClosing || row.cells.length === 0) {
        continue;
      }

      const nextRowXml = transformRowXml(
        index.xml,
        row,
        this.name,
        columnNumber,
        count,
        0,
        0,
      );
      nextSheetXml = nextSheetXml.slice(0, row.start) + nextRowXml + nextSheetXml.slice(row.end);
    }

    nextSheetXml = updateMergedRanges(nextSheetXml, nextMergedRanges);
    nextSheetXml = transformWorksheetStructureReferences(
      nextSheetXml,
      columnNumber,
      count,
      0,
      0,
      "shift",
    );
    this.writeSheetXml(nextSheetXml);
    this.syncReferencedFormulasInOtherSheets((formula) =>
      shiftFormulaReferences(formula, this.name, columnNumber, count, 0, 0, false),
    );
    this.workbook.rewriteDefinedNamesForSheetStructure(this.name, columnNumber, count, 0, 0, "shift");
    this.updateTableReferences(columnNumber, count, 0, 0, "shift");
  }

  deleteRow(rowNumber: number, count = 1): void {
    assertRowNumber(rowNumber);
    assertInsertCount(count);

    const index = this.getSheetIndex();
    const deleteEndRow = rowNumber + count - 1;
    let nextSheetXml = index.xml;
    const nextMergedRanges = this.getMergedRanges()
      .map((range) => deleteRangeRefRows(range, rowNumber, count))
      .filter((range): range is string => range !== null);

    for (const sourceRowNumber of [...index.rowNumbers].sort((left, right) => right - left)) {
      const row = index.rows.get(sourceRowNumber);
      if (!row) {
        continue;
      }

      if (sourceRowNumber >= rowNumber && sourceRowNumber <= deleteEndRow) {
        nextSheetXml = nextSheetXml.slice(0, row.start) + nextSheetXml.slice(row.end);
        continue;
      }

      const nextRowXml = deleteRowTransform(index.xml, row, this.name, rowNumber, count);
      nextSheetXml = nextSheetXml.slice(0, row.start) + nextRowXml + nextSheetXml.slice(row.end);
    }

    nextSheetXml = updateMergedRanges(nextSheetXml, nextMergedRanges);
    nextSheetXml = transformWorksheetStructureReferences(nextSheetXml, 0, 0, rowNumber, count, "delete");
    this.writeSheetXml(nextSheetXml);
    this.syncReferencedFormulasInOtherSheets((formula) =>
      deleteFormulaReferences(formula, this.name, 0, 0, rowNumber, count, false),
    );
    this.workbook.rewriteDefinedNamesForSheetStructure(this.name, 0, 0, rowNumber, count, "delete");
    this.updateTableReferences(0, 0, rowNumber, count, "delete");
  }

  deleteColumn(column: number | string, count = 1): void {
    const columnNumber = normalizeColumnNumber(column);
    assertInsertCount(count);

    const index = this.getSheetIndex();
    let nextSheetXml = index.xml;
    const nextMergedRanges = this.getMergedRanges()
      .map((range) => deleteRangeRefColumns(range, columnNumber, count))
      .filter((range): range is string => range !== null);

    for (const rowNumber of [...index.rowNumbers].sort((left, right) => right - left)) {
      const row = index.rows.get(rowNumber);
      if (!row) {
        continue;
      }

      const nextRowXml = deleteColumnTransform(index.xml, row, this.name, columnNumber, count);
      nextSheetXml = nextSheetXml.slice(0, row.start) + nextRowXml + nextSheetXml.slice(row.end);
    }

    nextSheetXml = updateMergedRanges(nextSheetXml, nextMergedRanges);
    nextSheetXml = transformWorksheetStructureReferences(
      nextSheetXml,
      columnNumber,
      count,
      0,
      0,
      "delete",
    );
    this.writeSheetXml(nextSheetXml);
    this.syncReferencedFormulasInOtherSheets((formula) =>
      deleteFormulaReferences(formula, this.name, columnNumber, count, 0, 0, false),
    );
    this.workbook.rewriteDefinedNamesForSheetStructure(this.name, columnNumber, count, 0, 0, "delete");
    this.updateTableReferences(columnNumber, count, 0, 0, "delete");
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

  getRevision(): number {
    return this.revision;
  }

  setHeaders(headers: string[], headerRowNumber = 1, startColumn = 1): void {
    assertRowNumber(headerRowNumber);
    assertColumnNumber(startColumn);
    this.setRow(headerRowNumber, headers, startColumn);
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

  appendRow(values: CellValue[], startColumn = 1): number {
    assertColumnNumber(startColumn);
    const rowNumber = (this.getSheetIndex().rowNumbers.at(-1) ?? 0) + 1;
    this.setRow(rowNumber, values, startColumn);
    return rowNumber;
  }

  appendRows(rows: CellValue[][], startColumn = 1): number[] {
    assertColumnNumber(startColumn);

    const rowNumbers: number[] = [];
    let nextRowNumber = (this.getSheetIndex().rowNumbers.at(-1) ?? 0) + 1;

    for (const row of rows) {
      this.setRow(nextRowNumber, row, startColumn);
      rowNumbers.push(nextRowNumber);
      nextRowNumber += 1;
    }

    return rowNumbers;
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
    if (Object.keys(record).length === 0) {
      return;
    }

    const nextRowNumber = Math.max(headerRowNumber + 1, (this.getSheetIndex().rowNumbers.at(-1) ?? headerRowNumber) + 1);
    this.writeRecordRow(nextRowNumber, record, headerMap, false);
  }

  addRecords(records: Array<Record<string, CellValue>>, headerRowNumber = 1): void {
    if (records.length === 0) {
      return;
    }

    const headerMap = this.getHeaderMap(headerRowNumber);
    let nextRowNumber = Math.max(headerRowNumber + 1, (this.getSheetIndex().rowNumbers.at(-1) ?? headerRowNumber) + 1);

    for (const record of records) {
      if (Object.keys(record).length === 0) {
        nextRowNumber += 1;
        continue;
      }

      this.writeRecordRow(nextRowNumber, record, headerMap, false);
      nextRowNumber += 1;
    }
  }

  setRecord(rowNumber: number, record: Record<string, CellValue>, headerRowNumber = 1): void {
    assertRowNumber(rowNumber);

    const headerMap = this.getHeaderMap(headerRowNumber);
    if (Object.keys(record).length === 0) {
      return;
    }

    this.writeRecordRow(rowNumber, record, headerMap, false);
  }

  setRecords(records: Array<Record<string, CellValue>>, headerRowNumber = 1): void {
    const headerMap = this.getHeaderMap(headerRowNumber);
    const existingRecordRows = this.getSheetIndex().rowNumbers.filter(
      (rowNumber) => rowNumber > headerRowNumber && this.getRecord(rowNumber, headerRowNumber) !== null,
    );
    const targetRows: number[] = [];

    for (let index = 0; index < records.length; index += 1) {
      const rowNumber = headerRowNumber + 1 + index;
      this.writeRecordRow(rowNumber, records[index], headerMap, true);
      targetRows.push(rowNumber);
    }

    const rowsToDelete = existingRecordRows.filter((rowNumber) => !targetRows.includes(rowNumber));
    this.deleteRecords(rowsToDelete, headerRowNumber);
  }

  deleteRecord(rowNumber: number, headerRowNumber = 1): void {
    assertRowNumber(rowNumber);
    assertRowNumber(headerRowNumber);

    if (rowNumber <= headerRowNumber) {
      throw new XlsxError(`Cannot delete header row: ${rowNumber}`);
    }

    const row = this.getSheetIndex().rows.get(rowNumber);
    if (!row) {
      return;
    }

    const nextSheetXml = this.getSheetIndex().xml.slice(0, row.start) + this.getSheetIndex().xml.slice(row.end);
    this.writeSheetXml(nextSheetXml);
  }

  deleteRecords(rowNumbers: number[], headerRowNumber = 1): void {
    assertRowNumber(headerRowNumber);

    const uniqueRows = [...new Set(rowNumbers)];
    uniqueRows.sort((left, right) => right - left);

    for (const rowNumber of uniqueRows) {
      this.deleteRecord(rowNumber, headerRowNumber);
    }
  }

  readCellSnapshot(address: string): {
    exists: boolean;
    formula: string | null;
    rawType: string | null;
    styleId: number | null;
    type: CellType;
    value: CellValue;
  } {
    const locatedCell = this.getSheetIndex().cells.get(normalizeCellAddress(address));
    return parseCellSnapshot(this.workbook, locatedCell);
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

  private writeRecordRow(
    rowNumber: number,
    record: Record<string, CellValue>,
    headerMap: Map<string, number>,
    replaceMissingKeys: boolean,
  ): void {
    const keys = Object.keys(record);

    for (const key of keys) {
      if (!headerMap.has(key)) {
        throw new XlsxError(`Header not found: ${key}`);
      }
    }

    if (replaceMissingKeys) {
      for (const [header, columnNumber] of headerMap) {
        const nextValue = Object.hasOwn(record, header) ? record[header] ?? null : null;
        this.setCell(makeCellAddress(rowNumber, columnNumber), nextValue);
      }
      return;
    }

    for (const key of keys) {
      const columnNumber = headerMap.get(key);
      if (!columnNumber) {
        continue;
      }

      this.setCell(makeCellAddress(rowNumber, columnNumber), record[key] ?? null);
    }
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

  private syncReferencedFormulasInOtherSheets(
    transformFormula: (formula: string) => string,
  ): void {
    for (const sheet of this.workbook.getSheets()) {
      if (sheet.path === this.path) {
        continue;
      }

      sheet.rewriteFormulaTexts(transformFormula);
    }
  }

  private rewriteFormulaTexts(transformFormula: (formula: string) => string): void {
    const sheetXml = this.getSheetIndex().xml;
    let changed = false;
    const nextSheetXml = sheetXml.replace(/<f\b([^>]*)>([\s\S]*?)<\/f>/g, (match, attributesSource, formulaSource) => {
      const formula = decodeXmlText(formulaSource);
      const nextFormula = transformFormula(formula);

      if (nextFormula === formula) {
        return match;
      }

      changed = true;
      return `<f${attributesSource}>${escapeXmlText(nextFormula)}</f>`;
    });

    if (changed) {
      this.writeSheetXml(nextSheetXml);
    }
  }

  private getTableReferences(): TableReference[] {
    const sheetRelationshipIds = Array.from(
      this.getSheetIndex().xml.matchAll(/<tablePart\b[^>]*\br:id="([^"]+)"[^>]*\/>/g),
      (match) => match[1],
    );
    if (sheetRelationshipIds.length === 0) {
      return [];
    }

    const relationshipsPath = `${dirnamePosix(this.path)}/_rels/${basenamePosix(this.path)}.rels`;
    if (!this.workbook.listEntries().includes(relationshipsPath)) {
      return [];
    }

    const relationshipsXml = this.workbook.readEntryText(relationshipsPath);
    const baseDir = dirnamePosix(this.path);
    const tables: TableReference[] = [];

    for (const match of relationshipsXml.matchAll(/<Relationship\b([^>]*?)\/>/g)) {
      const attributesSource = match[1];
      const relationshipId = getXmlAttr(attributesSource, "Id");
      const type = getXmlAttr(attributesSource, "Type");
      const target = getXmlAttr(attributesSource, "Target");

      if (
        !relationshipId ||
        !type ||
        !target ||
        !sheetRelationshipIds.includes(relationshipId) ||
        !/\/table$/.test(type)
      ) {
        continue;
      }

      tables.push({
        relationshipId,
        path: resolvePosix(baseDir, target.replace(/^\/+/, "")),
      });
    }

    return tables;
  }

  private updateTableReferences(
    targetColumnNumber: number,
    columnCount: number,
    targetRowNumber: number,
    rowCount: number,
    mode: "shift" | "delete",
  ): void {
    const transformRange = (range: string) =>
      mode === "shift"
        ? shiftRangeRef(range, targetColumnNumber, columnCount, targetRowNumber, rowCount)
        : deleteRangeRef(range, targetColumnNumber, columnCount, targetRowNumber, rowCount);
    const removedRelationshipIds: string[] = [];

    for (const table of this.getTableReferences()) {
      const tableXml = this.workbook.readEntryText(table.path);
      const nextTableXml = rewriteTableReferenceXml(tableXml, transformRange);

      if (nextTableXml === null) {
        removedRelationshipIds.push(table.relationshipId);
        continue;
      }

      if (nextTableXml !== tableXml) {
        this.workbook.writeEntryText(table.path, nextTableXml);
      }
    }

    if (removedRelationshipIds.length === 0) {
      return;
    }

    const nextSheetXml = removeTablePartsFromSheetXml(this.getSheetIndex().xml, removedRelationshipIds);
    if (nextSheetXml !== this.getSheetIndex().xml) {
      this.writeSheetXml(nextSheetXml);
    }
  }

  private writeSheetXml(nextSheetXml: string): void {
    const indexedSheet = buildSheetIndex(nextSheetXml);
    const normalizedSheetXml = updateDimensionRef(indexedSheet);

    this.workbook.writeEntryText(this.path, normalizedSheetXml);
    this.sheetIndex = buildSheetIndex(normalizedSheetXml);
    this.revision += 1;
  }
}

function parseCellSnapshot(
  workbook: Workbook,
  cell: LocatedCell | undefined,
): {
  exists: boolean;
  formula: string | null;
  rawType: string | null;
  styleId: number | null;
  type: CellType;
  value: CellValue;
} {
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

  const rawType = getXmlAttr(cell.attributesSource, "t") ?? null;
  const styleIdText = getXmlAttr(cell.attributesSource, "s");
  const styleId = styleIdText === undefined ? null : Number(styleIdText);
  const formulaText = extractTagText(cell.innerXml, "f");
  const formula = formulaText === undefined ? null : decodeXmlText(formulaText);

  if (formula !== null) {
    return {
      exists: true,
      formula,
      rawType,
      styleId: Number.isFinite(styleId) ? styleId : null,
      type: "formula",
      value: parseCellValue(workbook, cell, rawType),
    };
  }

  const value = parseCellValue(workbook, cell, rawType);
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
    styleId: Number.isFinite(styleId) ? styleId : null,
    type,
    value,
  };
}

function parseCellValue(workbook: Workbook, cell: LocatedCell, rawType: string | null): CellValue {
  if (rawType === "inlineStr") {
    return extractAllTagTexts(cell.innerXml, "t").map(decodeXmlText).join("");
  }

  if (rawType === "str") {
    const rawString = extractTagText(cell.innerXml, "v");
    return rawString === undefined ? null : decodeXmlText(rawString);
  }

  if (rawType === "s") {
    const indexText = extractTagText(cell.innerXml, "v");
    if (!indexText) {
      return null;
    }

    const value = workbook.readSharedStrings()[Number(indexText)];
    return value ?? null;
  }

  if (rawType === "b") {
    return extractTagText(cell.innerXml, "v") === "1";
  }

  const rawValue = extractTagText(cell.innerXml, "v");
  if (rawValue === undefined) {
    return null;
  }

  const numericValue = Number(rawValue);
  return Number.isFinite(numericValue) ? numericValue : decodeXmlText(rawValue);
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

function transformRowXml(
  sheetXml: string,
  row: LocatedRow,
  sheetName: string,
  targetColumnNumber: number,
  columnCount: number,
  targetRowNumber: number,
  rowCount: number,
): string {
  const rowAttributes = parseAttributes(row.attributesSource);
  const nextRowAttributes = rowAttributes.map(([name, value]) => {
    if (name === "r") {
      return [name, String(shiftRowNumber(Number(value), targetRowNumber, rowCount))] as [string, string];
    }

    if (name === "spans") {
      return [name, shiftRowSpans(value, targetColumnNumber, columnCount)] as [string, string];
    }

    return [name, value] as [string, string];
  });

  const rowOpenTag = `<row ${serializeAttributes(nextRowAttributes)}>`;
  let nextInnerXml = "";
  let cursor = row.innerStart;

  for (const cell of row.cells) {
    nextInnerXml += sheetXml.slice(cursor, cell.start);
    nextInnerXml += transformCellXml(
      sheetXml.slice(cell.start, cell.end),
      cell,
      sheetName,
      targetColumnNumber,
      columnCount,
      targetRowNumber,
      rowCount,
    );
    cursor = cell.end;
  }

  nextInnerXml += sheetXml.slice(cursor, row.innerEnd);
  return `${rowOpenTag}${nextInnerXml}</row>`;
}

function transformCellXml(
  cellXml: string,
  cell: LocatedCell,
  sheetName: string,
  targetColumnNumber: number,
  columnCount: number,
  targetRowNumber: number,
  rowCount: number,
): string {
  const attributes = parseAttributes(cell.attributesSource);
  const nextAttributes = attributes.map(([name, value]) => {
    if (name === "r") {
      return [
        name,
        shiftCellAddress(
          value,
          targetColumnNumber,
          columnCount,
          targetRowNumber,
          rowCount,
        ),
      ] as [string, string];
    }

    return [name, value] as [string, string];
  });

  const nextCellOpenTag = `<c ${serializeAttributes(nextAttributes)}`;
  if (!cellXml.includes("</c>")) {
    return `${nextCellOpenTag}/>`;
  }

  const innerStart = cellXml.indexOf(">") + 1;
  const innerEnd = cellXml.lastIndexOf("</c>");
  let nextInnerXml = cellXml.slice(innerStart, innerEnd);

  nextInnerXml = nextInnerXml.replace(/<f\b([^>]*)>([\s\S]*?)<\/f>/g, (_match, attributesSource, formulaSource) => {
    const formulaAttributes = parseAttributes(attributesSource);
    const nextFormulaAttributes = formulaAttributes.map(([name, value]) => {
      if (name === "ref") {
        return [
          name,
          shiftRangeRef(
            value,
            targetColumnNumber,
            columnCount,
            targetRowNumber,
            rowCount,
          ),
        ] as [string, string];
      }

      return [name, value] as [string, string];
    });
    const serializedAttributes = serializeAttributes(nextFormulaAttributes);
    const shiftedFormula = shiftFormulaReferences(
      decodeXmlText(formulaSource),
      sheetName,
      targetColumnNumber,
      columnCount,
      targetRowNumber,
      rowCount,
    );

    return `<f${serializedAttributes ? ` ${serializedAttributes}` : ""}>${escapeXmlText(shiftedFormula)}</f>`;
  });

  return `${nextCellOpenTag}>${nextInnerXml}</c>`;
}

function deleteRowTransform(
  sheetXml: string,
  row: LocatedRow,
  sheetName: string,
  targetRowNumber: number,
  count: number,
): string {
  const nextRowNumber = deleteShiftRowNumber(row.rowNumber, targetRowNumber, count);
  const rowAttributes = parseAttributes(row.attributesSource)
    .filter(([name]) => name !== "r")
    .map(([name, value]) => [name, value] as [string, string]);
  const nextRowAttributes: Array<[string, string]> = [["r", String(nextRowNumber)], ...rowAttributes];

  if (row.selfClosing || row.cells.length === 0) {
    return `<row ${serializeAttributes(nextRowAttributes)}/>`;
  }

  const nextCells = row.cells.map((cell) =>
    deleteRowCellTransform(sheetXml.slice(cell.start, cell.end), cell, sheetName, targetRowNumber, count),
  );
  return `<row ${serializeAttributes(nextRowAttributes)}>${nextCells.join("")}</row>`;
}

function deleteColumnTransform(
  sheetXml: string,
  row: LocatedRow,
  sheetName: string,
  targetColumnNumber: number,
  count: number,
): string {
  const keptCells = row.cells
    .filter((cell) => !isColumnDeleted(cell.columnNumber, targetColumnNumber, count))
    .map((cell) => ({
      columnNumber: deleteShiftColumnNumber(cell.columnNumber, targetColumnNumber, count),
      xml: deleteColumnCellTransform(sheetXml.slice(cell.start, cell.end), cell, sheetName, targetColumnNumber, count),
    }));

  const baseAttributes = parseAttributes(row.attributesSource)
    .filter(([name]) => name !== "spans")
    .map(([name, value]) => [name, value] as [string, string]);

  if (keptCells.length === 0) {
    return `<row ${serializeAttributes(baseAttributes)}/>`;
  }

  const nextAttributes = [...baseAttributes];
  const spansIndex = nextAttributes.findIndex(([name]) => name === "spans");
  const spansValue = `${keptCells[0].columnNumber}:${keptCells[keptCells.length - 1].columnNumber}`;

  if (spansIndex === -1) {
    nextAttributes.push(["spans", spansValue]);
  } else {
    nextAttributes[spansIndex] = ["spans", spansValue];
  }

  return `<row ${serializeAttributes(nextAttributes)}>${keptCells.map((cell) => cell.xml).join("")}</row>`;
}

function deleteRowCellTransform(
  cellXml: string,
  cell: LocatedCell,
  sheetName: string,
  targetRowNumber: number,
  count: number,
): string {
  const attributes = parseAttributes(cell.attributesSource).map(([name, value]) => {
    if (name === "r") {
      const { columnNumber, rowNumber } = splitCellAddress(value);
      return [name, makeCellAddress(deleteShiftRowNumber(rowNumber, targetRowNumber, count), columnNumber)] as [string, string];
    }

    return [name, value] as [string, string];
  });

  return deleteTransformCellInnerXml(cellXml, attributes, sheetName, 0, 0, targetRowNumber, count);
}

function deleteColumnCellTransform(
  cellXml: string,
  cell: LocatedCell,
  sheetName: string,
  targetColumnNumber: number,
  count: number,
): string {
  const attributes = parseAttributes(cell.attributesSource).map(([name, value]) => {
    if (name === "r") {
      const { rowNumber, columnNumber } = splitCellAddress(value);
      return [name, makeCellAddress(rowNumber, deleteShiftColumnNumber(columnNumber, targetColumnNumber, count))] as [string, string];
    }

    return [name, value] as [string, string];
  });

  return deleteTransformCellInnerXml(cellXml, attributes, sheetName, targetColumnNumber, count, 0, 0);
}

function deleteTransformCellInnerXml(
  cellXml: string,
  attributes: Array<[string, string]>,
  sheetName: string,
  targetColumnNumber: number,
  columnCount: number,
  targetRowNumber: number,
  rowCount: number,
): string {
  const cellOpenTag = `<c ${serializeAttributes(attributes)}`;
  if (!cellXml.includes("</c>")) {
    return `${cellOpenTag}/>`;
  }

  const innerStart = cellXml.indexOf(">") + 1;
  const innerEnd = cellXml.lastIndexOf("</c>");
  let nextInnerXml = cellXml.slice(innerStart, innerEnd);

  nextInnerXml = nextInnerXml.replace(/<f\b([^>]*)>([\s\S]*?)<\/f>/g, (_match, attributesSource, formulaSource) => {
    const formulaAttributes = parseAttributes(attributesSource);
    const nextFormulaAttributes = formulaAttributes.map(([name, value]) => {
      if (name === "ref") {
        const nextRange = deleteRangeRef(
          value,
          targetColumnNumber,
          columnCount,
          targetRowNumber,
          rowCount,
        );

        return nextRange === null ? [name, "#REF!"] as [string, string] : [name, nextRange] as [string, string];
      }

      return [name, value] as [string, string];
    });

    const nextFormula = deleteFormulaReferences(
      decodeXmlText(formulaSource),
      sheetName,
      targetColumnNumber,
      columnCount,
      targetRowNumber,
      rowCount,
    );
    const serializedAttributes = serializeAttributes(nextFormulaAttributes);
    return `<f${serializedAttributes ? ` ${serializedAttributes}` : ""}>${escapeXmlText(nextFormula)}</f>`;
  });

  return `${cellOpenTag}>${nextInnerXml}</c>`;
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
  const match = address.match(/^\$?([A-Z]+)\$?(\d+)$/i);
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

function shiftCellAddress(
  address: string,
  targetColumnNumber: number,
  columnCount: number,
  targetRowNumber: number,
  rowCount: number,
): string {
  const { rowNumber, columnNumber } = splitCellAddress(address);
  return makeCellAddress(
    shiftRowNumber(rowNumber, targetRowNumber, rowCount),
    shiftColumnNumber(columnNumber, targetColumnNumber, columnCount),
  );
}

function shiftRangeRef(
  range: string,
  targetColumnNumber: number,
  columnCount: number,
  targetRowNumber: number,
  rowCount: number,
): string {
  const { startRow, endRow, startColumn, endColumn } = parseRangeRef(range);

  return formatRangeRef(
    shiftRowNumber(startRow, targetRowNumber, rowCount),
    shiftColumnNumber(startColumn, targetColumnNumber, columnCount),
    shiftRowNumber(endRow, targetRowNumber, rowCount),
    shiftColumnNumber(endColumn, targetColumnNumber, columnCount),
  );
}

function shiftCellAddressColumns(address: string, targetColumnNumber: number, count: number): string {
  return shiftCellAddress(address, targetColumnNumber, count, 0, 0);
}

function shiftRangeRefColumns(range: string, targetColumnNumber: number, count: number): string {
  return shiftRangeRef(range, targetColumnNumber, count, 0, 0);
}

function shiftRangeRefRows(range: string, targetRowNumber: number, count: number): string {
  return shiftRangeRef(range, 0, 0, targetRowNumber, count);
}

function deleteRangeRefColumns(range: string, targetColumnNumber: number, count: number): string | null {
  return deleteRangeRef(range, targetColumnNumber, count, 0, 0);
}

function deleteRangeRefRows(range: string, targetRowNumber: number, count: number): string | null {
  return deleteRangeRef(range, 0, 0, targetRowNumber, count);
}

function shiftColumnNumber(columnNumber: number, targetColumnNumber: number, count: number): number {
  if (targetColumnNumber <= 0 || count <= 0) {
    return columnNumber;
  }

  return columnNumber >= targetColumnNumber ? columnNumber + count : columnNumber;
}

function shiftRowNumber(rowNumber: number, targetRowNumber: number, count: number): number {
  if (targetRowNumber <= 0 || count <= 0) {
    return rowNumber;
  }

  return rowNumber >= targetRowNumber ? rowNumber + count : rowNumber;
}

function deleteShiftColumnNumber(columnNumber: number, targetColumnNumber: number, count: number): number {
  if (targetColumnNumber <= 0 || count <= 0) {
    return columnNumber;
  }

  return columnNumber > targetColumnNumber + count - 1 ? columnNumber - count : columnNumber;
}

function deleteShiftRowNumber(rowNumber: number, targetRowNumber: number, count: number): number {
  if (targetRowNumber <= 0 || count <= 0) {
    return rowNumber;
  }

  return rowNumber > targetRowNumber + count - 1 ? rowNumber - count : rowNumber;
}

function isColumnDeleted(columnNumber: number, targetColumnNumber: number, count: number): boolean {
  return targetColumnNumber > 0 && columnNumber >= targetColumnNumber && columnNumber <= targetColumnNumber + count - 1;
}

function isRowDeleted(rowNumber: number, targetRowNumber: number, count: number): boolean {
  return targetRowNumber > 0 && rowNumber >= targetRowNumber && rowNumber <= targetRowNumber + count - 1;
}

function deleteRangeRef(
  range: string,
  targetColumnNumber: number,
  columnCount: number,
  targetRowNumber: number,
  rowCount: number,
): string | null {
  const { startRow, endRow, startColumn, endColumn } = parseRangeRef(range);
  const nextColumns = deleteRangeAxis(startColumn, endColumn, targetColumnNumber, columnCount);
  const nextRows = deleteRangeAxis(startRow, endRow, targetRowNumber, rowCount);

  if (!nextColumns || !nextRows) {
    return null;
  }

  return formatRangeRef(nextRows.start, nextColumns.start, nextRows.end, nextColumns.end);
}

function deleteRangeAxis(
  start: number,
  end: number,
  target: number,
  count: number,
): { start: number; end: number } | null {
  if (target <= 0 || count <= 0) {
    return { start, end };
  }

  const deleteEnd = target + count - 1;

  if (end < target) {
    return { start, end };
  }

  if (start > deleteEnd) {
    return { start: start - count, end: end - count };
  }

  const hasLeft = start < target;
  const hasRight = end > deleteEnd;

  if (!hasLeft && !hasRight) {
    return null;
  }

  const nextStart = hasLeft ? start : target;
  const nextEnd = hasRight ? end - count : deleteEnd >= start ? target - 1 : end;

  if (nextStart > nextEnd) {
    return null;
  }

  return { start: nextStart, end: nextEnd };
}

function transformWorksheetStructureReferences(
  sheetXml: string,
  targetColumnNumber: number,
  columnCount: number,
  targetRowNumber: number,
  rowCount: number,
  mode: "shift" | "delete",
): string {
  const transformRange = (range: string) =>
    mode === "shift"
      ? shiftRangeRef(range, targetColumnNumber, columnCount, targetRowNumber, rowCount)
      : deleteRangeRef(range, targetColumnNumber, columnCount, targetRowNumber, rowCount);

  let nextSheetXml = sheetXml;

  for (const tagName of WORKSHEET_REF_TAGS) {
    nextSheetXml = rewriteWorksheetReferenceTag(nextSheetXml, tagName, "ref", false, transformRange);
  }

  for (const tagName of WORKSHEET_SQREF_TAGS) {
    nextSheetXml = rewriteWorksheetReferenceTag(nextSheetXml, tagName, "sqref", true, transformRange);
  }

  for (const [tagName, attributeName] of WORKSHEET_CELL_REF_ATTRIBUTES) {
    nextSheetXml = rewriteWorksheetCellReferenceAttribute(
      nextSheetXml,
      tagName,
      attributeName,
      mode,
      targetColumnNumber,
      columnCount,
      targetRowNumber,
      rowCount,
    );
  }

  nextSheetXml = rewriteCountedContainer(nextSheetXml, "dataValidations", "dataValidation", "count");
  nextSheetXml = rewriteEmptyContainer(nextSheetXml, "hyperlinks", "hyperlink");
  return nextSheetXml;
}

function rewriteWorksheetReferenceTag(
  sheetXml: string,
  tagName: string,
  attributeName: string,
  multipleRanges: boolean,
  transformRange: (range: string) => string | null,
): string {
  const regex = new RegExp(
    `<${escapeRegex(tagName)}\\b([^>]*?)(\\/?>|>[\\s\\S]*?<\\/${escapeRegex(tagName)}>)`,
    "g",
  );

  return sheetXml.replace(regex, (match, attributesSource, bodySource) => {
    const attributes = parseAttributes(attributesSource);
    const attributeIndex = attributes.findIndex(([name]) => name === attributeName);

    if (attributeIndex === -1) {
      return match;
    }

    const currentValue = attributes[attributeIndex]?.[1] ?? "";
    const nextValue = transformWorksheetReferenceValue(currentValue, multipleRanges, transformRange);

    if (nextValue === currentValue) {
      return match;
    }

    if (nextValue === null) {
      return "";
    }

    const nextAttributes = [...attributes];
    nextAttributes[attributeIndex] = [attributeName, nextValue];
    const serializedAttributes = serializeAttributes(nextAttributes);
    const tagOpen = serializedAttributes.length > 0 ? `<${tagName} ${serializedAttributes}` : `<${tagName}`;

    if (bodySource === "/>") {
      return `${tagOpen}/>`;
    }

    const closingTag = `</${tagName}>`;
    const innerXml = bodySource.slice(1, -closingTag.length);
    return `${tagOpen}>${innerXml}${closingTag}`;
  });
}

function transformWorksheetReferenceValue(
  value: string,
  multipleRanges: boolean,
  transformRange: (range: string) => string | null,
): string | null {
  if (multipleRanges) {
    const nextRanges = value
      .trim()
      .split(/\s+/)
      .filter((range) => range.length > 0)
      .map((range) => transformRange(range))
      .filter((range): range is string => range !== null);

    return nextRanges.length > 0 ? nextRanges.join(" ") : null;
  }

  return transformRange(value);
}

function rewriteCountedContainer(
  sheetXml: string,
  containerTagName: string,
  childTagName: string,
  countAttributeName: string,
): string {
  const regex = new RegExp(
    `<${escapeRegex(containerTagName)}\\b([^>]*)>([\\s\\S]*?)<\\/${escapeRegex(containerTagName)}>`,
    "g",
  );

  return sheetXml.replace(regex, (_match, attributesSource, innerXml) => {
    const childMatches = innerXml.match(new RegExp(`<${escapeRegex(childTagName)}\\b`, "g")) ?? [];
    if (childMatches.length === 0) {
      return "";
    }

    const attributes = parseAttributes(attributesSource);
    const countIndex = attributes.findIndex(([name]) => name === countAttributeName);
    const nextAttributes = [...attributes];

    if (countIndex === -1) {
      nextAttributes.push([countAttributeName, String(childMatches.length)]);
    } else {
      nextAttributes[countIndex] = [countAttributeName, String(childMatches.length)];
    }

    const serializedAttributes = serializeAttributes(nextAttributes);
    return `<${containerTagName}${serializedAttributes ? ` ${serializedAttributes}` : ""}>${innerXml}</${containerTagName}>`;
  });
}

function rewriteEmptyContainer(
  sheetXml: string,
  containerTagName: string,
  childTagName: string,
): string {
  const regex = new RegExp(
    `<${escapeRegex(containerTagName)}\\b([^>]*)>([\\s\\S]*?)<\\/${escapeRegex(containerTagName)}>`,
    "g",
  );

  return sheetXml.replace(regex, (match, _attributesSource, innerXml) => {
    return new RegExp(`<${escapeRegex(childTagName)}\\b`).test(innerXml) ? match : "";
  });
}

function rewriteWorksheetCellReferenceAttribute(
  sheetXml: string,
  tagName: string,
  attributeName: string,
  mode: "shift" | "delete",
  targetColumnNumber: number,
  columnCount: number,
  targetRowNumber: number,
  rowCount: number,
): string {
  const regex = new RegExp(
    `<${escapeRegex(tagName)}\\b([^>]*?)(\\/?>|>[\\s\\S]*?<\\/${escapeRegex(tagName)}>)`,
    "g",
  );

  return sheetXml.replace(regex, (match, attributesSource, bodySource) => {
    const attributes = parseAttributes(attributesSource);
    const attributeIndex = attributes.findIndex(([name]) => name === attributeName);

    if (attributeIndex === -1) {
      return match;
    }

    const currentValue = attributes[attributeIndex]?.[1] ?? "";
    const nextValue =
      mode === "shift"
        ? shiftCellAddress(
            currentValue,
            targetColumnNumber,
            columnCount,
            targetRowNumber,
            rowCount,
          )
        : deleteCellReferenceAddress(
            currentValue,
            targetColumnNumber,
            columnCount,
            targetRowNumber,
            rowCount,
          );

    if (nextValue === currentValue) {
      return match;
    }

    const nextAttributes = [...attributes];
    if (nextValue === null) {
      nextAttributes.splice(attributeIndex, 1);
    } else {
      nextAttributes[attributeIndex] = [attributeName, nextValue];
    }

    const serializedAttributes = serializeAttributes(nextAttributes);
    const tagOpen = serializedAttributes.length > 0 ? `<${tagName} ${serializedAttributes}` : `<${tagName}`;

    if (bodySource === "/>") {
      return `${tagOpen}/>`;
    }

    const closingTag = `</${tagName}>`;
    const innerXml = bodySource.slice(1, -closingTag.length);
    return `${tagOpen}>${innerXml}${closingTag}`;
  });
}

function deleteCellReferenceAddress(
  address: string,
  targetColumnNumber: number,
  columnCount: number,
  targetRowNumber: number,
  rowCount: number,
): string | null {
  const { rowNumber, columnNumber } = splitCellAddress(address);

  if (isColumnDeleted(columnNumber, targetColumnNumber, columnCount) || isRowDeleted(rowNumber, targetRowNumber, rowCount)) {
    return null;
  }

  return makeCellAddress(
    deleteShiftRowNumber(rowNumber, targetRowNumber, rowCount),
    deleteShiftColumnNumber(columnNumber, targetColumnNumber, columnCount),
  );
}

function shiftRowSpans(spans: string, targetColumnNumber: number, count: number): string {
  const match = spans.match(/^(\d+):(\d+)$/);
  if (!match) {
    return spans;
  }

  const startColumn = Number(match[1]);
  const endColumn = Number(match[2]);
  return `${startColumn >= targetColumnNumber ? startColumn + count : startColumn}:${endColumn >= targetColumnNumber ? endColumn + count : endColumn}`;
}

export function shiftFormulaReferences(
  formula: string,
  currentSheetName: string,
  targetColumnNumber: number,
  columnCount: number,
  targetRowNumber: number,
  rowCount: number,
  includeUnqualifiedReferences = true,
): string {
  let nextFormula = "";
  let cursor = 0;
  let inString = false;

  while (cursor < formula.length) {
    const character = formula[cursor];

    if (character === "\"") {
      nextFormula += character;

      if (inString && formula[cursor + 1] === "\"") {
        nextFormula += "\"";
        cursor += 2;
        continue;
      }

      inString = !inString;
      cursor += 1;
      continue;
    }

    if (inString) {
      nextFormula += character;
      cursor += 1;
      continue;
    }

    const remaining = formula.slice(cursor);
    const rangeMatch = remaining.match(
      /^((?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?)([A-Z]+)(\$?)(\d+):((?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?)([A-Z]+)(\$?)(\d+)/,
    );
    const previous = formula[cursor - 1];

    if (rangeMatch) {
      const [
        fullMatch,
        startSheetRef,
        startColumnDollar,
        startColumnLabel,
        startRowDollar,
        startRowText,
        endSheetRef,
        endColumnDollar,
        endColumnLabel,
        endRowDollar,
        endRowText,
      ] = rangeMatch;

      if (
        !matchesFormulaReference(startSheetRef, currentSheetName, includeUnqualifiedReferences, previous) ||
        (endSheetRef !== undefined && !matchesSheetReference(endSheetRef, currentSheetName))
      ) {
        nextFormula += fullMatch;
        cursor += fullMatch.length;
        continue;
      }

      const nextStartColumn = shiftColumnNumber(
        columnLabelToNumber(startColumnLabel),
        targetColumnNumber,
        columnCount,
      );
      const nextEndColumn = shiftColumnNumber(
        columnLabelToNumber(endColumnLabel),
        targetColumnNumber,
        columnCount,
      );
      const nextStartRow = shiftRowNumber(Number(startRowText), targetRowNumber, rowCount);
      const nextEndRow = shiftRowNumber(Number(endRowText), targetRowNumber, rowCount);
      const leftRef = `${startSheetRef ?? ""}${startColumnDollar}${numberToColumnLabel(nextStartColumn)}${startRowDollar}${nextStartRow}`;
      const rightRef = `${endSheetRef ?? ""}${endColumnDollar}${numberToColumnLabel(nextEndColumn)}${endRowDollar}${nextEndRow}`;
      nextFormula += `${leftRef}:${rightRef}`;
      cursor += fullMatch.length;
      continue;
    }

    const match = remaining.match(/^((?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?)([A-Z]+)(\$?)(\d+)/);

    if (!match) {
      nextFormula += character;
      cursor += 1;
      continue;
    }

    const [fullMatch, sheetRef, columnDollar, columnLabel, rowDollar, rowNumber] = match;

    if (!matchesFormulaReference(sheetRef, currentSheetName, includeUnqualifiedReferences, previous)) {
      nextFormula += fullMatch;
      cursor += fullMatch.length;
      continue;
    }

    const columnNumber = columnLabelToNumber(columnLabel);
    const nextColumnNumber = shiftColumnNumber(columnNumber, targetColumnNumber, columnCount);
    const nextRowNumber = shiftRowNumber(Number(rowNumber), targetRowNumber, rowCount);
    nextFormula += `${sheetRef ?? ""}${columnDollar}${numberToColumnLabel(nextColumnNumber)}${rowDollar}${String(nextRowNumber)}`;
    cursor += fullMatch.length;
  }

  return nextFormula;
}

export function deleteFormulaReferences(
  formula: string,
  currentSheetName: string,
  targetColumnNumber: number,
  columnCount: number,
  targetRowNumber: number,
  rowCount: number,
  includeUnqualifiedReferences = true,
): string {
  let nextFormula = "";
  let cursor = 0;
  let inString = false;

  while (cursor < formula.length) {
    const character = formula[cursor];

    if (character === "\"") {
      nextFormula += character;

      if (inString && formula[cursor + 1] === "\"") {
        nextFormula += "\"";
        cursor += 2;
        continue;
      }

      inString = !inString;
      cursor += 1;
      continue;
    }

    if (inString) {
      nextFormula += character;
      cursor += 1;
      continue;
    }

    const remaining = formula.slice(cursor);
    const rangeMatch = remaining.match(
      /^((?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?)([A-Z]+)(\$?)(\d+):((?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?)([A-Z]+)(\$?)(\d+)/,
    );

    if (rangeMatch) {
      const [
        fullMatch,
        startSheetRef,
        startColumnDollar,
        startColumnLabel,
        startRowDollar,
        startRowText,
        endSheetRef,
        endColumnDollar,
        endColumnLabel,
        endRowDollar,
        endRowText,
      ] = rangeMatch;
      const previous = formula[cursor - 1];

      if (
        !matchesFormulaReference(
          startSheetRef,
          currentSheetName,
          includeUnqualifiedReferences,
          previous,
        ) ||
        (endSheetRef !== undefined && !matchesSheetReference(endSheetRef, currentSheetName))
      ) {
        nextFormula += fullMatch;
        cursor += fullMatch.length;
        continue;
      }

      const nextRange = deleteRangeAxis(
        columnLabelToNumber(startColumnLabel),
        columnLabelToNumber(endColumnLabel),
        targetColumnNumber,
        columnCount,
      );
      const nextRows = deleteRangeAxis(
        Number(startRowText),
        Number(endRowText),
        targetRowNumber,
        rowCount,
      );

      if (!nextRange || !nextRows) {
        nextFormula += "#REF!";
      } else {
        const leftRef = `${startSheetRef ?? ""}${startColumnDollar}${numberToColumnLabel(nextRange.start)}${startRowDollar}${nextRows.start}`;
        const rightPrefix = endSheetRef ?? "";
        const rightRef = `${rightPrefix}${endColumnDollar}${numberToColumnLabel(nextRange.end)}${endRowDollar}${nextRows.end}`;
        nextFormula += `${leftRef}:${rightRef}`;
      }

      cursor += fullMatch.length;
      continue;
    }

    const refMatch = remaining.match(/^((?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?)([A-Z]+)(\$?)(\d+)/);
    if (!refMatch) {
      nextFormula += character;
      cursor += 1;
      continue;
    }

    const [fullMatch, sheetRef, columnDollar, columnLabel, rowDollar, rowText] = refMatch;
    const previous = formula[cursor - 1];

    if (
      !matchesFormulaReference(
        sheetRef,
        currentSheetName,
        includeUnqualifiedReferences,
        previous,
      )
    ) {
      nextFormula += fullMatch;
      cursor += fullMatch.length;
      continue;
    }

    const columnNumber = columnLabelToNumber(columnLabel);
    const rowNumber = Number(rowText);

    if (isColumnDeleted(columnNumber, targetColumnNumber, columnCount) || isRowDeleted(rowNumber, targetRowNumber, rowCount)) {
      nextFormula += "#REF!";
    } else {
      nextFormula += `${sheetRef ?? ""}${columnDollar}${numberToColumnLabel(deleteShiftColumnNumber(columnNumber, targetColumnNumber, columnCount))}${rowDollar}${deleteShiftRowNumber(rowNumber, targetRowNumber, rowCount)}`;
    }

    cursor += fullMatch.length;
  }

  return nextFormula;
}

export function deleteSheetFormulaReferences(formula: string, deletedSheetName: string): string {
  let nextFormula = "";
  let cursor = 0;
  let inString = false;

  while (cursor < formula.length) {
    const character = formula[cursor];

    if (character === "\"") {
      nextFormula += character;

      if (inString && formula[cursor + 1] === "\"") {
        nextFormula += "\"";
        cursor += 2;
        continue;
      }

      inString = !inString;
      cursor += 1;
      continue;
    }

    if (inString) {
      nextFormula += character;
      cursor += 1;
      continue;
    }

    const remaining = formula.slice(cursor);
    const rangeMatch = remaining.match(
      /^((?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?)([A-Z]+)(\$?)(\d+):((?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?)([A-Z]+)(\$?)(\d+)/,
    );

    if (rangeMatch) {
      const [fullMatch, startSheetRef, , , , , endSheetRef] = rangeMatch;

      if (
        matchesSheetReference(startSheetRef, deletedSheetName) ||
        matchesSheetReference(endSheetRef, deletedSheetName)
      ) {
        nextFormula += "#REF!";
      } else {
        nextFormula += fullMatch;
      }

      cursor += fullMatch.length;
      continue;
    }

    const refMatch = remaining.match(/^((?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?)([A-Z]+)(\$?)(\d+)/);
    if (!refMatch) {
      nextFormula += character;
      cursor += 1;
      continue;
    }

    const [fullMatch, sheetRef] = refMatch;

    nextFormula += matchesSheetReference(sheetRef, deletedSheetName) ? "#REF!" : fullMatch;
    cursor += fullMatch.length;
  }

  return nextFormula;
}

export function renameSheetFormulaReferences(
  formula: string,
  previousSheetName: string,
  nextSheetName: string,
): string {
  let nextFormula = "";
  let cursor = 0;
  let inString = false;

  while (cursor < formula.length) {
    const character = formula[cursor];

    if (character === "\"") {
      nextFormula += character;

      if (inString && formula[cursor + 1] === "\"") {
        nextFormula += "\"";
        cursor += 2;
        continue;
      }

      inString = !inString;
      cursor += 1;
      continue;
    }

    if (inString) {
      nextFormula += character;
      cursor += 1;
      continue;
    }

    const remaining = formula.slice(cursor);
    const rangeMatch = remaining.match(
      /^((?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?)([A-Z]+)(\$?)(\d+):((?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?)([A-Z]+)(\$?)(\d+)/,
    );

    if (rangeMatch) {
      const [
        fullMatch,
        startSheetRef,
        startColumnDollar,
        startColumnLabel,
        startRowDollar,
        startRowText,
        endSheetRef,
        endColumnDollar,
        endColumnLabel,
        endRowDollar,
        endRowText,
      ] = rangeMatch;
      const nextStartSheetRef = renameSheetReferencePrefix(startSheetRef, previousSheetName, nextSheetName);
      const nextEndSheetRef = renameSheetReferencePrefix(endSheetRef, previousSheetName, nextSheetName);

      nextFormula +=
        nextStartSheetRef === startSheetRef && nextEndSheetRef === endSheetRef
          ? fullMatch
          : `${nextStartSheetRef ?? ""}${startColumnDollar}${startColumnLabel}${startRowDollar}${startRowText}:${nextEndSheetRef ?? ""}${endColumnDollar}${endColumnLabel}${endRowDollar}${endRowText}`;
      cursor += fullMatch.length;
      continue;
    }

    const refMatch = remaining.match(/^((?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?)([A-Z]+)(\$?)(\d+)/);
    if (!refMatch) {
      nextFormula += character;
      cursor += 1;
      continue;
    }

    const [fullMatch, sheetRef, columnDollar, columnLabel, rowDollar, rowText] = refMatch;
    const nextSheetRef = renameSheetReferencePrefix(sheetRef, previousSheetName, nextSheetName);

    nextFormula +=
      nextSheetRef === sheetRef
        ? fullMatch
        : `${nextSheetRef ?? ""}${columnDollar}${columnLabel}${rowDollar}${rowText}`;
    cursor += fullMatch.length;
  }

  return nextFormula;
}

function matchesFormulaReference(
  sheetRef: string | undefined,
  targetSheetName: string,
  includeUnqualifiedReferences: boolean,
  previousCharacter: string | undefined,
): boolean {
  if (!sheetRef) {
    return includeUnqualifiedReferences && !(previousCharacter && /[A-Za-z0-9_.]/.test(previousCharacter));
  }

  return matchesSheetReference(sheetRef, targetSheetName);
}

function matchesSheetReference(sheetRef: string | undefined, targetSheetName: string): boolean {
  if (!sheetRef) {
    return false;
  }

  const rawSheetName = sheetRef.slice(0, -1);
  const normalizedSheetName =
    rawSheetName.startsWith("'") && rawSheetName.endsWith("'")
      ? rawSheetName.slice(1, -1).replaceAll("''", "'")
      : rawSheetName;

  return normalizedSheetName === targetSheetName;
}

function renameSheetReferencePrefix(
  sheetRef: string | undefined,
  previousSheetName: string,
  nextSheetName: string,
): string | undefined {
  if (!matchesSheetReference(sheetRef, previousSheetName)) {
    return sheetRef;
  }

  return `${formatSheetReference(nextSheetName)}!`;
}

function formatSheetReference(sheetName: string): string {
  if (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheetName)) {
    return sheetName;
  }

  return `'${sheetName.replaceAll("'", "''")}'`;
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

function rewriteTableReferenceXml(
  tableXml: string,
  transformRange: (range: string) => string | null,
): string | null {
  const tableMatch = tableXml.match(/<table\b([^>]*?)>/);
  if (!tableMatch) {
    return tableXml;
  }

  const tableAttributes = parseAttributes(tableMatch[1]);
  const refIndex = tableAttributes.findIndex(([name]) => name === "ref");
  if (refIndex === -1) {
    return tableXml;
  }

  const currentRange = tableAttributes[refIndex]?.[1] ?? "";
  const nextRange = transformRange(currentRange);
  if (nextRange === null) {
    return null;
  }

  const nextTableAttributes = [...tableAttributes];
  nextTableAttributes[refIndex] = ["ref", nextRange];
  let nextTableXml =
    tableXml.slice(0, tableMatch.index) +
    `<table ${serializeAttributes(nextTableAttributes)}>` +
    tableXml.slice((tableMatch.index ?? 0) + tableMatch[0].length);

  nextTableXml = nextTableXml.replace(/<autoFilter\b([^>]*?)\/>/g, (match, attributesSource) => {
    const attributes = parseAttributes(attributesSource);
    const autoFilterRefIndex = attributes.findIndex(([name]) => name === "ref");

    if (autoFilterRefIndex === -1) {
      return match;
    }

    const autoFilterRange = attributes[autoFilterRefIndex]?.[1] ?? "";
    const nextAutoFilterRange = transformRange(autoFilterRange);
    if (nextAutoFilterRange === null) {
      return "";
    }

    const nextAttributes = [...attributes];
    nextAttributes[autoFilterRefIndex] = ["ref", nextAutoFilterRange];
    return `<autoFilter ${serializeAttributes(nextAttributes)}/>`;
  });

  return nextTableXml;
}

function removeTablePartsFromSheetXml(sheetXml: string, relationshipIds: string[]): string {
  const tablePartsMatch = sheetXml.match(/<tableParts\b[^>]*>([\s\S]*?)<\/tableParts>/);
  if (!tablePartsMatch || tablePartsMatch.index === undefined) {
    return sheetXml;
  }

  const keptTableParts = Array.from(
    tablePartsMatch[1].matchAll(/<tablePart\b([^>]*?)\/>/g),
    (match) => {
      const attributesSource = match[1];
      return {
        relationshipId: getXmlAttr(attributesSource, "r:id"),
        xml: `<tablePart${attributesSource ? ` ${attributesSource.trim()}` : ""}/>`,
      };
    },
  ).filter((tablePart) => tablePart.relationshipId && !relationshipIds.includes(tablePart.relationshipId));

  const nextTablePartsXml =
    keptTableParts.length === 0
      ? ""
      : `<tableParts count="${keptTableParts.length}">${keptTableParts.map((tablePart) => tablePart.xml).join("")}</tableParts>`;

  return (
    sheetXml.slice(0, tablePartsMatch.index) +
    nextTablePartsXml +
    sheetXml.slice(tablePartsMatch.index + tablePartsMatch[0].length)
  );
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

function assertInsertCount(count: number): void {
  if (!Number.isInteger(count) || count < 1) {
    throw new XlsxError(`Invalid insert count: ${count}`);
  }
}

const WORKSHEET_REF_TAGS = ["autoFilter", "sortState", "hyperlink"];
const WORKSHEET_SQREF_TAGS = [
  "conditionalFormatting",
  "dataValidation",
  "selection",
  "protectedRange",
  "ignoredError",
];
const WORKSHEET_CELL_REF_ATTRIBUTES: Array<[string, string]> = [
  ["selection", "activeCell"],
  ["pane", "topLeftCell"],
];
