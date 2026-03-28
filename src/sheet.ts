import { Cell } from "./cell.js";
import type {
  CellEntry,
  CellSnapshot,
  CellStyleDefinition,
  CellStylePatch,
  CellValue,
  DataValidation,
  FreezePane,
  Hyperlink,
  SetDataValidationOptions,
  SetFormulaOptions,
  SetHyperlinkOptions,
  SheetSelection,
} from "./types.js";
import { XlsxError } from "./errors.js";
import {
  buildSheetIndex,
  parseCellAddressFast,
  parseCellSnapshot,
  type LocatedCell,
  type LocatedRow,
  type SheetIndex,
} from "./sheet-index.js";
import type { Workbook } from "./workbook.js";
import { basenamePosix, dirnamePosix, resolvePosix } from "./utils/path.js";
import {
  decodeXmlText,
  escapeRegex,
  escapeXmlText,
  extractTagText,
  getXmlAttr,
  parseAttributes,
  serializeAttributes,
} from "./utils/xml.js";

interface TableReference {
  relationshipId: string;
  path: string;
}

interface UsedRangeBounds {
  minRow: number;
  maxRow: number;
  minColumn: number;
  maxColumn: number;
}

interface ColumnDefinition {
  min: number;
  max: number;
  attributes: Array<[string, string]>;
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

  cell(address: string): Cell;
  cell(rowNumber: number, column: number | string): Cell;
  cell(addressOrRowNumber: string | number, column?: number | string): Cell {
    const normalizedAddress = resolveCellAddress(addressOrRowNumber, column);
    let cell = this.cellHandles.get(normalizedAddress);

    if (!cell) {
      cell = new Cell(this, normalizedAddress);
      this.cellHandles.set(normalizedAddress, cell);
    }

    return cell;
  }

  getCell(address: string): CellValue;
  getCell(rowNumber: number, column: number | string): CellValue;
  getCell(addressOrRowNumber: string | number, column?: number | string): CellValue {
    if (typeof addressOrRowNumber === "number") {
      return this.readCellSnapshotByIndexes(addressOrRowNumber, column).value;
    }

    return this.readCellSnapshot(resolveCellAddress(addressOrRowNumber, column)).value;
  }

  getStyleId(address: string): number | null;
  getStyleId(rowNumber: number, column: number | string): number | null;
  getStyleId(addressOrRowNumber: string | number, column?: number | string): number | null {
    if (typeof addressOrRowNumber === "number") {
      return this.readCellSnapshotByIndexes(addressOrRowNumber, column).styleId;
    }

    return this.readCellSnapshot(resolveCellAddress(addressOrRowNumber, column)).styleId;
  }

  getStyle(address: string): CellStyleDefinition | null;
  getStyle(rowNumber: number, column: number | string): CellStyleDefinition | null;
  getStyle(addressOrRowNumber: string | number, column?: number | string): CellStyleDefinition | null {
    const styleId =
      typeof addressOrRowNumber === "number"
        ? this.readCellSnapshotByIndexes(addressOrRowNumber, column).styleId
        : this.readCellSnapshot(resolveCellAddress(addressOrRowNumber, column)).styleId;
    return this.workbook.getStyle(styleId ?? 0);
  }

  getColumnStyleId(column: number | string): number | null {
    const columnNumber = normalizeColumnNumber(column);
    return parseColumnStyleId(this.getSheetIndex().xml, columnNumber);
  }

  copyStyle(sourceAddress: string, targetAddress: string): void;
  copyStyle(
    sourceRowNumber: number,
    sourceColumn: number | string,
    targetRowNumber: number,
    targetColumn: number | string,
  ): void;
  copyStyle(
    sourceAddressOrRowNumber: string | number,
    sourceColumnOrTargetAddress: number | string,
    targetRowNumber?: number,
    targetColumn?: number | string,
  ): void {
    const { sourceAddress, targetAddress } = resolveCopyStyleArguments(
      sourceAddressOrRowNumber,
      sourceColumnOrTargetAddress,
      targetRowNumber,
      targetColumn,
    );
    this.setStyleId(targetAddress, this.getStyleId(sourceAddress));
  }

  cloneStyle(address: string, patch?: CellStylePatch): number;
  cloneStyle(rowNumber: number, column: number | string, patch?: CellStylePatch): number;
  cloneStyle(
    addressOrRowNumber: string | number,
    columnOrPatch?: number | string | CellStylePatch,
    patch?: CellStylePatch,
  ): number {
    const normalizedAddress = resolveCellAddress(
      addressOrRowNumber,
      typeof addressOrRowNumber === "number" ? (columnOrPatch as number | string) : undefined,
    );
    const nextPatch = resolveCloneStylePatch(addressOrRowNumber, columnOrPatch, patch);
    const nextStyleId = this.workbook.cloneStyle(this.getStyleId(normalizedAddress) ?? 0, nextPatch);
    this.setStyleId(normalizedAddress, nextStyleId);
    return nextStyleId;
  }

  rename(name: string): void {
    this.workbook.renameSheet(this.name, name);
  }

  getFormula(address: string): string | null;
  getFormula(rowNumber: number, column: number | string): string | null;
  getFormula(addressOrRowNumber: string | number, column?: number | string): string | null {
    if (typeof addressOrRowNumber === "number") {
      return this.readCellSnapshotByIndexes(addressOrRowNumber, column).formula;
    }

    return this.readCellSnapshot(resolveCellAddress(addressOrRowNumber, column)).formula;
  }

  get rowCount(): number {
    return this.getSheetIndex().usedBounds?.maxRow ?? 0;
  }

  get columnCount(): number {
    return this.getSheetIndex().usedBounds?.maxColumn ?? 0;
  }

  getHeaders(headerRowNumber = 1): string[] {
    assertRowNumber(headerRowNumber);
    return this.getRow(headerRowNumber).map((value) => (typeof value === "string" ? value : ""));
  }

  getRowStyleId(rowNumber: number): number | null {
    assertRowNumber(rowNumber);
    return parseRowStyleId(this.getSheetIndex().rows.get(rowNumber)?.attributesSource);
  }

  getRow(rowNumber: number): CellValue[] {
    assertRowNumber(rowNumber);

    const row = this.getSheetIndex().rows.get(rowNumber);
    if (!row || row.cells.length === 0) {
      return [];
    }

    const values: CellValue[] = [];
    const maxColumn = row.maxColumnNumber;

    for (let columnNumber = 1; columnNumber <= maxColumn; columnNumber += 1) {
      values.push(this.getCell(rowNumber, columnNumber));
    }

    return values;
  }

  getRowEntries(rowNumber: number): CellEntry[] {
    assertRowNumber(rowNumber);

    const row = this.getSheetIndex().rows.get(rowNumber);
    if (!row) {
      return [];
    }

    return row.cells.map((cell) => createCellEntry(cell));
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
      values.push(this.getCell(rowNumber, columnNumber));
    }

    return values;
  }

  getColumnEntries(column: number | string): CellEntry[] {
    const columnNumber = normalizeColumnNumber(column);
    const entries: CellEntry[] = [];
    const index = this.getSheetIndex();

    for (const rowNumber of index.rowNumbers) {
      const cell = index.rows.get(rowNumber)?.cellsByColumn[columnNumber];
      if (cell) {
        entries.push(createCellEntry(cell));
      }
    }

    return entries;
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

  getCellEntries(): CellEntry[] {
    return Array.from(this.iterCellEntries());
  }

  *iterCellEntries(): IterableIterator<CellEntry> {
    const index = this.getSheetIndex();

    for (const rowNumber of index.rowNumbers) {
      const row = index.rows.get(rowNumber);
      if (!row) {
        continue;
      }

      for (const cell of row.cells) {
        yield createCellEntry(cell);
      }
    }
  }

  getUsedRange(): string | null {
    return formatUsedRangeBounds(this.getSheetIndex().usedBounds);
  }

  getMergedRanges(): string[] {
    return parseMergedRanges(this.getSheetIndex().xml);
  }

  getAutoFilter(): string | null {
    return parseSheetAutoFilter(this.getSheetIndex().xml);
  }

  getFreezePane(): FreezePane | null {
    return parseSheetFreezePane(this.getSheetIndex().xml);
  }

  getSelection(): SheetSelection | null {
    return parseSheetSelection(this.getSheetIndex().xml);
  }

  getDataValidations(): DataValidation[] {
    return parseSheetDataValidations(this.getSheetIndex().xml);
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

  getHyperlinks(): Hyperlink[] {
    return parseSheetHyperlinks(this.getSheetIndex().xml, parseHyperlinkRelationshipTargets(this.readSheetRelationshipsXml()));
  }

  setHyperlink(address: string, target: string, options: SetHyperlinkOptions = {}): void {
    const normalizedAddress = normalizeCellAddress(address);
    if (options.text !== undefined) {
      this.setCell(normalizedAddress, options.text);
    }

    const currentRelationshipId = getHyperlinkRelationshipId(this.getSheetIndex().xml, normalizedAddress);
    let relationshipsXml = this.readSheetRelationshipsXml();
    let relationshipId: string | null = currentRelationshipId;

    if (target.startsWith("#")) {
      if (currentRelationshipId) {
        relationshipsXml = removeRelationshipById(relationshipsXml, currentRelationshipId);
      }

      this.writeSheetXml(
        upsertHyperlinkInSheetXml(
          this.getSheetIndex().xml,
          buildInternalHyperlinkXml(normalizedAddress, target, options.tooltip),
          normalizedAddress,
        ),
      );
      this.writeSheetRelationshipsXml(relationshipsXml);
      return;
    }

    relationshipId ??= getNextRelationshipIdFromXml(relationshipsXml);
    relationshipsXml = upsertRelationship(
      relationshipsXml,
      relationshipId,
      HYPERLINK_RELATIONSHIP_TYPE,
      target,
      "External",
    );
    this.writeSheetXml(
      upsertHyperlinkInSheetXml(
        this.getSheetIndex().xml,
        buildExternalHyperlinkXml(normalizedAddress, relationshipId, options.tooltip),
        normalizedAddress,
      ),
    );
    this.writeSheetRelationshipsXml(relationshipsXml);
  }

  removeHyperlink(address: string): void {
    const normalizedAddress = normalizeCellAddress(address);
    const currentRelationshipId = getHyperlinkRelationshipId(this.getSheetIndex().xml, normalizedAddress);
    const nextSheetXml = removeHyperlinkFromSheetXml(this.getSheetIndex().xml, normalizedAddress);

    if (nextSheetXml !== this.getSheetIndex().xml) {
      this.writeSheetXml(nextSheetXml);
    }

    if (currentRelationshipId) {
      this.writeSheetRelationshipsXml(removeRelationshipById(this.readSheetRelationshipsXml(), currentRelationshipId));
    }
  }

  setAutoFilter(range: string): void {
    const normalizedRange = normalizeRangeRef(range);
    const nextSheetXml = upsertAutoFilterInSheetXml(this.getSheetIndex().xml, normalizedRange);

    if (nextSheetXml !== this.getSheetIndex().xml) {
      this.writeSheetXml(nextSheetXml);
    }
  }

  removeAutoFilter(): void {
    const nextSheetXml = removeAutoFilterFromSheetXml(this.getSheetIndex().xml);

    if (nextSheetXml !== this.getSheetIndex().xml) {
      this.writeSheetXml(nextSheetXml);
    }
  }

  freezePane(columnCount: number, rowCount = 0): void {
    assertFreezeSplit(columnCount, rowCount);
    const nextSheetXml = upsertFreezePaneInSheetXml(this.getSheetIndex().xml, columnCount, rowCount);

    if (nextSheetXml !== this.getSheetIndex().xml) {
      this.writeSheetXml(nextSheetXml);
    }
  }

  unfreezePane(): void {
    const nextSheetXml = removeFreezePaneFromSheetXml(this.getSheetIndex().xml);

    if (nextSheetXml !== this.getSheetIndex().xml) {
      this.writeSheetXml(nextSheetXml);
    }
  }

  setSelection(activeCell: string, range = activeCell): void {
    const normalizedActiveCell = normalizeCellAddress(activeCell);
    const normalizedRange = normalizeSqref(range);
    const nextSheetXml = upsertSheetSelectionInSheetXml(
      this.getSheetIndex().xml,
      normalizedActiveCell,
      normalizedRange,
    );

    if (nextSheetXml !== this.getSheetIndex().xml) {
      this.writeSheetXml(nextSheetXml);
    }
  }

  setDataValidation(range: string, options: SetDataValidationOptions = {}): void {
    const normalizedRange = normalizeSqref(range);
    const nextSheetXml = upsertDataValidationInSheetXml(
      this.getSheetIndex().xml,
      buildDataValidationXml(normalizedRange, options),
      normalizedRange,
    );

    if (nextSheetXml !== this.getSheetIndex().xml) {
      this.writeSheetXml(nextSheetXml);
    }
  }

  removeDataValidation(range: string): void {
    const normalizedRange = normalizeSqref(range);
    const nextSheetXml = removeDataValidationFromSheetXml(this.getSheetIndex().xml, normalizedRange);

    if (nextSheetXml !== this.getSheetIndex().xml) {
      this.writeSheetXml(nextSheetXml);
    }
  }

  addTable(
    range: string,
    options: { name?: string } = {},
  ): { name: string; displayName: string; range: string; path: string } {
    const normalizedRange = normalizeRangeRef(range);
    const existingTables = this.getTables();
    const name = options.name ?? getNextTableName(this.workbook.listEntries());
    assertTableName(name);

    if (existingTables.some((table) => table.name === name || table.displayName === name)) {
      throw new XlsxError(`Table already exists: ${name}`);
    }

    const tablePath = getNextTablePath(this.workbook.listEntries());
    const tableId = getNextTableId(this.workbook.listEntries(), this.workbook);
    const relationshipId = getNextRelationshipIdFromXml(this.readSheetRelationshipsXml());
    const tableXml = buildTableXml(normalizedRange, tableId, name, this.getRange(normalizedRange)[0] ?? []);

    this.workbook.writeEntryText(tablePath, tableXml);
    this.writeSheetRelationshipsXml(
      appendRelationship(
        this.readSheetRelationshipsXml(),
        relationshipId,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table",
        makeRelativeSheetRelationshipTarget(this.path, tablePath),
      ),
    );
    this.writeSheetXml(appendTablePart(this.getSheetIndex().xml, relationshipId));
    this.writeContentTypesXml(addContentTypeOverride(this.readContentTypesXml(), tablePath, TABLE_CONTENT_TYPE));

    return {
      name,
      displayName: name,
      range: normalizedRange,
      path: tablePath,
    };
  }

  removeTable(name: string): void {
    const tableReference = this.getTableReferences().find((table) => {
      const tableXml = this.workbook.readEntryText(table.path);
      const tableTagMatch = tableXml.match(/<table\b([^>]*?)>/);
      const attributesSource = tableTagMatch?.[1] ?? "";
      return getXmlAttr(attributesSource, "name") === name || getXmlAttr(attributesSource, "displayName") === name;
    });

    if (!tableReference) {
      throw new XlsxError(`Table not found: ${name}`);
    }

    this.writeSheetXml(removeTablePartsFromSheetXml(this.getSheetIndex().xml, [tableReference.relationshipId]));
    this.writeSheetRelationshipsXml(removeRelationshipById(this.readSheetRelationshipsXml(), tableReference.relationshipId));
    this.writeContentTypesXml(removeContentTypeOverride(this.readContentTypesXml(), tableReference.path));
    this.workbook.removeEntry(tableReference.path);
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
    nextSheetXml = transformColumnStyleDefinitions(nextSheetXml, columnNumber, count, "shift");
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
    nextSheetXml = transformColumnStyleDefinitions(nextSheetXml, columnNumber, count, "delete");
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

  setCell(address: string, value: CellValue): void;
  setCell(rowNumber: number, column: number | string, value: CellValue): void;
  setCell(addressOrRowNumber: string | number, columnOrValue: number | string | CellValue, value?: CellValue): void {
    const normalizedAddress = resolveCellAddress(addressOrRowNumber, typeof addressOrRowNumber === "number" ? columnOrValue as number | string : undefined);
    const existingCell = this.getSheetIndex().cells.get(normalizedAddress);
    const nextValue = resolveSetCellValue(addressOrRowNumber, columnOrValue, value);
    this.writeCellXml(
      normalizedAddress,
      buildValueCellXml(normalizedAddress, nextValue, existingCell?.attributesSource),
    );
  }

  setStyleId(address: string, styleId: number | null): void;
  setStyleId(rowNumber: number, column: number | string, styleId: number | null): void;
  setStyleId(
    addressOrRowNumber: string | number,
    columnOrStyleId: number | string | null,
    styleId?: number | null,
  ): void {
    const normalizedAddress = resolveCellAddress(
      addressOrRowNumber,
      typeof addressOrRowNumber === "number" ? (columnOrStyleId as number | string) : undefined,
    );
    const nextStyleId = resolveSetStyleId(addressOrRowNumber, columnOrStyleId, styleId);
    const index = this.getSheetIndex();
    const existingCell = index.cells.get(normalizedAddress);

    this.writeCellXml(
      normalizedAddress,
      buildStyledCellXml(
        normalizedAddress,
        nextStyleId,
        existingCell?.attributesSource,
        existingCell ? index.xml.slice(existingCell.start, existingCell.end) : undefined,
      ),
    );
  }

  setColumnStyleId(column: number | string, styleId: number | null): void {
    const columnNumber = normalizeColumnNumber(column);
    assertStyleId(styleId);

    const nextSheetXml = updateColumnStyleIdInSheetXml(
      this.getSheetIndex().xml,
      columnNumber,
      styleId,
    );

    if (nextSheetXml !== this.getSheetIndex().xml) {
      this.writeSheetXml(nextSheetXml);
    }
  }

  deleteCell(address: string): void;
  deleteCell(rowNumber: number, column: number | string): void;
  deleteCell(addressOrRowNumber: string | number, column?: number | string): void {
    const normalizedAddress = resolveCellAddress(addressOrRowNumber, column);
    const index = this.getSheetIndex();
    const existingCell = index.cells.get(normalizedAddress);

    if (!existingCell) {
      return;
    }

    const row = index.rows.get(existingCell.rowNumber);
    if (row && row.cells.length === 1) {
      const nextRowXml = normalizeEmptyRowXml(
        index.xml.slice(row.start, existingCell.start) + index.xml.slice(existingCell.end, row.end),
      );
      this.writeSheetXml(index.xml.slice(0, row.start) + nextRowXml + index.xml.slice(row.end));
      return;
    }

    this.writeSheetXml(index.xml.slice(0, existingCell.start) + index.xml.slice(existingCell.end));
  }

  setFormula(address: string, formula: string, options?: SetFormulaOptions): void;
  setFormula(rowNumber: number, column: number | string, formula: string, options?: SetFormulaOptions): void;
  setFormula(
    addressOrRowNumber: string | number,
    columnOrFormula: number | string,
    formulaOrOptions?: string | SetFormulaOptions,
    options: SetFormulaOptions = {},
  ): void {
    const normalizedAddress = resolveCellAddress(addressOrRowNumber, typeof addressOrRowNumber === "number" ? columnOrFormula as number | string : undefined);
    const existingCell = this.getSheetIndex().cells.get(normalizedAddress);
    const { formula, formulaOptions } = resolveSetFormulaArguments(
      addressOrRowNumber,
      columnOrFormula,
      formulaOrOptions,
      options,
    );
    this.writeCellXml(
      normalizedAddress,
      buildFormulaCellXml(
        normalizedAddress,
        formula,
        formulaOptions.cachedValue ?? null,
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

  setRowStyleId(rowNumber: number, styleId: number | null): void {
    assertRowNumber(rowNumber);
    assertStyleId(styleId);

    const index = this.getSheetIndex();
    const row = index.rows.get(rowNumber);

    if (!row) {
      if (styleId === null) {
        return;
      }

      const insertionIndex = findRowInsertionIndex(index, rowNumber);
      this.writeSheetXml(
        index.xml.slice(0, insertionIndex) +
          buildEmptyStyledRowXml(rowNumber, styleId) +
          index.xml.slice(insertionIndex),
      );
      return;
    }

    this.writeSheetXml(
      index.xml.slice(0, row.start) +
        buildStyledRowXml(index.xml, row, styleId) +
        index.xml.slice(row.end),
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

  readCellSnapshot(address: string): CellSnapshot {
    const locatedCell = this.getSheetIndex().cells.get(normalizeCellAddress(address));
    return parseCellSnapshot(locatedCell);
  }

  private readCellSnapshotByIndexes(
    rowNumber: number,
    column: number | string | undefined,
  ): CellSnapshot {
    assertRowNumber(rowNumber);
    if (column === undefined) {
      throw new XlsxError(`Missing column index for row: ${rowNumber}`);
    }

    const columnNumber = normalizeColumnNumber(column);
    const row = this.getSheetIndex().rows.get(rowNumber);
    return parseCellSnapshot(row?.cellsByColumn[columnNumber]);
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

    this.sheetIndex = buildSheetIndex(this.workbook, this.workbook.readEntryText(this.path));
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

    const relationshipsPath = getSheetRelationshipsPath(this.path);
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

  private readSheetRelationshipsXml(): string {
    const relationshipsPath = getSheetRelationshipsPath(this.path);
    return this.workbook.listEntries().includes(relationshipsPath)
      ? this.workbook.readEntryText(relationshipsPath)
      : EMPTY_RELATIONSHIPS_XML;
  }

  private writeSheetRelationshipsXml(relationshipsXml: string): void {
    this.workbook.writeEntryText(getSheetRelationshipsPath(this.path), relationshipsXml);
  }

  private readContentTypesXml(): string {
    return this.workbook.readEntryText("[Content_Types].xml");
  }

  private writeContentTypesXml(contentTypesXml: string): void {
    this.workbook.writeEntryText("[Content_Types].xml", contentTypesXml);
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
    const removedTables: TableReference[] = [];

    for (const table of this.getTableReferences()) {
      const tableXml = this.workbook.readEntryText(table.path);
      const nextTableXml = rewriteTableReferenceXml(tableXml, transformRange);

      if (nextTableXml === null) {
        removedTables.push(table);
        continue;
      }

      if (nextTableXml !== tableXml) {
        this.workbook.writeEntryText(table.path, nextTableXml);
      }
    }

    if (removedTables.length === 0) {
      return;
    }

    const removedRelationshipIds = removedTables.map((table) => table.relationshipId);
    const nextSheetXml = removeTablePartsFromSheetXml(this.getSheetIndex().xml, removedRelationshipIds);
    if (nextSheetXml !== this.getSheetIndex().xml) {
      this.writeSheetXml(nextSheetXml);
    }

    this.writeSheetRelationshipsXml(
      removedRelationshipIds.reduce(
        (relationshipsXml, relationshipId) => removeRelationshipById(relationshipsXml, relationshipId),
        this.readSheetRelationshipsXml(),
      ),
    );

    let nextContentTypesXml = this.readContentTypesXml();
    for (const table of removedTables) {
      nextContentTypesXml = removeContentTypeOverride(nextContentTypesXml, table.path);
      this.workbook.removeEntry(table.path);
    }
    this.writeContentTypesXml(nextContentTypesXml);
  }

  private writeSheetXml(nextSheetXml: string): void {
    const indexedSheet = buildSheetIndex(this.workbook, nextSheetXml);
    const normalizedSheetXml = updateDimensionRef(indexedSheet);

    this.workbook.writeEntryText(this.path, normalizedSheetXml);
    this.sheetIndex = buildSheetIndex(this.workbook, normalizedSheetXml);
    this.revision += 1;
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

function buildStyledCellXml(
  address: string,
  styleId: number | null,
  existingAttributesSource?: string,
  existingCellXml?: string,
): string {
  const serializedAttributes = serializeAttributes(
    buildCellAttributesWithStyle(address, styleId, existingAttributesSource),
  );

  if (!existingCellXml || existingCellXml.endsWith("/>")) {
    return `<c ${serializedAttributes}/>`;
  }

  const openTagEnd = existingCellXml.indexOf(">");
  if (openTagEnd === -1) {
    throw new XlsxError("Cell XML is missing opening tag");
  }

  return `<c ${serializedAttributes}>${existingCellXml.slice(openTagEnd + 1)}`;
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

function buildCellAttributesWithStyle(
  address: string,
  styleId: number | null,
  existingAttributesSource = "",
): Array<[string, string]> {
  const attributes = parseAttributes(existingAttributesSource);
  const preserved = attributes.filter(([name]) => name !== "r" && name !== "s");
  const nextAttributes: Array<[string, string]> = [["r", address]];

  if (styleId !== null) {
    nextAttributes.push(["s", String(styleId)]);
  }

  nextAttributes.push(...preserved);
  return nextAttributes;
}

function parseRowStyleId(attributesSource: string | undefined): number | null {
  if (!attributesSource) {
    return null;
  }

  const styleId = getXmlAttr(attributesSource, "s");
  return styleId === undefined ? null : Number(styleId);
}

function parseColumnStyleId(sheetXml: string, columnNumber: number): number | null {
  let styleId: number | null = null;

  for (const definition of parseColumnDefinitions(sheetXml)) {
    if (columnNumber < definition.min || columnNumber > definition.max) {
      continue;
    }

    const styleText = getXmlAttr(serializeAttributes(definition.attributes), "style");
    styleId = styleText === undefined ? null : Number(styleText);
  }

  return styleId;
}

function buildStyledRowXml(sheetXml: string, row: LocatedRow, styleId: number | null): string {
  const serializedAttributes = serializeAttributes(
    buildRowAttributesWithStyle(row.rowNumber, styleId, row.attributesSource),
  );

  if (row.selfClosing) {
    return `<row ${serializedAttributes}/>`;
  }

  return `<row ${serializedAttributes}>${sheetXml.slice(row.innerStart, row.innerEnd)}</row>`;
}

function buildEmptyStyledRowXml(rowNumber: number, styleId: number): string {
  return `<row ${serializeAttributes(buildRowAttributesWithStyle(rowNumber, styleId))}/>`;
}

function buildRowAttributesWithStyle(
  rowNumber: number,
  styleId: number | null,
  existingAttributesSource = "",
): Array<[string, string]> {
  const attributes = parseAttributes(existingAttributesSource);
  const preserved = attributes.filter(
    ([name]) => name !== "r" && name !== "s" && name !== "customFormat",
  );
  const nextAttributes: Array<[string, string]> = [["r", String(rowNumber)]];

  if (styleId !== null) {
    nextAttributes.push(["s", String(styleId)], ["customFormat", "1"]);
  }

  nextAttributes.push(...preserved);
  return nextAttributes;
}

function parseColumnDefinitions(sheetXml: string): ColumnDefinition[] {
  const colsMatch = sheetXml.match(/<cols\b[^>]*>([\s\S]*?)<\/cols>/);
  if (!colsMatch) {
    return [];
  }

  return Array.from(colsMatch[1].matchAll(/<col\b([^>]*?)\/>/g), (match) => {
    const attributes = parseAttributes(match[1]);
    const min = Number(getXmlAttr(match[1], "min") ?? "0");
    const max = Number(getXmlAttr(match[1], "max") ?? "0");

    return {
      min,
      max,
      attributes,
    };
  }).filter((definition) => Number.isInteger(definition.min) && Number.isInteger(definition.max) && definition.min > 0 && definition.max >= definition.min);
}

function updateColumnStyleIdInSheetXml(
  sheetXml: string,
  columnNumber: number,
  styleId: number | null,
): string {
  const existingDefinitions = parseColumnDefinitions(sheetXml);
  if (existingDefinitions.length === 0 && styleId === null) {
    return sheetXml;
  }

  const nextDefinitions: ColumnDefinition[] = [];
  let handled = false;

  for (const definition of existingDefinitions) {
    if (columnNumber < definition.min || columnNumber > definition.max) {
      nextDefinitions.push(definition);
      continue;
    }

    handled = true;
    if (definition.min < columnNumber) {
      nextDefinitions.push(buildColumnDefinition(definition.min, columnNumber - 1, definition.attributes));
    }

    const styledDefinition = buildColumnDefinitionWithStyle(columnNumber, columnNumber, definition.attributes, styleId);
    if (styledDefinition) {
      nextDefinitions.push(styledDefinition);
    }

    if (columnNumber < definition.max) {
      nextDefinitions.push(buildColumnDefinition(columnNumber + 1, definition.max, definition.attributes));
    }
  }

  if (!handled && styleId !== null) {
    nextDefinitions.push(buildColumnDefinitionWithStyle(columnNumber, columnNumber, [], styleId)!);
  }

  return replaceColumnDefinitions(sheetXml, normalizeColumnDefinitions(nextDefinitions));
}

function transformColumnStyleDefinitions(
  sheetXml: string,
  targetColumnNumber: number,
  count: number,
  mode: "shift" | "delete",
): string {
  const existingDefinitions = parseColumnDefinitions(sheetXml);
  if (existingDefinitions.length === 0) {
    return sheetXml;
  }

  const nextDefinitions: ColumnDefinition[] = [];

  for (const definition of existingDefinitions) {
    if (mode === "shift") {
      if (definition.max < targetColumnNumber) {
        nextDefinitions.push(definition);
        continue;
      }

      if (definition.min >= targetColumnNumber) {
        nextDefinitions.push(buildColumnDefinition(definition.min + count, definition.max + count, definition.attributes));
        continue;
      }

      nextDefinitions.push(buildColumnDefinition(definition.min, targetColumnNumber - 1, definition.attributes));
      nextDefinitions.push(buildColumnDefinition(targetColumnNumber + count, definition.max + count, definition.attributes));
      continue;
    }

    const deleteEnd = targetColumnNumber + count - 1;
    if (definition.max < targetColumnNumber) {
      nextDefinitions.push(definition);
      continue;
    }

    if (definition.min > deleteEnd) {
      nextDefinitions.push(buildColumnDefinition(definition.min - count, definition.max - count, definition.attributes));
      continue;
    }

    if (definition.min < targetColumnNumber) {
      nextDefinitions.push(buildColumnDefinition(definition.min, targetColumnNumber - 1, definition.attributes));
    }

    if (definition.max > deleteEnd) {
      nextDefinitions.push(buildColumnDefinition(targetColumnNumber, definition.max - count, definition.attributes));
    }
  }

  return replaceColumnDefinitions(sheetXml, normalizeColumnDefinitions(nextDefinitions));
}

function replaceColumnDefinitions(sheetXml: string, definitions: ColumnDefinition[]): string {
  const colsMatch = sheetXml.match(/<cols\b[^>]*>[\s\S]*?<\/cols>/);
  const colsXml =
    definitions.length === 0
      ? ""
      : `<cols>${definitions.map((definition) => serializeColumnDefinition(definition)).join("")}</cols>`;

  if (colsMatch && colsMatch.index !== undefined) {
    return (
      sheetXml.slice(0, colsMatch.index) +
      colsXml +
      sheetXml.slice(colsMatch.index + colsMatch[0].length)
    );
  }

  if (definitions.length === 0) {
    return sheetXml;
  }

  const insertionIndex = findWorksheetChildInsertionIndex(sheetXml, COLS_FOLLOWING_TAGS);
  return sheetXml.slice(0, insertionIndex) + colsXml + sheetXml.slice(insertionIndex);
}

function normalizeColumnDefinitions(definitions: ColumnDefinition[]): ColumnDefinition[] {
  const filtered = definitions
    .filter((definition) => definition.min <= definition.max)
    .sort((left, right) => left.min - right.min || left.max - right.max);
  const merged: ColumnDefinition[] = [];

  for (const definition of filtered) {
    const previous = merged.at(-1);
    if (
      previous &&
      previous.max + 1 === definition.min &&
      haveEquivalentColumnDefinitionAttributes(previous.attributes, definition.attributes)
    ) {
      previous.max = definition.max;
      continue;
    }

    merged.push({
      min: definition.min,
      max: definition.max,
      attributes: [...definition.attributes],
    });
  }

  return merged;
}

function buildColumnDefinition(
  min: number,
  max: number,
  existingAttributes: Array<[string, string]>,
): ColumnDefinition {
  const preserved = existingAttributes.filter(([name]) => name !== "min" && name !== "max");
  return {
    min,
    max,
    attributes: [["min", String(min)], ["max", String(max)], ...preserved],
  };
}

function buildColumnDefinitionWithStyle(
  min: number,
  max: number,
  existingAttributes: Array<[string, string]>,
  styleId: number | null,
): ColumnDefinition | null {
  const preserved = existingAttributes.filter(
    ([name]) => name !== "min" && name !== "max" && name !== "style",
  );

  if (styleId === null && preserved.length === 0) {
    return null;
  }

  const attributes: Array<[string, string]> = [["min", String(min)], ["max", String(max)]];
  if (styleId !== null) {
    attributes.push(["style", String(styleId)]);
  }
  attributes.push(...preserved);

  return { min, max, attributes };
}

function serializeColumnDefinition(definition: ColumnDefinition): string {
  const attributes = definition.attributes.map(([name, value]) => {
    if (name === "min") {
      return [name, String(definition.min)] as [string, string];
    }
    if (name === "max") {
      return [name, String(definition.max)] as [string, string];
    }
    return [name, value] as [string, string];
  });

  return `<col ${serializeAttributes(attributes)}/>`;
}

function haveEquivalentColumnDefinitionAttributes(
  left: Array<[string, string]>,
  right: Array<[string, string]>,
): boolean {
  const normalize = (attributes: Array<[string, string]>) =>
    serializeAttributes(attributes.filter(([name]) => name !== "min" && name !== "max"));

  return normalize(left) === normalize(right);
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

function splitCellAddress(address: string): { rowNumber: number; columnNumber: number } {
  return parseCellAddressFast(address);
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

function resolveCellAddress(addressOrRowNumber: string | number, column?: number | string): string {
  if (typeof addressOrRowNumber === "string") {
    if (column !== undefined) {
      throw new XlsxError("Column argument is not allowed when address is a string");
    }

    return normalizeCellAddress(addressOrRowNumber);
  }

  assertRowNumber(addressOrRowNumber);
  if (column === undefined) {
    throw new XlsxError(`Missing column index for row: ${addressOrRowNumber}`);
  }

  return makeCellAddress(addressOrRowNumber, normalizeColumnNumber(column));
}

function resolveSetCellValue(
  addressOrRowNumber: string | number,
  columnOrValue: number | string | CellValue,
  value: CellValue | undefined,
): CellValue {
  if (typeof addressOrRowNumber === "string") {
    return columnOrValue as CellValue;
  }

  if (value === undefined) {
    throw new XlsxError(`Missing cell value for row ${addressOrRowNumber}`);
  }

  return value;
}

function resolveSetFormulaArguments(
  addressOrRowNumber: string | number,
  columnOrFormula: number | string,
  formulaOrOptions: string | SetFormulaOptions | undefined,
  options: SetFormulaOptions,
): { formula: string; formulaOptions: SetFormulaOptions } {
  if (typeof addressOrRowNumber === "string") {
    if (typeof columnOrFormula !== "string") {
      throw new XlsxError(`Invalid formula: ${String(columnOrFormula)}`);
    }

    return {
      formula: columnOrFormula,
      formulaOptions: (formulaOrOptions as SetFormulaOptions | undefined) ?? {},
    };
  }

  if (typeof formulaOrOptions !== "string") {
    throw new XlsxError(`Missing formula for row ${addressOrRowNumber}`);
  }

  return {
    formula: formulaOrOptions,
    formulaOptions: options,
  };
}

function normalizeRangeRef(range: string): string {
  const { startRow, endRow, startColumn, endColumn } = parseRangeRef(range);
  return formatRangeRef(startRow, startColumn, endRow, endColumn);
}

function normalizeSqref(rangeList: string): string {
  const ranges = rangeList
    .trim()
    .split(/\s+/)
    .filter((range) => range.length > 0)
    .map((range) => normalizeRangeRef(range));

  if (ranges.length === 0) {
    throw new XlsxError(`Invalid sqref: ${rangeList}`);
  }

  return ranges.join(" ");
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
    `<${escapeRegex(tagName)}\\b([^>]*?)(\\/>|>[\\s\\S]*?<\\/${escapeRegex(tagName)}>)`,
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

function createCellEntry(cell: LocatedCell): CellEntry {
  return {
    address: cell.address,
    rowNumber: cell.rowNumber,
    columnNumber: cell.columnNumber,
    ...cell.snapshot,
  };
}

function normalizeEmptyRowXml(rowXml: string): string {
  return rowXml.replace(/>\s*<\/row>$/, "></row>");
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

function parseSheetAutoFilter(sheetXml: string): string | null {
  const autoFilterMatch = sheetXml.match(/<autoFilter\b([^>]*?)\/>/);
  if (!autoFilterMatch) {
    return null;
  }

  const ref = getXmlAttr(autoFilterMatch[1], "ref");
  return ref ? normalizeRangeRef(ref) : null;
}

function parseSheetFreezePane(sheetXml: string): FreezePane | null {
  const paneMatch = sheetXml.match(/<pane\b([^>]*?)\/>/);
  if (!paneMatch) {
    return null;
  }

  const attributesSource = paneMatch[1];
  const state = getXmlAttr(attributesSource, "state");
  if (state && state !== "frozen" && state !== "frozenSplit") {
    return null;
  }

  const columnCount = Number(getXmlAttr(attributesSource, "xSplit") ?? "0");
  const rowCount = Number(getXmlAttr(attributesSource, "ySplit") ?? "0");
  if ((!Number.isFinite(columnCount) || columnCount < 0) && (!Number.isFinite(rowCount) || rowCount < 0)) {
    return null;
  }

  if (columnCount === 0 && rowCount === 0) {
    return null;
  }

  return {
    columnCount: Number.isFinite(columnCount) ? columnCount : 0,
    rowCount: Number.isFinite(rowCount) ? rowCount : 0,
    topLeftCell: getXmlAttr(attributesSource, "topLeftCell") ?? makeCellAddress(rowCount + 1, columnCount + 1),
    activePane: normalizePaneName(getXmlAttr(attributesSource, "activePane")),
  };
}

function parseSheetSelection(sheetXml: string): SheetSelection | null {
  const freezePane = parseSheetFreezePane(sheetXml);
  const selections = parseSheetSelectionEntries(sheetXml);
  if (selections.length === 0) {
    return null;
  }

  const targetPane = freezePane?.activePane ?? null;
  const selection =
    selections.find((candidate) => candidate.pane === targetPane) ??
    selections.find((candidate) => candidate.activeCell !== null || candidate.range !== null) ??
    selections[0];

  return selection ?? null;
}

function upsertFreezePaneInSheetXml(sheetXml: string, columnCount: number, rowCount: number): string {
  const paneXml = buildFreezePaneXml(columnCount, rowCount);
  const selectionsXml = buildFreezePaneSelectionsXml(columnCount, rowCount);
  const sheetViewsMatch = sheetXml.match(/<sheetViews\b[^>]*>([\s\S]*?)<\/sheetViews>/);

  if (!sheetViewsMatch || sheetViewsMatch.index === undefined) {
    const insertionIndex = findWorksheetChildInsertionIndex(sheetXml, SHEET_VIEWS_FOLLOWING_TAGS);
    return (
      sheetXml.slice(0, insertionIndex) +
      `<sheetViews><sheetView workbookViewId="0">${paneXml}${selectionsXml}</sheetView></sheetViews>` +
      sheetXml.slice(insertionIndex)
    );
  }

  const sheetViewMatch = sheetViewsMatch[1].match(/<sheetView\b([^>]*?)(?:\/>|>([\s\S]*?)<\/sheetView>)/);
  if (!sheetViewMatch || sheetViewMatch.index === undefined) {
    return (
      sheetXml.slice(0, sheetViewsMatch.index) +
      `<sheetViews><sheetView workbookViewId="0">${paneXml}${selectionsXml}</sheetView></sheetViews>` +
      sheetXml.slice(sheetViewsMatch.index + sheetViewsMatch[0].length)
    );
  }

  const attributes = parseAttributes(sheetViewMatch[1]);
  if (!attributes.some(([name]) => name === "workbookViewId")) {
    attributes.push(["workbookViewId", "0"]);
  }

  const innerXml = sheetViewMatch[2] ?? "";
  const cleanedInnerXml = innerXml
    .replace(/<pane\b[^>]*\/>/g, "")
    .replace(/<selection\b[^>]*\/>/g, "");
  const nextSheetViewXml = `<sheetView ${serializeAttributes(attributes)}>${paneXml}${selectionsXml}${cleanedInnerXml}</sheetView>`;
  const relativeStart = sheetViewMatch.index;
  const absoluteStart = sheetViewsMatch.index + sheetViewsMatch[0].indexOf(sheetViewMatch[0], relativeStart);

  return (
    sheetXml.slice(0, absoluteStart) +
    nextSheetViewXml +
    sheetXml.slice(absoluteStart + sheetViewMatch[0].length)
  );
}

function removeFreezePaneFromSheetXml(sheetXml: string): string {
  const sheetViewsMatch = sheetXml.match(/<sheetViews\b[^>]*>([\s\S]*?)<\/sheetViews>/);
  if (!sheetViewsMatch || sheetViewsMatch.index === undefined) {
    return sheetXml;
  }

  const sheetViewMatch = sheetViewsMatch[1].match(/<sheetView\b([^>]*?)(?:\/>|>([\s\S]*?)<\/sheetView>)/);
  if (!sheetViewMatch || sheetViewMatch.index === undefined) {
    return sheetXml;
  }

  const attributes = parseAttributes(sheetViewMatch[1]);
  if (!attributes.some(([name]) => name === "workbookViewId")) {
    attributes.push(["workbookViewId", "0"]);
  }

  const innerXml = sheetViewMatch[2] ?? "";
  const paneMatch = innerXml.match(/<pane\b([^>]*?)\/>/);
  if (!paneMatch) {
    return sheetXml;
  }

  const activePane = normalizePaneName(getXmlAttr(paneMatch[1], "activePane"));
  const selections = Array.from(innerXml.matchAll(/<selection\b([^>]*?)\/>/g), (match) => ({
    xml: match[0],
    attributes: parseAttributes(match[1]),
  }));
  const preferredSelection =
    selections.find((selection) => selection.attributes.find(([name]) => name === "pane")?.[1] === activePane) ??
    selections.find((selection) => selection.attributes.some(([name]) => name === "activeCell" || name === "sqref")) ??
    selections[0];
  const nextSelectionXml = preferredSelection
    ? buildSelectionXml(preferredSelection.attributes.filter(([name]) => name !== "pane"))
    : "";
  const cleanedInnerXml = innerXml
    .replace(/<pane\b[^>]*\/>/g, "")
    .replace(/<selection\b[^>]*\/>/g, "");
  const nextSheetViewXml = `<sheetView ${serializeAttributes(attributes)}>${nextSelectionXml}${cleanedInnerXml}</sheetView>`;
  const relativeStart = sheetViewMatch.index;
  const absoluteStart = sheetViewsMatch.index + sheetViewsMatch[0].indexOf(sheetViewMatch[0], relativeStart);

  return (
    sheetXml.slice(0, absoluteStart) +
    nextSheetViewXml +
    sheetXml.slice(absoluteStart + sheetViewMatch[0].length)
  );
}

function upsertSheetSelectionInSheetXml(
  sheetXml: string,
  activeCell: string,
  range: string,
): string {
  const freezePane = parseSheetFreezePane(sheetXml);
  const targetPane = freezePane?.activePane ?? null;
  const nextSelectionXml = buildSelectionXml([
    ...(targetPane ? [["pane", targetPane] as [string, string]] : []),
    ["activeCell", activeCell],
    ["sqref", range],
  ]);
  const sheetViewsMatch = sheetXml.match(/<sheetViews\b[^>]*>([\s\S]*?)<\/sheetViews>/);

  if (!sheetViewsMatch || sheetViewsMatch.index === undefined) {
    const insertionIndex = findWorksheetChildInsertionIndex(sheetXml, SHEET_VIEWS_FOLLOWING_TAGS);
    return (
      sheetXml.slice(0, insertionIndex) +
      `<sheetViews><sheetView workbookViewId="0">${nextSelectionXml}</sheetView></sheetViews>` +
      sheetXml.slice(insertionIndex)
    );
  }

  const sheetViewMatch = sheetViewsMatch[1].match(/<sheetView\b([^>]*?)(?:\/>|>([\s\S]*?)<\/sheetView>)/);
  if (!sheetViewMatch || sheetViewMatch.index === undefined) {
    return (
      sheetXml.slice(0, sheetViewsMatch.index) +
      `<sheetViews><sheetView workbookViewId="0">${nextSelectionXml}</sheetView></sheetViews>` +
      sheetXml.slice(sheetViewsMatch.index + sheetViewsMatch[0].length)
    );
  }

  const attributes = parseAttributes(sheetViewMatch[1]);
  if (!attributes.some(([name]) => name === "workbookViewId")) {
    attributes.push(["workbookViewId", "0"]);
  }

  const innerXml = sheetViewMatch[2] ?? "";
  let replaced = false;
  const nextInnerXml =
    innerXml.replace(/<selection\b([^>]*?)\/>/g, (match, attributesSource) => {
      const selectionPane = normalizePaneName(getXmlAttr(attributesSource, "pane"));
      const matchesTargetPane = selectionPane === targetPane;

      if (matchesTargetPane || (!replaced && targetPane === null && selectionPane === null)) {
        replaced = true;
        return nextSelectionXml;
      }

      return match;
    }) + (!replaced ? nextSelectionXml : "");
  const nextSheetViewXml = `<sheetView ${serializeAttributes(attributes)}>${nextInnerXml}</sheetView>`;
  const relativeStart = sheetViewMatch.index;
  const absoluteStart = sheetViewsMatch.index + sheetViewsMatch[0].indexOf(sheetViewMatch[0], relativeStart);

  return (
    sheetXml.slice(0, absoluteStart) +
    nextSheetViewXml +
    sheetXml.slice(absoluteStart + sheetViewMatch[0].length)
  );
}

function upsertAutoFilterInSheetXml(sheetXml: string, range: string): string {
  const normalizedRange = normalizeRangeRef(range);
  const autoFilterXml = `<autoFilter ref="${normalizedRange}"/>`;
  const autoFilterMatch = sheetXml.match(/<autoFilter\b[^>]*\/>/);

  if (autoFilterMatch && autoFilterMatch.index !== undefined) {
    return (
      sheetXml.slice(0, autoFilterMatch.index) +
      autoFilterXml +
      sheetXml.slice(autoFilterMatch.index + autoFilterMatch[0].length)
    );
  }

  const insertionIndex = findWorksheetChildInsertionIndex(sheetXml, AUTO_FILTER_FOLLOWING_TAGS);
  return sheetXml.slice(0, insertionIndex) + autoFilterXml + sheetXml.slice(insertionIndex);
}

function removeAutoFilterFromSheetXml(sheetXml: string): string {
  return sheetXml
    .replace(/<autoFilter\b[^>]*\/>/, "")
    .replace(/<sortState\b[^>]*\/>/, "");
}

function parseSheetDataValidations(sheetXml: string): DataValidation[] {
  const dataValidationsMatch = sheetXml.match(/<dataValidations\b[^>]*>([\s\S]*?)<\/dataValidations>/);
  if (!dataValidationsMatch) {
    return [];
  }

  return parseDataValidationEntries(dataValidationsMatch[1])
    .map(({ attributesSource, innerXml }) => {
      const sqref = getXmlAttr(attributesSource, "sqref");
      if (!sqref) {
        return null;
      }

      const errorTitle = getXmlAttr(attributesSource, "errorTitle");
      const error = getXmlAttr(attributesSource, "error");
      const promptTitle = getXmlAttr(attributesSource, "promptTitle");
      const prompt = getXmlAttr(attributesSource, "prompt");
      const formula1 = extractTagText(innerXml, "formula1");
      const formula2 = extractTagText(innerXml, "formula2");

      return {
        range: normalizeSqref(sqref),
        type: getXmlAttr(attributesSource, "type") ?? null,
        operator: getXmlAttr(attributesSource, "operator") ?? null,
        allowBlank: parseOptionalXmlBoolean(getXmlAttr(attributesSource, "allowBlank")),
        showInputMessage: parseOptionalXmlBoolean(getXmlAttr(attributesSource, "showInputMessage")),
        showErrorMessage: parseOptionalXmlBoolean(getXmlAttr(attributesSource, "showErrorMessage")),
        showDropDown: parseOptionalXmlBoolean(getXmlAttr(attributesSource, "showDropDown")),
        errorStyle: getXmlAttr(attributesSource, "errorStyle") ?? null,
        errorTitle: errorTitle ? decodeXmlText(errorTitle) : null,
        error: error ? decodeXmlText(error) : null,
        promptTitle: promptTitle ? decodeXmlText(promptTitle) : null,
        prompt: prompt ? decodeXmlText(prompt) : null,
        imeMode: getXmlAttr(attributesSource, "imeMode") ?? null,
        formula1: formula1 ? decodeXmlText(formula1) : null,
        formula2: formula2 ? decodeXmlText(formula2) : null,
      };
    })
    .filter((validation): validation is DataValidation => validation !== null);
}

function parseDataValidationEntries(innerXml: string): Array<{
  attributesSource: string;
  innerXml: string;
  xml: string;
}> {
  return Array.from(
    innerXml.matchAll(/<dataValidation\b([^>]*?)(?:\/>|>([\s\S]*?)<\/dataValidation>)/g),
    (match) => ({
      attributesSource: match[1],
      innerXml: match[2] ?? "",
      xml: match[0],
    }),
  );
}

function buildDataValidationXml(range: string, options: SetDataValidationOptions): string {
  const attributes: Array<[string, string]> = [["sqref", normalizeSqref(range)]];
  appendOptionalAttribute(attributes, "type", options.type);
  appendOptionalAttribute(attributes, "operator", options.operator);
  appendOptionalBooleanAttribute(attributes, "allowBlank", options.allowBlank);
  appendOptionalBooleanAttribute(attributes, "showInputMessage", options.showInputMessage);
  appendOptionalBooleanAttribute(attributes, "showErrorMessage", options.showErrorMessage);
  appendOptionalBooleanAttribute(attributes, "showDropDown", options.showDropDown);
  appendOptionalAttribute(attributes, "errorStyle", options.errorStyle);
  appendOptionalAttribute(attributes, "errorTitle", options.errorTitle);
  appendOptionalAttribute(attributes, "error", options.error);
  appendOptionalAttribute(attributes, "promptTitle", options.promptTitle);
  appendOptionalAttribute(attributes, "prompt", options.prompt);
  appendOptionalAttribute(attributes, "imeMode", options.imeMode);

  const formulas: string[] = [];
  if (options.formula1 !== undefined) {
    formulas.push(`<formula1>${escapeXmlText(options.formula1)}</formula1>`);
  }
  if (options.formula2 !== undefined) {
    formulas.push(`<formula2>${escapeXmlText(options.formula2)}</formula2>`);
  }

  return formulas.length === 0
    ? `<dataValidation ${serializeAttributes(attributes)}/>`
    : `<dataValidation ${serializeAttributes(attributes)}>${formulas.join("")}</dataValidation>`;
}

function upsertDataValidationInSheetXml(sheetXml: string, dataValidationXml: string, range: string): string {
  const normalizedRange = normalizeSqref(range);
  const dataValidationsMatch = sheetXml.match(/<dataValidations\b[^>]*>([\s\S]*?)<\/dataValidations>/);
  const dataValidations = (dataValidationsMatch
    ? parseDataValidationEntries(dataValidationsMatch[1]).map((validation) => ({
        range: normalizeSqref(getXmlAttr(validation.attributesSource, "sqref") ?? ""),
        xml: validation.xml,
      }))
    : []
  ).filter((validation) => validation.range !== normalizedRange);

  dataValidations.push({ range: normalizedRange, xml: dataValidationXml });
  const nextDataValidationsXml =
    `<dataValidations count="${dataValidations.length}">` +
    dataValidations.map((validation) => validation.xml).join("") +
    `</dataValidations>`;

  if (dataValidationsMatch && dataValidationsMatch.index !== undefined) {
    return (
      sheetXml.slice(0, dataValidationsMatch.index) +
      nextDataValidationsXml +
      sheetXml.slice(dataValidationsMatch.index + dataValidationsMatch[0].length)
    );
  }

  const insertionIndex = findWorksheetChildInsertionIndex(sheetXml, DATA_VALIDATIONS_FOLLOWING_TAGS);
  return sheetXml.slice(0, insertionIndex) + nextDataValidationsXml + sheetXml.slice(insertionIndex);
}

function removeDataValidationFromSheetXml(sheetXml: string, range: string): string {
  const normalizedRange = normalizeSqref(range);
  const dataValidationsMatch = sheetXml.match(/<dataValidations\b[^>]*>([\s\S]*?)<\/dataValidations>/);
  if (!dataValidationsMatch || dataValidationsMatch.index === undefined) {
    return sheetXml;
  }

  const keptDataValidations = parseDataValidationEntries(dataValidationsMatch[1]).filter(
    (validation) => normalizeSqref(getXmlAttr(validation.attributesSource, "sqref") ?? "") !== normalizedRange,
  );

  const nextDataValidationsXml =
    keptDataValidations.length === 0
      ? ""
      : `<dataValidations count="${keptDataValidations.length}">${keptDataValidations
          .map((validation) => validation.xml)
          .join("")}</dataValidations>`;

  return (
    sheetXml.slice(0, dataValidationsMatch.index) +
    nextDataValidationsXml +
    sheetXml.slice(dataValidationsMatch.index + dataValidationsMatch[0].length)
  );
}

function parseSheetHyperlinks(
  sheetXml: string,
  relationshipTargets: Map<string, string>,
): Hyperlink[] {
  return Array.from(sheetXml.matchAll(/<hyperlink\b([^>]*?)\/>/g), (match) => {
    const attributesSource = match[1];
    const address = getXmlAttr(attributesSource, "ref");
    const relationshipId = getXmlAttr(attributesSource, "r:id");
    const location = getXmlAttr(attributesSource, "location");
    const tooltip = getXmlAttr(attributesSource, "tooltip") ?? null;

    if (!address) {
      return null;
    }

    if (relationshipId) {
      const target = relationshipTargets.get(relationshipId);
      if (!target) {
        return null;
      }

      return {
        address: normalizeCellAddress(address),
        target,
        tooltip,
        type: "external" as const,
      };
    }

    if (!location) {
      return null;
    }

    return {
      address: normalizeCellAddress(address),
      target: location,
      tooltip,
      type: "internal" as const,
    };
  }).filter((hyperlink): hyperlink is Hyperlink => hyperlink !== null);
}

function parseHyperlinkRelationshipTargets(relationshipsXml: string): Map<string, string> {
  const targets = new Map<string, string>();

  for (const match of relationshipsXml.matchAll(/<Relationship\b([^>]*?)\/>/g)) {
    const attributesSource = match[1];
    const relationshipId = getXmlAttr(attributesSource, "Id");
    const type = getXmlAttr(attributesSource, "Type");
    const target = getXmlAttr(attributesSource, "Target");

    if (!relationshipId || !type || !target || type !== HYPERLINK_RELATIONSHIP_TYPE) {
      continue;
    }

    targets.set(relationshipId, decodeXmlText(target));
  }

  return targets;
}

function getHyperlinkRelationshipId(sheetXml: string, address: string): string | null {
  const normalizedAddress = normalizeCellAddress(address);

  for (const match of sheetXml.matchAll(/<hyperlink\b([^>]*?)\/>/g)) {
    const attributesSource = match[1];
    const ref = getXmlAttr(attributesSource, "ref");
    if (!ref || normalizeCellAddress(ref) !== normalizedAddress) {
      continue;
    }

    return getXmlAttr(attributesSource, "r:id") ?? null;
  }

  return null;
}

function buildInternalHyperlinkXml(address: string, location: string, tooltip?: string): string {
  const attributes: Array<[string, string]> = [["ref", address], ["location", location]];
  if (tooltip) {
    attributes.push(["tooltip", tooltip]);
  }

  return `<hyperlink ${serializeAttributes(attributes)}/>`;
}

function buildExternalHyperlinkXml(address: string, relationshipId: string, tooltip?: string): string {
  const attributes: Array<[string, string]> = [["ref", address], ["r:id", relationshipId]];
  if (tooltip) {
    attributes.push(["tooltip", tooltip]);
  }

  return `<hyperlink ${serializeAttributes(attributes)}/>`;
}

function upsertHyperlinkInSheetXml(sheetXml: string, hyperlinkXml: string, address: string): string {
  const normalizedAddress = normalizeCellAddress(address);
  const hyperlinksMatch = sheetXml.match(/<hyperlinks\b[^>]*>([\s\S]*?)<\/hyperlinks>/);

  const hyperlinks = (hyperlinksMatch
    ? Array.from(hyperlinksMatch[1].matchAll(/<hyperlink\b([^>]*?)\/>/g), (match) => {
        const attributesSource = match[1];
        const ref = getXmlAttr(attributesSource, "ref");
        return {
          address: ref ? normalizeCellAddress(ref) : "",
          xml: match[0],
        };
      })
    : []
  ).filter((hyperlink) => hyperlink.address !== normalizedAddress);
  hyperlinks.push({ address: normalizedAddress, xml: hyperlinkXml });
  hyperlinks.sort((left, right) => compareCellAddresses(left.address, right.address));

  const nextHyperlinksXml = `<hyperlinks>${hyperlinks.map((hyperlink) => hyperlink.xml).join("")}</hyperlinks>`;

  if (hyperlinksMatch && hyperlinksMatch.index !== undefined) {
    return (
      sheetXml.slice(0, hyperlinksMatch.index) +
      nextHyperlinksXml +
      sheetXml.slice(hyperlinksMatch.index + hyperlinksMatch[0].length)
    );
  }

  const closingTag = "</worksheet>";
  const insertionIndex = sheetXml.indexOf(closingTag);
  if (insertionIndex === -1) {
    throw new XlsxError("Worksheet is missing </worksheet>");
  }

  return sheetXml.slice(0, insertionIndex) + nextHyperlinksXml + sheetXml.slice(insertionIndex);
}

function removeHyperlinkFromSheetXml(sheetXml: string, address: string): string {
  const normalizedAddress = normalizeCellAddress(address);
  const hyperlinksMatch = sheetXml.match(/<hyperlinks\b[^>]*>([\s\S]*?)<\/hyperlinks>/);
  if (!hyperlinksMatch || hyperlinksMatch.index === undefined) {
    return sheetXml;
  }

  const keptHyperlinks = Array.from(
    hyperlinksMatch[1].matchAll(/<hyperlink\b([^>]*?)\/>/g),
    (match) => {
      const attributesSource = match[1];
      const ref = getXmlAttr(attributesSource, "ref");
      return {
        address: ref ? normalizeCellAddress(ref) : "",
        xml: match[0],
      };
    },
  ).filter((hyperlink) => hyperlink.address !== normalizedAddress);

  const nextHyperlinksXml =
    keptHyperlinks.length === 0
      ? ""
      : `<hyperlinks>${keptHyperlinks.map((hyperlink) => hyperlink.xml).join("")}</hyperlinks>`;

  return (
    sheetXml.slice(0, hyperlinksMatch.index) +
    nextHyperlinksXml +
    sheetXml.slice(hyperlinksMatch.index + hyperlinksMatch[0].length)
  );
}

function getSheetRelationshipsPath(sheetPath: string): string {
  return `${dirnamePosix(sheetPath)}/_rels/${basenamePosix(sheetPath)}.rels`;
}

function getNextTablePath(entryPaths: string[]): string {
  let nextIndex = 1;

  for (const path of entryPaths) {
    const match = path.match(/^xl\/tables\/table(\d+)\.xml$/);
    if (match) {
      nextIndex = Math.max(nextIndex, Number(match[1]) + 1);
    }
  }

  return `xl/tables/table${nextIndex}.xml`;
}

function getNextTableId(entryPaths: string[], workbook: Workbook): number {
  let nextId = 1;

  for (const path of entryPaths) {
    if (!/^xl\/tables\/table\d+\.xml$/.test(path)) {
      continue;
    }

    const tableXml = workbook.readEntryText(path);
    const idText = getXmlAttr(tableXml.match(/<table\b([^>]*?)>/)?.[1] ?? "", "id");
    if (idText) {
      nextId = Math.max(nextId, Number(idText) + 1);
    }
  }

  return nextId;
}

function getNextTableName(entryPaths: string[]): string {
  let nextIndex = 1;

  for (const path of entryPaths) {
    const match = path.match(/^xl\/tables\/table(\d+)\.xml$/);
    if (match) {
      nextIndex = Math.max(nextIndex, Number(match[1]) + 1);
    }
  }

  return `Table${nextIndex}`;
}

function assertTableName(name: string): void {
  if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(name)) {
    throw new XlsxError(`Invalid table name: ${name}`);
  }
}

function buildTableXml(
  range: string,
  id: number,
  name: string,
  headerValues: CellValue[],
): string {
  const columnNames = buildTableColumnNames(headerValues, parseRangeRef(range).endColumn - parseRangeRef(range).startColumn + 1);

  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="${id}" name="${escapeXmlText(name)}" displayName="${escapeXmlText(name)}" ref="${range}" totalsRowShown="0">` +
    `<autoFilter ref="${range}"/>` +
    `<tableColumns count="${columnNames.length}">` +
    columnNames
      .map((columnName, index) => `<tableColumn id="${index + 1}" name="${escapeXmlText(columnName)}"/>`)
      .join("") +
    `</tableColumns>` +
    `<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>` +
    `</table>`
  );
}

function buildTableColumnNames(headerValues: CellValue[], width: number): string[] {
  const names: string[] = [];
  const seen = new Map<string, number>();

  for (let index = 0; index < width; index += 1) {
    const rawValue = headerValues[index];
    const baseName =
      typeof rawValue === "string" && rawValue.trim().length > 0 ? rawValue.trim() : `Column${index + 1}`;
    const nextCount = (seen.get(baseName) ?? 0) + 1;
    seen.set(baseName, nextCount);
    names.push(nextCount === 1 ? baseName : `${baseName}_${nextCount}`);
  }

  return names;
}

function getNextRelationshipIdFromXml(relationshipsXml: string): string {
  let nextId = 1;

  for (const match of relationshipsXml.matchAll(/\bId="rId(\d+)"/g)) {
    nextId = Math.max(nextId, Number(match[1]) + 1);
  }

  return `rId${nextId}`;
}

function appendRelationship(
  relationshipsXml: string,
  relationshipId: string,
  type: string,
  target: string,
  targetMode?: string,
): string {
  const closingTag = "</Relationships>";
  const insertionIndex = relationshipsXml.indexOf(closingTag);
  if (insertionIndex === -1) {
    throw new XlsxError("Worksheet relationships file is missing </Relationships>");
  }

  const attributes: Array<[string, string]> = [
    ["Id", relationshipId],
    ["Type", type],
    ["Target", target],
  ];
  if (targetMode) {
    attributes.push(["TargetMode", targetMode]);
  }

  const relationshipXml = `<Relationship ${serializeAttributes(attributes)}/>`;
  return relationshipsXml.slice(0, insertionIndex) + relationshipXml + relationshipsXml.slice(insertionIndex);
}

function upsertRelationship(
  relationshipsXml: string,
  relationshipId: string,
  type: string,
  target: string,
  targetMode?: string,
): string {
  const nextRelationshipXml = buildRelationshipXml(relationshipId, type, target, targetMode);
  const relationshipRegex = new RegExp(`<Relationship\\b[^>]*\\bId="${escapeRegex(relationshipId)}"[^>]*/>`);

  return relationshipRegex.test(relationshipsXml)
    ? relationshipsXml.replace(relationshipRegex, nextRelationshipXml)
    : appendRelationship(relationshipsXml, relationshipId, type, target, targetMode);
}

function removeRelationshipById(relationshipsXml: string, relationshipId: string): string {
  return relationshipsXml.replace(
    new RegExp(`<Relationship\\b[^>]*\\bId="${escapeRegex(relationshipId)}"[^>]*/>`),
    "",
  );
}

function buildRelationshipXml(
  relationshipId: string,
  type: string,
  target: string,
  targetMode?: string,
): string {
  const attributes: Array<[string, string]> = [
    ["Id", relationshipId],
    ["Type", type],
    ["Target", target],
  ];
  if (targetMode) {
    attributes.push(["TargetMode", targetMode]);
  }

  return `<Relationship ${serializeAttributes(attributes)}/>`;
}

function makeRelativeSheetRelationshipTarget(sheetPath: string, targetPath: string): string {
  const fromParts = dirnamePosix(sheetPath).split("/").filter((part) => part.length > 0);
  const toParts = targetPath.split("/").filter((part) => part.length > 0);
  let commonLength = 0;

  while (
    commonLength < fromParts.length &&
    commonLength < toParts.length &&
    fromParts[commonLength] === toParts[commonLength]
  ) {
    commonLength += 1;
  }

  const upward = fromParts.slice(commonLength).map(() => "..");
  const downward = toParts.slice(commonLength);
  return [...upward, ...downward].join("/");
}

function appendTablePart(sheetXml: string, relationshipId: string): string {
  const tablePartsMatch = sheetXml.match(/<tableParts\b[^>]*>([\s\S]*?)<\/tableParts>/);
  if (tablePartsMatch && tablePartsMatch.index !== undefined) {
    const tableParts = Array.from(
      tablePartsMatch[1].matchAll(/<tablePart\b([^>]*?)\/>/g),
      (match) => `<tablePart${match[1] ? ` ${match[1].trim()}` : ""}/>`,
    );
    tableParts.push(`<tablePart r:id="${relationshipId}"/>`);
    const nextTablePartsXml = `<tableParts count="${tableParts.length}">${tableParts.join("")}</tableParts>`;
    return (
      sheetXml.slice(0, tablePartsMatch.index) +
      nextTablePartsXml +
      sheetXml.slice(tablePartsMatch.index + tablePartsMatch[0].length)
    );
  }

  const closingTag = "</worksheet>";
  const insertionIndex = sheetXml.indexOf(closingTag);
  if (insertionIndex === -1) {
    throw new XlsxError("Worksheet is missing </worksheet>");
  }

  return (
    sheetXml.slice(0, insertionIndex) +
    `<tableParts count="1"><tablePart r:id="${relationshipId}"/></tableParts>` +
    sheetXml.slice(insertionIndex)
  );
}

function findWorksheetChildInsertionIndex(sheetXml: string, followingTagNames: string[]): number {
  let insertionIndex = -1;

  for (const tagName of followingTagNames) {
    const match = sheetXml.match(new RegExp(`<${escapeRegex(tagName)}\\b`));
    if (!match || match.index === undefined) {
      continue;
    }

    if (insertionIndex === -1 || match.index < insertionIndex) {
      insertionIndex = match.index;
    }
  }

  if (insertionIndex !== -1) {
    return insertionIndex;
  }

  const closingTag = "</worksheet>";
  const closingTagIndex = sheetXml.indexOf(closingTag);
  if (closingTagIndex === -1) {
    throw new XlsxError("Worksheet is missing </worksheet>");
  }

  return closingTagIndex;
}

function appendOptionalAttribute(attributes: Array<[string, string]>, name: string, value: string | undefined): void {
  if (value !== undefined) {
    attributes.push([name, value]);
  }
}

function appendOptionalBooleanAttribute(attributes: Array<[string, string]>, name: string, value: boolean | undefined): void {
  if (value !== undefined) {
    attributes.push([name, value ? "1" : "0"]);
  }
}

function buildFreezePaneXml(columnCount: number, rowCount: number): string {
  const attributes: Array<[string, string]> = [["state", "frozen"]];
  if (columnCount > 0) {
    attributes.push(["xSplit", String(columnCount)]);
  }
  if (rowCount > 0) {
    attributes.push(["ySplit", String(rowCount)]);
  }
  attributes.push(["topLeftCell", makeCellAddress(rowCount + 1, columnCount + 1)]);
  const activePane = getFreezePaneActivePane(columnCount, rowCount);
  if (activePane) {
    attributes.push(["activePane", activePane]);
  }

  return `<pane ${serializeAttributes(attributes)}/>`;
}

function buildFreezePaneSelectionsXml(columnCount: number, rowCount: number): string {
  const topLeftCell = makeCellAddress(rowCount + 1, columnCount + 1);

  if (columnCount > 0 && rowCount > 0) {
    return [
      buildSelectionXml([["pane", "topRight"]]),
      buildSelectionXml([["pane", "bottomLeft"]]),
      buildSelectionXml([["pane", "bottomRight"], ["activeCell", topLeftCell], ["sqref", topLeftCell]]),
    ].join("");
  }

  if (columnCount > 0) {
    return buildSelectionXml([["pane", "topRight"], ["activeCell", topLeftCell], ["sqref", topLeftCell]]);
  }

  return buildSelectionXml([["pane", "bottomLeft"], ["activeCell", topLeftCell], ["sqref", topLeftCell]]);
}

function parseSheetSelectionEntries(sheetXml: string): SheetSelection[] {
  return Array.from(sheetXml.matchAll(/<selection\b([^>]*?)\/>/g), (match) => {
    const attributesSource = match[1];
    const activeCell = getXmlAttr(attributesSource, "activeCell");
    const sqref = getXmlAttr(attributesSource, "sqref");

    return {
      activeCell: activeCell ? normalizeCellAddress(activeCell) : null,
      range: sqref ? normalizeSqref(sqref) : null,
      pane: normalizePaneName(getXmlAttr(attributesSource, "pane")),
    };
  });
}

function buildSelectionXml(attributes: Array<[string, string]>): string {
  return attributes.length === 0 ? "<selection/>" : `<selection ${serializeAttributes(attributes)}/>`;
}

function getFreezePaneActivePane(
  columnCount: number,
  rowCount: number,
): "bottomLeft" | "topRight" | "bottomRight" | null {
  if (columnCount > 0 && rowCount > 0) {
    return "bottomRight";
  }

  if (columnCount > 0) {
    return "topRight";
  }

  if (rowCount > 0) {
    return "bottomLeft";
  }

  return null;
}

function normalizePaneName(
  value: string | undefined,
): "bottomLeft" | "topRight" | "bottomRight" | null {
  if (value === "bottomLeft" || value === "topRight" || value === "bottomRight") {
    return value;
  }

  return null;
}

function parseOptionalXmlBoolean(value: string | undefined): boolean | null {
  if (value === undefined) {
    return null;
  }

  return value === "1" || value.toLowerCase() === "true";
}

function addContentTypeOverride(contentTypesXml: string, partPath: string, contentType: string): string {
  if (new RegExp(`PartName="/${escapeRegex(partPath)}"`).test(contentTypesXml)) {
    return contentTypesXml;
  }

  const closingTag = "</Types>";
  const insertionIndex = contentTypesXml.indexOf(closingTag);
  if (insertionIndex === -1) {
    throw new XlsxError("Content types file is missing </Types>");
  }

  return (
    contentTypesXml.slice(0, insertionIndex) +
    `<Override PartName="/${escapeXmlText(partPath)}" ContentType="${escapeXmlText(contentType)}"/>` +
    contentTypesXml.slice(insertionIndex)
  );
}

function removeContentTypeOverride(contentTypesXml: string, partPath: string): string {
  return contentTypesXml.replace(
    new RegExp(`<Override\\b[^>]*\\bPartName="/${escapeRegex(partPath)}"[^>]*/>`),
    "",
  );
}

function assertFreezeSplit(columnCount: number, rowCount: number): void {
  if (!Number.isInteger(columnCount) || columnCount < 0) {
    throw new XlsxError(`Invalid freeze column count: ${columnCount}`);
  }

  if (!Number.isInteger(rowCount) || rowCount < 0) {
    throw new XlsxError(`Invalid freeze row count: ${rowCount}`);
  }

  if (columnCount === 0 && rowCount === 0) {
    throw new XlsxError("Freeze pane requires at least one frozen row or column");
  }
}

function assertStyleId(styleId: number | null): void {
  if (styleId !== null && (!Number.isInteger(styleId) || styleId < 0)) {
    throw new XlsxError(`Invalid style id: ${styleId}`);
  }
}

function resolveCloneStylePatch(
  addressOrRowNumber: string | number,
  columnOrPatch: number | string | CellStylePatch | undefined,
  patch: CellStylePatch | undefined,
): CellStylePatch {
  return typeof addressOrRowNumber === "number" ? (patch ?? {}) : ((columnOrPatch as CellStylePatch | undefined) ?? {});
}

function resolveSetStyleId(
  addressOrRowNumber: string | number,
  columnOrStyleId: number | string | null,
  styleId?: number | null,
): number | null {
  const nextStyleId =
    typeof addressOrRowNumber === "number" ? (styleId ?? null) : (columnOrStyleId as number | null);
  assertStyleId(nextStyleId);
  return nextStyleId;
}

function resolveCopyStyleArguments(
  sourceAddressOrRowNumber: string | number,
  sourceColumnOrTargetAddress: number | string,
  targetRowNumber?: number,
  targetColumn?: number | string,
): { sourceAddress: string; targetAddress: string } {
  if (typeof sourceAddressOrRowNumber === "string") {
    if (typeof sourceColumnOrTargetAddress !== "string") {
      throw new XlsxError("Missing target address for copyStyle");
    }

    return {
      sourceAddress: resolveCellAddress(sourceAddressOrRowNumber),
      targetAddress: resolveCellAddress(sourceColumnOrTargetAddress),
    };
  }

  if (targetRowNumber === undefined || targetColumn === undefined) {
    throw new XlsxError("Missing target row or column for copyStyle");
  }

  return {
    sourceAddress: resolveCellAddress(sourceAddressOrRowNumber, sourceColumnOrTargetAddress),
    targetAddress: resolveCellAddress(targetRowNumber, targetColumn),
  };
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

function compareCellAddresses(left: string, right: string): number {
  const leftCell = splitCellAddress(left);
  const rightCell = splitCellAddress(right);
  return leftCell.rowNumber - rightCell.rowNumber || leftCell.columnNumber - rightCell.columnNumber;
}

function updateDimensionRef(sheetIndex: SheetIndex): string {
  const usedRange = formatUsedRangeBounds(sheetIndex.usedBounds);
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

function formatUsedRangeBounds(bounds: UsedRangeBounds | null): string | null {
  return bounds ? formatRangeRef(bounds.minRow, bounds.minColumn, bounds.maxRow, bounds.maxColumn) : null;
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
const TABLE_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
const HYPERLINK_RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const ROW_CLOSE_TAG = "</row>";
const CELL_CLOSE_TAG = "</c>";
const AUTO_FILTER_FOLLOWING_TAGS = [
  "sortState",
  "mergeCells",
  "phoneticPr",
  "conditionalFormatting",
  "dataValidations",
  "hyperlinks",
  "printOptions",
  "pageMargins",
  "pageSetup",
  "headerFooter",
  "rowBreaks",
  "colBreaks",
  "customProperties",
  "cellWatches",
  "ignoredErrors",
  "smartTags",
  "drawing",
  "legacyDrawing",
  "legacyDrawingHF",
  "picture",
  "oleObjects",
  "controls",
  "webPublishItems",
  "tableParts",
  "extLst",
];
const SHEET_VIEWS_FOLLOWING_TAGS = [
  "sheetFormatPr",
  "cols",
  "sheetData",
  "autoFilter",
  "sortState",
  "mergeCells",
  "phoneticPr",
  "conditionalFormatting",
  "dataValidations",
  "hyperlinks",
  "printOptions",
  "pageMargins",
  "pageSetup",
  "headerFooter",
  "rowBreaks",
  "colBreaks",
  "customProperties",
  "cellWatches",
  "ignoredErrors",
  "smartTags",
  "drawing",
  "legacyDrawing",
  "legacyDrawingHF",
  "picture",
  "oleObjects",
  "controls",
  "webPublishItems",
  "tableParts",
  "extLst",
];
const COLS_FOLLOWING_TAGS = [
  "sheetData",
  "autoFilter",
  "sortState",
  "mergeCells",
  "phoneticPr",
  "conditionalFormatting",
  "dataValidations",
  "hyperlinks",
  "printOptions",
  "pageMargins",
  "pageSetup",
  "headerFooter",
  "rowBreaks",
  "colBreaks",
  "customProperties",
  "cellWatches",
  "ignoredErrors",
  "smartTags",
  "drawing",
  "legacyDrawing",
  "legacyDrawingHF",
  "picture",
  "oleObjects",
  "controls",
  "webPublishItems",
  "tableParts",
  "extLst",
];
const DATA_VALIDATIONS_FOLLOWING_TAGS = [
  "hyperlinks",
  "printOptions",
  "pageMargins",
  "pageSetup",
  "headerFooter",
  "rowBreaks",
  "colBreaks",
  "customProperties",
  "cellWatches",
  "ignoredErrors",
  "smartTags",
  "drawing",
  "legacyDrawing",
  "legacyDrawingHF",
  "picture",
  "oleObjects",
  "controls",
  "webPublishItems",
  "tableParts",
  "extLst",
];
const EMPTY_RELATIONSHIPS_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
