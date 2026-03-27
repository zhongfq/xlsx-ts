import type { ArchiveEntry, CellValue } from "./types.js";
import { XlsxError } from "./errors.js";
import { Sheet } from "./sheet.js";
import { CliZipAdapter } from "./zip-cli.js";
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

const XML_DECODER = new TextDecoder();
const XML_ENCODER = new TextEncoder();

interface WorkbookContext {
  workbookPath: string;
  workbookRelsPath: string;
  sharedStringsPath?: string;
  sheets: Sheet[];
}

interface LocatedCell {
  start: number;
  end: number;
  attributesSource: string;
  innerXml: string;
}

interface LocatedRow {
  start: number;
  end: number;
  attributesSource: string;
  innerXml: string;
  selfClosing: boolean;
}

export class Workbook {
  private readonly adapter: CliZipAdapter;
  private readonly entryOrder: string[];
  private readonly entries: Map<string, Uint8Array>;
  private workbookContext?: WorkbookContext;

  constructor(entries: Iterable<ArchiveEntry>, adapter = new CliZipAdapter()) {
    this.adapter = adapter;
    this.entries = new Map();
    this.entryOrder = [];

    for (const entry of entries) {
      this.entryOrder.push(entry.path);
      this.entries.set(entry.path, new Uint8Array(entry.data));
    }
  }

  static async open(filePath: string): Promise<Workbook> {
    const adapter = new CliZipAdapter();
    const entries = await adapter.readArchive(filePath);
    return new Workbook(entries, adapter);
  }

  static fromEntries(entries: Iterable<ArchiveEntry>): Workbook {
    return new Workbook(entries);
  }

  listEntries(): string[] {
    return [...this.entryOrder];
  }

  getSheets(): Sheet[] {
    return [...this.getWorkbookContext().sheets];
  }

  getSheet(sheetName: string): Sheet {
    const sheet = this.getWorkbookContext().sheets.find((candidate) => candidate.name === sheetName);

    if (!sheet) {
      throw new XlsxError(`Sheet not found: ${sheetName}`);
    }

    return sheet;
  }

  readCell(sheet: Sheet, address: string): CellValue {
    const sheetXml = this.requireText(sheet.path);
    const cell = locateCell(sheetXml, address);

    if (!cell) {
      return null;
    }

    const type = getXmlAttr(cell.attributesSource, "t");

    if (type === "inlineStr" || type === "str") {
      const text = extractAllTagTexts(cell.innerXml, "t").map(decodeXmlText).join("");
      return text;
    }

    if (type === "s") {
      const indexText = extractTagText(cell.innerXml, "v");
      if (!indexText) {
        return null;
      }

      const sharedStrings = this.readSharedStrings();
      const value = sharedStrings[Number(indexText)];
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

  writeCell(sheet: Sheet, address: string, value: CellValue): void {
    assertCellAddress(address);

    const sheetXml = this.requireText(sheet.path);
    const existingCell = locateCell(sheetXml, address);
    const nextCellXml = buildCellXml(address, value, existingCell?.attributesSource);

    let nextSheetXml: string;

    if (existingCell) {
      nextSheetXml =
        sheetXml.slice(0, existingCell.start) + nextCellXml + sheetXml.slice(existingCell.end);
    } else {
      nextSheetXml = insertCell(sheetXml, address, nextCellXml);
    }

    this.writeText(sheet.path, nextSheetXml);
  }

  async save(filePath: string): Promise<void> {
    await this.adapter.writeArchive(filePath, this.toEntries());
  }

  toEntries(): ArchiveEntry[] {
    return this.entryOrder.map((path) => {
      const data = this.entries.get(path);
      if (!data) {
        throw new XlsxError(`Entry missing from map: ${path}`);
      }

      return { path, data: new Uint8Array(data) };
    });
  }

  private getWorkbookContext(): WorkbookContext {
    if (this.workbookContext) {
      return this.workbookContext;
    }

    const rootRels = this.requireText("_rels/.rels");
    const workbookTarget = findRelationshipTarget(rootRels, /\/officeDocument$/) ?? "xl/workbook.xml";
    const workbookPath = workbookTarget.replace(/^\/+/, "");
    const workbookDir = dirnamePosix(workbookPath);
    const workbookRelsPath = `${workbookDir}/_rels/${basenamePosix(workbookPath)}.rels`;
    const workbookXml = this.requireText(workbookPath);
    const workbookRelsXml = this.requireText(workbookRelsPath);
    const relationships = parseRelationships(workbookRelsXml, workbookDir);
    const sheets = parseSheets(this, workbookXml, relationships);
    const sharedStringsPath = findRelationshipTarget(workbookRelsXml, /\/sharedStrings$/, workbookDir);

    this.workbookContext = {
      workbookPath,
      workbookRelsPath,
      sharedStringsPath,
      sheets,
    };

    return this.workbookContext;
  }

  private readSharedStrings(): string[] {
    const sharedStringsPath = this.getWorkbookContext().sharedStringsPath;
    if (!sharedStringsPath || !this.entries.has(sharedStringsPath)) {
      return [];
    }

    const xml = this.requireText(sharedStringsPath);
    return Array.from(xml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g), (match) =>
      extractAllTagTexts(match[1], "t").map(decodeXmlText).join(""),
    );
  }

  private requireText(path: string): string {
    const entry = this.entries.get(path);
    if (!entry) {
      throw new XlsxError(`Entry not found: ${path}`);
    }

    return XML_DECODER.decode(entry);
  }

  private writeText(path: string, text: string): void {
    if (!this.entries.has(path)) {
      this.entryOrder.push(path);
    }

    this.entries.set(path, XML_ENCODER.encode(text));
  }
}

function parseRelationships(xml: string, baseDir: string): Map<string, string> {
  const relationships = new Map<string, string>();

  for (const match of xml.matchAll(/<Relationship\b([^>]*?)\/>/g)) {
    const attributesSource = match[1];
    const id = getXmlAttr(attributesSource, "Id");
    const target = getXmlAttr(attributesSource, "Target");

    if (!id || !target) {
      continue;
    }

    relationships.set(id, resolvePosix(baseDir, target.replace(/^\/+/, "")));
  }

  return relationships;
}

function parseSheets(
  workbook: Workbook,
  workbookXml: string,
  relationships: Map<string, string>,
): Sheet[] {
  const sheets: Sheet[] = [];

  for (const match of workbookXml.matchAll(/<sheet\b([^>]*?)(?:\/>|>[\s\S]*?<\/sheet>)/g)) {
    const attributesSource = match[1];
    const name = getXmlAttr(attributesSource, "name");
    const relationshipId = getXmlAttr(attributesSource, "r:id");

    if (!name || !relationshipId) {
      continue;
    }

    const path = relationships.get(relationshipId);
    if (!path) {
      continue;
    }

    sheets.push(
      new Sheet(workbook, {
        name,
        path,
        relationshipId,
      }),
    );
  }

  return sheets;
}

function findRelationshipTarget(
  xml: string,
  typePattern: RegExp,
  baseDir = "",
): string | undefined {
  for (const match of xml.matchAll(/<Relationship\b([^>]*?)\/>/g)) {
    const attributesSource = match[1];
    const type = getXmlAttr(attributesSource, "Type");
    const target = getXmlAttr(attributesSource, "Target");

    if (!type || !target || !typePattern.test(type)) {
      continue;
    }

    return baseDir ? resolvePosix(baseDir, target.replace(/^\/+/, "")) : target;
  }

  return undefined;
}

function locateCell(sheetXml: string, address: string): LocatedCell | undefined {
  const regex = new RegExp(
    `<c\\b([^>]*\\br="${escapeRegex(address)}"[^>]*)\\s*(?:>([\\s\\S]*?)<\\/c>|\\/>)`,
    "g",
  );
  const match = regex.exec(sheetXml);

  if (!match || match.index === undefined) {
    return undefined;
  }

  return {
    start: match.index,
    end: match.index + match[0].length,
    attributesSource: match[1].trim(),
    innerXml: match[2] ?? "",
  };
}

function locateRow(sheetXml: string, rowNumber: number): LocatedRow | undefined {
  const regex = new RegExp(
    `<row\\b([^>]*\\br="${rowNumber}"[^>]*)\\s*(?:>([\\s\\S]*?)<\\/row>|\\/>)`,
    "g",
  );
  const match = regex.exec(sheetXml);

  if (!match || match.index === undefined) {
    return undefined;
  }

  return {
    start: match.index,
    end: match.index + match[0].length,
    attributesSource: match[1].trim(),
    innerXml: match[2] ?? "",
    selfClosing: !match[0].includes("</row>"),
  };
}

function buildCellXml(address: string, value: CellValue, existingAttributesSource?: string): string {
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

function insertCell(sheetXml: string, address: string, cellXml: string): string {
  const { rowNumber, columnNumber } = splitCellAddress(address);
  const row = locateRow(sheetXml, rowNumber);

  if (row) {
    if (row.selfClosing) {
      const nextRowXml = `<row ${row.attributesSource}>${cellXml}</row>`;
      return sheetXml.slice(0, row.start) + nextRowXml + sheetXml.slice(row.end);
    }

    const insertionIndex = findCellInsertionIndex(row.innerXml, columnNumber);
    const nextRowXml =
      `<row ${row.attributesSource}>` +
      row.innerXml.slice(0, insertionIndex) +
      cellXml +
      row.innerXml.slice(insertionIndex) +
      `</row>`;

    return sheetXml.slice(0, row.start) + nextRowXml + sheetXml.slice(row.end);
  }

  const sheetDataMatch = sheetXml.match(/<sheetData\b[^>]*>([\s\S]*?)<\/sheetData>/);
  if (!sheetDataMatch || sheetDataMatch.index === undefined) {
    throw new XlsxError("Worksheet is missing <sheetData>");
  }

  const sheetDataInner = sheetDataMatch[1];
  const relativeRowIndex = findRowInsertionIndex(sheetDataInner, rowNumber);
  const rowXml = `<row r="${rowNumber}">${cellXml}</row>`;
  const start = sheetDataMatch.index + sheetDataMatch[0].indexOf(sheetDataInner);

  return sheetXml.slice(0, start + relativeRowIndex) + rowXml + sheetXml.slice(start + relativeRowIndex);
}

function findCellInsertionIndex(rowInnerXml: string, columnNumber: number): number {
  const regex = /<c\b([^>]*\br="([A-Z]+)\d+"[^>]*)\s*(?:>[\s\S]*?<\/c>|\/>)/g;

  for (const match of rowInnerXml.matchAll(regex)) {
    const cellReference = match[2];
    if (columnLabelToNumber(cellReference) > columnNumber) {
      return match.index ?? rowInnerXml.length;
    }
  }

  return rowInnerXml.length;
}

function findRowInsertionIndex(sheetDataInnerXml: string, rowNumber: number): number {
  const regex = /<row\b([^>]*\br="(\d+)"[^>]*)\s*(?:>[\s\S]*?<\/row>|\/>)/g;

  for (const match of sheetDataInnerXml.matchAll(regex)) {
    const candidateRow = Number(match[2]);
    if (candidateRow > rowNumber) {
      return match.index ?? sheetDataInnerXml.length;
    }
  }

  return sheetDataInnerXml.length;
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
