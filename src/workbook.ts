import type { ArchiveEntry } from "./types.js";
import { XlsxError } from "./errors.js";
import { Sheet } from "./sheet.js";
import { CliZipAdapter } from "./zip-cli.js";
import { basenamePosix, dirnamePosix, resolvePosix } from "./utils/path.js";
import {
  decodeXmlText,
  extractAllTagTexts,
  getXmlAttr,
} from "./utils/xml.js";

const XML_DECODER = new TextDecoder();
const XML_ENCODER = new TextEncoder();

interface WorkbookContext {
  workbookPath: string;
  workbookRelsPath: string;
  sharedStringsPath?: string;
  sheets: Sheet[];
}

export class Workbook {
  private readonly adapter: CliZipAdapter;
  private readonly entryOrder: string[];
  private readonly entries: Map<string, Uint8Array>;
  private workbookContext?: WorkbookContext;
  private sharedStringsCache?: string[];

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

    const rootRels = this.readEntryText("_rels/.rels");
    const workbookTarget = findRelationshipTarget(rootRels, /\/officeDocument$/) ?? "xl/workbook.xml";
    const workbookPath = workbookTarget.replace(/^\/+/, "");
    const workbookDir = dirnamePosix(workbookPath);
    const workbookRelsPath = `${workbookDir}/_rels/${basenamePosix(workbookPath)}.rels`;
    const workbookXml = this.readEntryText(workbookPath);
    const workbookRelsXml = this.readEntryText(workbookRelsPath);
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

  readSharedStrings(): string[] {
    if (this.sharedStringsCache) {
      return [...this.sharedStringsCache];
    }

    const sharedStringsPath = this.getWorkbookContext().sharedStringsPath;
    if (!sharedStringsPath || !this.entries.has(sharedStringsPath)) {
      return [];
    }

    const xml = this.readEntryText(sharedStringsPath);
    this.sharedStringsCache = Array.from(xml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g), (match) =>
      extractAllTagTexts(match[1], "t").map(decodeXmlText).join(""),
    );
    return [...this.sharedStringsCache];
  }

  readEntryText(path: string): string {
    const entry = this.entries.get(path);
    if (!entry) {
      throw new XlsxError(`Entry not found: ${path}`);
    }

    return XML_DECODER.decode(entry);
  }

  writeEntryText(path: string, text: string): void {
    if (!this.entries.has(path)) {
      this.entryOrder.push(path);
    }

    if (this.workbookContext?.sharedStringsPath === path) {
      this.sharedStringsCache = undefined;
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
