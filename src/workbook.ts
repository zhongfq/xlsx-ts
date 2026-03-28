import type { ArchiveEntry, DefinedName, SetDefinedNameOptions, SheetVisibility } from "./types.js";
import { XlsxError } from "./errors.js";
import {
  Sheet,
  deleteFormulaReferences,
  deleteSheetFormulaReferences,
  renameSheetFormulaReferences,
  shiftFormulaReferences,
} from "./sheet.js";
import { CliZipAdapter } from "./zip-cli.js";
import { basenamePosix, dirnamePosix, resolvePosix } from "./utils/path.js";
import {
  escapeXmlText,
  decodeXmlText,
  escapeRegex,
  extractAllTagTexts,
  getXmlAttr,
  parseAttributes,
  serializeAttributes,
} from "./utils/xml.js";

const XML_DECODER = new TextDecoder();
const XML_ENCODER = new TextEncoder();
const WORKSHEET_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const WORKSHEET_RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";

interface WorkbookContext {
  workbookDir: string;
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

  getSheetVisibility(sheetName: string): SheetVisibility {
    const context = this.getWorkbookContext();
    const sheet = context.sheets.find((candidate) => candidate.name === sheetName);
    if (!sheet) {
      throw new XlsxError(`Sheet not found: ${sheetName}`);
    }

    return parseSheetVisibility(this.readEntryText(context.workbookPath), sheet.relationshipId);
  }

  setSheetVisibility(sheetName: string, visibility: SheetVisibility): void {
    assertSheetVisibility(visibility);

    const context = this.getWorkbookContext();
    const sheet = context.sheets.find((candidate) => candidate.name === sheetName);
    if (!sheet) {
      throw new XlsxError(`Sheet not found: ${sheetName}`);
    }

    const workbookPath = context.workbookPath;
    const workbookXml = this.readEntryText(workbookPath);
    const currentVisibility = parseSheetVisibility(workbookXml, sheet.relationshipId);

    if (currentVisibility === visibility) {
      return;
    }

    const visibleSheetCount = context.sheets.filter(
      (candidate) => parseSheetVisibility(workbookXml, candidate.relationshipId) === "visible",
    ).length;
    if (currentVisibility === "visible" && visibility !== "visible" && visibleSheetCount === 1) {
      throw new XlsxError("Workbook must contain at least one visible sheet");
    }

    this.writeEntryText(
      workbookPath,
      updateSheetVisibilityInWorkbookXml(workbookXml, sheet.relationshipId, visibility),
    );
  }

  getDefinedNames(): DefinedName[] {
    const context = this.getWorkbookContext();
    return parseDefinedNames(this.readEntryText(context.workbookPath), context.sheets);
  }

  getDefinedName(name: string, scope?: string): string | null {
    const definedName = this.getDefinedNames().find(
      (candidate) => candidate.name === name && candidate.scope === (scope ?? null),
    );
    return definedName?.value ?? null;
  }

  setDefinedName(name: string, value: string, options: SetDefinedNameOptions = {}): void {
    assertDefinedName(name);

    const context = this.getWorkbookContext();
    const scope = options.scope ?? null;
    const localSheetId = scope === null ? null : context.sheets.findIndex((sheet) => sheet.name === scope);

    if (scope !== null && localSheetId === -1) {
      throw new XlsxError(`Sheet not found: ${scope}`);
    }

    const workbookPath = context.workbookPath;
    const workbookXml = this.readEntryText(workbookPath);
    const nextDefinedNameXml = buildDefinedNameXml(name, value, localSheetId);
    let replaced = false;
    const nextWorkbookXml = workbookXml.replace(
      /<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g,
      (match, attributesSource) => {
        const candidateName = getXmlAttr(attributesSource, "name");
        const candidateLocalSheetId = getXmlAttr(attributesSource, "localSheetId");
        const candidateScope = candidateLocalSheetId === undefined ? null : Number(candidateLocalSheetId);

        if (candidateName !== name || candidateScope !== localSheetId) {
          return match;
        }

        replaced = true;
        return nextDefinedNameXml;
      },
    );

    if (replaced) {
      this.writeEntryText(workbookPath, nextWorkbookXml);
      return;
    }

    this.writeEntryText(workbookPath, insertDefinedNameIntoWorkbookXml(workbookXml, nextDefinedNameXml));
  }

  deleteDefinedName(name: string, scope?: string): void {
    const context = this.getWorkbookContext();
    const targetScope = scope ?? null;
    const localSheetId = targetScope === null ? null : context.sheets.findIndex((sheet) => sheet.name === targetScope);

    if (targetScope !== null && localSheetId === -1) {
      throw new XlsxError(`Sheet not found: ${targetScope}`);
    }

    const workbookPath = context.workbookPath;
    const workbookXml = this.readEntryText(workbookPath);
    const nextWorkbookXml = removeDefinedNameFromWorkbookXml(workbookXml, name, localSheetId);

    if (nextWorkbookXml !== workbookXml) {
      this.writeEntryText(workbookPath, nextWorkbookXml);
    }
  }

  renameSheet(currentSheetName: string, nextSheetName: string): Sheet {
    assertSheetName(nextSheetName);

    const context = this.getWorkbookContext();
    const renamedSheet = context.sheets.find((sheet) => sheet.name === currentSheetName);
    if (!renamedSheet) {
      throw new XlsxError(`Sheet not found: ${currentSheetName}`);
    }

    if (currentSheetName === nextSheetName) {
      return renamedSheet;
    }

    if (context.sheets.some((sheet) => sheet.name === nextSheetName)) {
      throw new XlsxError(`Sheet already exists: ${nextSheetName}`);
    }

    for (const sheet of context.sheets) {
      this.rewriteSheetFormulaTexts(sheet.path, (formula) =>
        renameSheetFormulaReferences(formula, currentSheetName, nextSheetName),
      );
      this.rewriteSheetHyperlinkLocations(sheet.path, currentSheetName, nextSheetName);
    }

    const workbookXml = this.readEntryText(context.workbookPath);
    this.writeEntryText(
      context.workbookPath,
      renameSheetInWorkbookXml(workbookXml, renamedSheet.relationshipId, currentSheetName, nextSheetName),
    );
    this.rewriteAppSheetNames(
      context.sheets.map((sheet) => (sheet.name === currentSheetName ? nextSheetName : sheet.name)),
    );
    renamedSheet.name = nextSheetName;
    return renamedSheet;
  }

  addSheet(sheetName: string): Sheet {
    assertSheetName(sheetName);

    const context = this.getWorkbookContext();
    if (context.sheets.some((sheet) => sheet.name === sheetName)) {
      throw new XlsxError(`Sheet already exists: ${sheetName}`);
    }

    const workbookXml = this.readEntryText(context.workbookPath);
    const workbookRelsXml = this.readEntryText(context.workbookRelsPath);
    const nextSheetId = getNextSheetId(workbookXml);
    const nextRelationshipId = getNextRelationshipId(workbookRelsXml);
    const nextSheetPath = getNextWorksheetPath(context.workbookDir, this.entryOrder);
    const relationshipTarget = toRelationshipTarget(context.workbookDir, nextSheetPath);
    const contentTypesXml = this.readEntryText("[Content_Types].xml");

    this.writeEntryText(nextSheetPath, buildEmptyWorksheetXml());
    this.writeEntryText(
      context.workbookPath,
      insertBeforeClosingTag(
        workbookXml,
        "sheets",
        `<sheet name="${escapeXmlText(sheetName)}" sheetId="${nextSheetId}" r:id="${nextRelationshipId}"/>`,
      ),
    );
    this.writeEntryText(
      context.workbookRelsPath,
      insertBeforeClosingTag(
        workbookRelsXml,
        "Relationships",
        `<Relationship Id="${nextRelationshipId}" Type="${WORKSHEET_RELATIONSHIP_TYPE}" Target="${escapeXmlText(relationshipTarget)}"/>`,
      ),
    );
    this.writeEntryText(
      "[Content_Types].xml",
      insertBeforeClosingTag(
        contentTypesXml,
        "Types",
        `<Override PartName="/${escapeXmlText(nextSheetPath)}" ContentType="${WORKSHEET_CONTENT_TYPE}"/>`,
      ),
    );
    this.rewriteAppSheetNames([...context.sheets.map((sheet) => sheet.name), sheetName]);

    return this.getSheet(sheetName);
  }

  deleteSheet(sheetName: string): void {
    const context = this.getWorkbookContext();
    if (context.sheets.length === 1) {
      throw new XlsxError("Cannot delete the last sheet");
    }

    const deletedSheetIndex = context.sheets.findIndex((sheet) => sheet.name === sheetName);
    if (deletedSheetIndex === -1) {
      throw new XlsxError(`Sheet not found: ${sheetName}`);
    }

    const deletedSheet = context.sheets[deletedSheetIndex];
    if (!deletedSheet) {
      throw new XlsxError(`Sheet not found: ${sheetName}`);
    }

    const workbookXml = this.readEntryText(context.workbookPath);
    const workbookRelsXml = this.readEntryText(context.workbookRelsPath);
    const contentTypesXml = this.readEntryText("[Content_Types].xml");

    for (const sheet of context.sheets) {
      if (sheet.path === deletedSheet.path) {
        continue;
      }

      this.rewriteSheetFormulaTexts(sheet.path, (formula) =>
        deleteSheetFormulaReferences(formula, sheetName),
      );
    }

    this.writeEntryText(
      context.workbookPath,
      removeSheetFromWorkbookXml(workbookXml, deletedSheet.relationshipId, sheetName, deletedSheetIndex),
    );
    this.writeEntryText(
      context.workbookRelsPath,
      removeRelationshipById(workbookRelsXml, deletedSheet.relationshipId),
    );
    this.writeEntryText(
      "[Content_Types].xml",
      removeContentTypeOverride(contentTypesXml, deletedSheet.path),
    );
    this.rewriteAppSheetNames(
      context.sheets.filter((sheet) => sheet.name !== sheetName).map((sheet) => sheet.name),
    );
    this.removeEntry(deletedSheet.path);

    const sheetRelsPath = `${dirnamePosix(deletedSheet.path)}/_rels/${basenamePosix(deletedSheet.path)}.rels`;
    if (this.entries.has(sheetRelsPath)) {
      this.removeEntry(sheetRelsPath);
    }
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
      workbookDir,
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

  rewriteDefinedNamesForSheetStructure(
    sheetName: string,
    targetColumnNumber: number,
    columnCount: number,
    targetRowNumber: number,
    rowCount: number,
    mode: "shift" | "delete",
  ): void {
    const context = this.getWorkbookContext();
    const localSheetIndex = context.sheets.findIndex((sheet) => sheet.name === sheetName);
    if (localSheetIndex === -1) {
      return;
    }

    const workbookXml = this.readEntryText(context.workbookPath);
    let changed = false;
    const nextWorkbookXml = workbookXml.replace(
      /<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g,
      (match, attributesSource, nameSource) => {
        const localSheetIdText = getXmlAttr(attributesSource, "localSheetId");
        const includeUnqualifiedReferences =
          localSheetIdText !== undefined && Number(localSheetIdText) === localSheetIndex;
        const nameText = decodeXmlText(nameSource);
        const nextNameText =
          mode === "shift"
            ? shiftFormulaReferences(
                nameText,
                sheetName,
                targetColumnNumber,
                columnCount,
                targetRowNumber,
                rowCount,
                includeUnqualifiedReferences,
              )
            : deleteFormulaReferences(
                nameText,
                sheetName,
                targetColumnNumber,
                columnCount,
                targetRowNumber,
                rowCount,
                includeUnqualifiedReferences,
              );

        if (nextNameText === nameText) {
          return match;
        }

        changed = true;
        return `<definedName${attributesSource}>${escapeXmlText(nextNameText)}</definedName>`;
      },
    );

    if (changed) {
      this.writeEntryText(context.workbookPath, nextWorkbookXml);
    }
  }

  private rewriteSheetFormulaTexts(
    path: string,
    transformFormula: (formula: string) => string,
  ): void {
    const sheetXml = this.readEntryText(path);
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
      this.writeEntryText(path, nextSheetXml);
    }
  }

  private rewriteSheetHyperlinkLocations(
    path: string,
    currentSheetName: string,
    nextSheetName: string,
  ): void {
    const sheetXml = this.readEntryText(path);
    let changed = false;
    const nextSheetXml = sheetXml.replace(/<hyperlink\b([^>]*?)\/>/g, (match, attributesSource) => {
      const attributes = parseAttributes(attributesSource);
      const locationIndex = attributes.findIndex(([name]) => name === "location");

      if (locationIndex === -1) {
        return match;
      }

      const location = attributes[locationIndex]?.[1] ?? "";
      const nextLocation = renameHyperlinkLocation(location, currentSheetName, nextSheetName);
      if (nextLocation === location) {
        return match;
      }

      changed = true;
      const nextAttributes = [...attributes];
      nextAttributes[locationIndex] = ["location", nextLocation];
      const serializedAttributes = serializeAttributes(nextAttributes);
      return `<hyperlink${serializedAttributes ? ` ${serializedAttributes}` : ""}/>`;
    });

    if (changed) {
      this.writeEntryText(path, nextSheetXml);
    }
  }

  private rewriteAppSheetNames(sheetNames: string[]): void {
    const appPath = "docProps/app.xml";
    if (!this.entries.has(appPath)) {
      return;
    }

    const appXml = this.readEntryText(appPath);
    const nextAppXml = updateAppSheetNames(appXml, sheetNames);

    if (nextAppXml !== appXml) {
      this.writeEntryText(appPath, nextAppXml);
    }
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

    if (
      this.workbookContext &&
      (this.workbookContext.workbookPath === path || this.workbookContext.workbookRelsPath === path)
    ) {
      this.workbookContext = undefined;
    }

    this.entries.set(path, XML_ENCODER.encode(text));
  }

  removeEntry(path: string): void {
    if (!this.entries.delete(path)) {
      return;
    }

    const entryIndex = this.entryOrder.indexOf(path);
    if (entryIndex !== -1) {
      this.entryOrder.splice(entryIndex, 1);
    }

    if (
      this.workbookContext &&
      (this.workbookContext.sharedStringsPath === path ||
        this.workbookContext.workbookPath === path ||
        this.workbookContext.workbookRelsPath === path)
    ) {
      this.sharedStringsCache = undefined;
      this.workbookContext = undefined;
    }
  }
}

function assertSheetName(sheetName: string): void {
  if (sheetName.length === 0 || sheetName.length > 31 || /[\\/*?:[\]]/.test(sheetName)) {
    throw new XlsxError(`Invalid sheet name: ${sheetName}`);
  }
}

function assertDefinedName(name: string): void {
  if (!/^[A-Za-z_\\][A-Za-z0-9_.\\]*$/.test(name)) {
    throw new XlsxError(`Invalid defined name: ${name}`);
  }
}

function assertSheetVisibility(visibility: string): asserts visibility is SheetVisibility {
  if (visibility !== "visible" && visibility !== "hidden" && visibility !== "veryHidden") {
    throw new XlsxError(`Invalid sheet visibility: ${visibility}`);
  }
}

function buildEmptyWorksheetXml(): string {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData></sheetData></worksheet>`
  );
}

function getNextSheetId(workbookXml: string): number {
  let nextSheetId = 1;

  for (const match of workbookXml.matchAll(/<sheet\b[^>]*\bsheetId="(\d+)"/g)) {
    nextSheetId = Math.max(nextSheetId, Number(match[1]) + 1);
  }

  return nextSheetId;
}

function getNextRelationshipId(relationshipsXml: string): string {
  let nextId = 1;

  for (const match of relationshipsXml.matchAll(/\bId="rId(\d+)"/g)) {
    nextId = Math.max(nextId, Number(match[1]) + 1);
  }

  return `rId${nextId}`;
}

function getNextWorksheetPath(workbookDir: string, entryOrder: string[]): string {
  let nextIndex = 1;
  const prefix = workbookDir ? `${workbookDir}/worksheets/` : "worksheets/";

  for (const path of entryOrder) {
    const match = path.match(new RegExp(`^${escapeRegex(prefix)}sheet(\\d+)\\.xml$`));
    if (match) {
      nextIndex = Math.max(nextIndex, Number(match[1]) + 1);
    }
  }

  return `${prefix}sheet${nextIndex}.xml`;
}

function toRelationshipTarget(workbookDir: string, path: string): string {
  return workbookDir && path.startsWith(`${workbookDir}/`) ? path.slice(workbookDir.length + 1) : path;
}

function insertBeforeClosingTag(xml: string, tagName: string, snippet: string): string {
  const closingTag = `</${tagName}>`;
  const insertionIndex = xml.indexOf(closingTag);
  if (insertionIndex === -1) {
    throw new XlsxError(`Missing closing tag: ${closingTag}`);
  }

  return xml.slice(0, insertionIndex) + snippet + xml.slice(insertionIndex);
}

function renameSheetInWorkbookXml(
  workbookXml: string,
  relationshipId: string,
  currentSheetName: string,
  nextSheetName: string,
): string {
  return workbookXml.replace(
    /<sheet\b([^>]*?)\/>/g,
    (match, attributesSource) => {
      const attributes = parseAttributes(attributesSource);
      const relationshipIndex = attributes.findIndex(([name]) => name === "r:id");

      if (relationshipIndex === -1 || attributes[relationshipIndex]?.[1] !== relationshipId) {
        return match;
      }

      const nextAttributes = attributes.map(([name, value]) => {
        if (name === "name") {
          return [name, nextSheetName] as [string, string];
        }

        return [name, value] as [string, string];
      });
      const serializedAttributes = serializeAttributes(nextAttributes);
      return `<sheet ${serializedAttributes}/>`;
    },
  ).replace(
    /<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g,
    (match, attributesSource, nameSource) => {
      const nameText = decodeXmlText(nameSource);
      const nextNameText = renameSheetFormulaReferences(nameText, currentSheetName, nextSheetName);
      return nextNameText === nameText
        ? match
        : `<definedName${attributesSource}>${escapeXmlText(nextNameText)}</definedName>`;
    },
  );
}

function parseSheetVisibility(workbookXml: string, relationshipId: string): SheetVisibility {
  for (const match of workbookXml.matchAll(/<sheet\b([^>]*?)\/>/g)) {
    const attributesSource = match[1];
    if (getXmlAttr(attributesSource, "r:id") !== relationshipId) {
      continue;
    }

    const state = getXmlAttr(attributesSource, "state");
    if (state === "hidden" || state === "veryHidden") {
      return state;
    }

    return "visible";
  }

  throw new XlsxError(`Sheet relationship not found: ${relationshipId}`);
}

function updateSheetVisibilityInWorkbookXml(
  workbookXml: string,
  relationshipId: string,
  visibility: SheetVisibility,
): string {
  let changed = false;

  const nextWorkbookXml = workbookXml.replace(
    /<sheet\b([^>]*?)\/>/g,
    (match, attributesSource) => {
      const attributes = parseAttributes(attributesSource);
      const relationshipIndex = attributes.findIndex(([name]) => name === "r:id");

      if (relationshipIndex === -1 || attributes[relationshipIndex]?.[1] !== relationshipId) {
        return match;
      }

      changed = true;
      const withoutState = attributes.filter(([name]) => name !== "state");
      const nextAttributes =
        visibility === "visible"
          ? withoutState
          : [...withoutState, ["state", visibility] as [string, string]];
      return `<sheet ${serializeAttributes(nextAttributes)}/>`;
    },
  );

  if (!changed) {
    throw new XlsxError(`Sheet relationship not found: ${relationshipId}`);
  }

  return nextWorkbookXml;
}

function parseDefinedNames(workbookXml: string, sheets: Sheet[]): DefinedName[] {
  return Array.from(
    workbookXml.matchAll(/<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g),
    (match) => {
      const attributesSource = match[1];
      const localSheetIdText = getXmlAttr(attributesSource, "localSheetId");
      const localSheetId = localSheetIdText === undefined ? null : Number(localSheetIdText);
      return {
        hidden: getXmlAttr(attributesSource, "hidden") === "1",
        name: getXmlAttr(attributesSource, "name") ?? "",
        scope: localSheetId === null ? null : (sheets[localSheetId]?.name ?? null),
        value: decodeXmlText(match[2]),
      };
    },
  ).filter((definedName) => definedName.name.length > 0);
}

function buildDefinedNameXml(name: string, value: string, localSheetId: number | null): string {
  const attributes: Array<[string, string]> = [["name", name]];
  if (localSheetId !== null) {
    attributes.push(["localSheetId", String(localSheetId)]);
  }

  return `<definedName ${serializeAttributes(attributes)}>${escapeXmlText(value)}</definedName>`;
}

function insertDefinedNameIntoWorkbookXml(workbookXml: string, definedNameXml: string): string {
  const definedNamesMatch = workbookXml.match(/<definedNames\b[^>]*>([\s\S]*?)<\/definedNames>/);
  if (definedNamesMatch && definedNamesMatch.index !== undefined) {
    const insertionIndex = definedNamesMatch.index + definedNamesMatch[0].length - "</definedNames>".length;
    return workbookXml.slice(0, insertionIndex) + definedNameXml + workbookXml.slice(insertionIndex);
  }

  return insertBeforeClosingTag(workbookXml, "workbook", `<definedNames>${definedNameXml}</definedNames>`);
}

function removeDefinedNameFromWorkbookXml(
  workbookXml: string,
  name: string,
  localSheetId: number | null,
): string {
  const definedNamesMatch = workbookXml.match(/<definedNames\b[^>]*>([\s\S]*?)<\/definedNames>/);
  if (!definedNamesMatch || definedNamesMatch.index === undefined) {
    return workbookXml;
  }

  const nextInnerXml = definedNamesMatch[1].replace(
    /<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g,
    (match, attributesSource) => {
      const candidateName = getXmlAttr(attributesSource, "name");
      const candidateLocalSheetIdText = getXmlAttr(attributesSource, "localSheetId");
      const candidateLocalSheetId = candidateLocalSheetIdText === undefined ? null : Number(candidateLocalSheetIdText);
      return candidateName === name && candidateLocalSheetId === localSheetId ? "" : match;
    },
  );

  const nextDefinedNamesXml = `<definedNames>${nextInnerXml}</definedNames>`;
  const nextWorkbookXml =
    workbookXml.slice(0, definedNamesMatch.index) +
    nextDefinedNamesXml +
    workbookXml.slice(definedNamesMatch.index + definedNamesMatch[0].length);

  return /<definedName\b/.test(nextInnerXml)
    ? nextWorkbookXml
    : workbookXml.slice(0, definedNamesMatch.index) +
        workbookXml.slice(definedNamesMatch.index + definedNamesMatch[0].length);
}

function removeSheetFromWorkbookXml(
  workbookXml: string,
  relationshipId: string,
  deletedSheetName: string,
  deletedSheetIndex: number,
): string {
  const withoutSheet = workbookXml.replace(
    new RegExp(`<sheet\\b[^>]*\\br:id="${escapeRegex(relationshipId)}"[^>]*/>`),
    "",
  );

  return withoutSheet.replace(
    /<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g,
    (match, attributesSource, nameSource) => {
      const attributes = parseAttributes(attributesSource);
      const localSheetIdIndex = attributes.findIndex(([name]) => name === "localSheetId");
      const localSheetIdText = localSheetIdIndex === -1 ? undefined : attributes[localSheetIdIndex]?.[1];

      if (localSheetIdText !== undefined) {
        const localSheetId = Number(localSheetIdText);
        if (localSheetId === deletedSheetIndex) {
          return "";
        }

        if (localSheetId > deletedSheetIndex) {
          attributes[localSheetIdIndex] = ["localSheetId", String(localSheetId - 1)];
        }
      }

      const nameText = decodeXmlText(nameSource);
      const nextNameText = deleteSheetFormulaReferences(nameText, deletedSheetName);
      const serializedAttributes = serializeAttributes(attributes);
      return `<definedName${serializedAttributes ? ` ${serializedAttributes}` : ""}>${escapeXmlText(nextNameText)}</definedName>`;
    },
  );
}

function renameHyperlinkLocation(
  location: string,
  currentSheetName: string,
  nextSheetName: string,
): string {
  const hashPrefix = location.startsWith("#") ? "#" : "";
  const target = hashPrefix ? location.slice(1) : location;
  const bangIndex = target.indexOf("!");

  if (bangIndex === -1) {
    return location;
  }

  const sheetToken = target.slice(0, bangIndex);
  const normalizedSheetName =
    sheetToken.startsWith("'") && sheetToken.endsWith("'")
      ? sheetToken.slice(1, -1).replaceAll("''", "'")
      : sheetToken;

  if (normalizedSheetName !== currentSheetName) {
    return location;
  }

  return `${hashPrefix}${formatSheetNameForReference(nextSheetName)}${target.slice(bangIndex)}`;
}

function removeRelationshipById(relationshipsXml: string, relationshipId: string): string {
  return relationshipsXml.replace(
    new RegExp(`<Relationship\\b[^>]*\\bId="${escapeRegex(relationshipId)}"[^>]*/>`),
    "",
  );
}

function removeContentTypeOverride(contentTypesXml: string, partPath: string): string {
  return contentTypesXml.replace(
    new RegExp(`<Override\\b[^>]*\\bPartName="/${escapeRegex(partPath)}"[^>]*/>`),
    "",
  );
}

function updateAppSheetNames(appXml: string, sheetNames: string[]): string {
  const headingPairsMatch = appXml.match(
    /<HeadingPairs>([\s\S]*?)<\/HeadingPairs>/,
  );
  const titlesOfPartsMatch = appXml.match(
    /<TitlesOfParts>([\s\S]*?)<\/TitlesOfParts>/,
  );

  if (!headingPairsMatch && !titlesOfPartsMatch) {
    return appXml;
  }

  let nextAppXml = appXml;

  if (headingPairsMatch) {
    const nextHeadingPairs =
      `<HeadingPairs><vt:vector size="2" baseType="variant">` +
      `<vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant>` +
      `<vt:variant><vt:i4>${sheetNames.length}</vt:i4></vt:variant>` +
      `</vt:vector></HeadingPairs>`;
    nextAppXml = nextAppXml.replace(/<HeadingPairs>[\s\S]*?<\/HeadingPairs>/, nextHeadingPairs);
  }

  if (titlesOfPartsMatch) {
    const nextTitlesOfParts =
      `<TitlesOfParts><vt:vector size="${sheetNames.length}" baseType="lpstr">` +
      sheetNames.map((sheetName) => `<vt:lpstr>${escapeXmlText(sheetName)}</vt:lpstr>`).join("") +
      `</vt:vector></TitlesOfParts>`;
    nextAppXml = nextAppXml.replace(/<TitlesOfParts>[\s\S]*?<\/TitlesOfParts>/, nextTitlesOfParts);
  }

  return nextAppXml;
}

function formatSheetNameForReference(sheetName: string): string {
  if (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheetName)) {
    return sheetName;
  }

  return `'${sheetName.replaceAll("'", "''")}'`;
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
