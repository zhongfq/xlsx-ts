import type {
  ArchiveEntry,
  CellBorderColor,
  CellBorderColorPatch,
  CellBorderDefinition,
  CellBorderPatch,
  CellBorderSideDefinition,
  CellBorderSidePatch,
  CellFillColor,
  CellFillColorPatch,
  CellFillDefinition,
  CellFillPatch,
  CellFontColor,
  CellFontColorPatch,
  CellFontDefinition,
  CellFontPatch,
  CellNumberFormatDefinition,
  CellStyleAlignment,
  CellStyleAlignmentPatch,
  CellStyleDefinition,
  CellStylePatch,
  DefinedName,
  SetDefinedNameOptions,
  SheetVisibility,
} from "./types.js";
import { XlsxError } from "./errors.js";
import {
  Sheet,
  deleteFormulaReferences,
  deleteSheetFormulaReferences,
  renameSheetFormulaReferences,
  shiftFormulaReferences,
} from "./sheet.js";
import { parseSharedStrings } from "./shared-strings.js";
import { Zip } from "./zip.js";
import type { WorkbookContext } from "./workbook-context.js";
import { resolveWorkbookContext } from "./workbook-context.js";
import { basenamePosix, dirnamePosix } from "./utils/path.js";
import { findFirstXmlTag, findXmlTags, getTagAttr } from "./utils/xml-read.js";
import {
  escapeXmlText,
  decodeXmlText,
  escapeRegex,
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

interface ParsedCellStyle {
  alignmentAttributes: Array<[string, string]> | null;
  attributes: Array<[string, string]>;
  definition: CellStyleDefinition;
  extraChildrenXml: string;
}

interface ParsedFont {
  definition: CellFontDefinition;
  extraChildrenXml: string;
}

interface ParsedFill {
  definition: CellFillDefinition;
  extraChildrenXml: string;
}

interface ParsedBorder {
  definition: CellBorderDefinition;
  extraChildrenXml: string;
}

interface StylesCache {
  borders: ParsedBorder[];
  cellXfs: ParsedCellStyle[];
  fills: ParsedFill[];
  fonts: ParsedFont[];
  numberFormats: Map<number, string>;
  path: string;
  xml: string;
}

const BUILTIN_NUMBER_FORMATS = new Map<number, string>([
  [0, "General"],
  [1, "0"],
  [2, "0.00"],
  [3, "#,##0"],
  [4, "#,##0.00"],
  [9, "0%"],
  [10, "0.00%"],
  [11, "0.00E+00"],
  [12, "# ?/?"],
  [13, "# ??/??"],
  [14, "mm-dd-yy"],
  [15, "d-mmm-yy"],
  [16, "d-mmm"],
  [17, "mmm-yy"],
  [18, "h:mm AM/PM"],
  [19, "h:mm:ss AM/PM"],
  [20, "h:mm"],
  [21, "h:mm:ss"],
  [22, "m/d/yy h:mm"],
  [37, "#,##0 ;(#,##0)"],
  [38, "#,##0 ;[Red](#,##0)"],
  [39, "#,##0.00;(#,##0.00)"],
  [40, "#,##0.00;[Red](#,##0.00)"],
  [45, "mm:ss"],
  [46, "[h]:mm:ss"],
  [47, "mmss.0"],
  [48, "##0.0E+0"],
  [49, "@"],
]);

export class Workbook {
  private readonly adapter: Zip;
  private readonly entryOrder: string[];
  private readonly entries: Map<string, Uint8Array>;
  private workbookContext?: WorkbookContext;
  private sharedStringsCache?: string[];
  private stylesCache?: StylesCache | null;

  constructor(entries: Iterable<ArchiveEntry>, adapter = new Zip()) {
    this.adapter = adapter;
    this.entries = new Map();
    this.entryOrder = [];

    for (const entry of entries) {
      this.entryOrder.push(entry.path);
      this.entries.set(entry.path, new Uint8Array(entry.data));
    }
  }

  static async open(filePath: string): Promise<Workbook> {
    const adapter = new Zip();
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

  getActiveSheet(): Sheet {
    const context = this.getWorkbookContext();
    const activeSheetIndex = parseActiveSheetIndex(this.readEntryText(context.workbookPath), context.sheets.length);
    return context.sheets[activeSheetIndex] ?? context.sheets[0]!;
  }

  getStyle(styleId: number): CellStyleDefinition | null {
    assertStyleId(styleId);
    return cloneCellStyleDefinition(this.getStylesCache()?.cellXfs[styleId]?.definition ?? null);
  }

  getNumberFormat(numFmtId: number): CellNumberFormatDefinition | null {
    assertStyleId(numFmtId);
    const styles = this.getStylesCache();
    const customCode = styles?.numberFormats.get(numFmtId);
    if (customCode !== undefined) {
      return {
        builtin: false,
        code: customCode,
        numFmtId,
      };
    }

    const builtinCode = BUILTIN_NUMBER_FORMATS.get(numFmtId);
    if (builtinCode !== undefined) {
      return {
        builtin: true,
        code: builtinCode,
        numFmtId,
      };
    }

    return null;
  }

  getFont(fontId: number): CellFontDefinition | null {
    assertStyleId(fontId);
    return cloneCellFontDefinition(this.getStylesCache()?.fonts[fontId]?.definition ?? null);
  }

  getFill(fillId: number): CellFillDefinition | null {
    assertStyleId(fillId);
    return cloneCellFillDefinition(this.getStylesCache()?.fills[fillId]?.definition ?? null);
  }

  getBorder(borderId: number): CellBorderDefinition | null {
    assertStyleId(borderId);
    return cloneCellBorderDefinition(this.getStylesCache()?.borders[borderId]?.definition ?? null);
  }

  updateNumberFormat(numFmtId: number, formatCode: string): void {
    assertStyleId(numFmtId);
    assertFormatCode(formatCode);

    const styles = this.getStylesCache();
    if (!styles) {
      throw new XlsxError("Workbook styles.xml not found");
    }

    if (!styles.numberFormats.has(numFmtId)) {
      if (BUILTIN_NUMBER_FORMATS.has(numFmtId)) {
        throw new XlsxError(`Cannot update builtin number format: ${numFmtId}`);
      }

      throw new XlsxError(`Number format not found: ${numFmtId}`);
    }

    this.writeEntryText(styles.path, upsertNumberFormatInStylesXml(styles.xml, numFmtId, formatCode));
  }

  cloneNumberFormat(numFmtId: number, formatCode?: string): number {
    assertStyleId(numFmtId);
    if (formatCode !== undefined) {
      assertFormatCode(formatCode);
    }

    const styles = this.getStylesCache();
    if (!styles) {
      throw new XlsxError("Workbook styles.xml not found");
    }

    const sourceCode = formatCode ?? styles.numberFormats.get(numFmtId) ?? BUILTIN_NUMBER_FORMATS.get(numFmtId);
    if (sourceCode === undefined) {
      throw new XlsxError(`Number format not found: ${numFmtId}`);
    }

    const nextNumFmtId = getNextCustomNumberFormatId(styles.numberFormats);
    this.writeEntryText(styles.path, upsertNumberFormatInStylesXml(styles.xml, nextNumFmtId, sourceCode));
    return nextNumFmtId;
  }

  ensureNumberFormat(formatCode: string): number {
    assertFormatCode(formatCode);

    for (const [numFmtId, builtinCode] of BUILTIN_NUMBER_FORMATS) {
      if (builtinCode === formatCode) {
        return numFmtId;
      }
    }

    const styles = this.getStylesCache();
    if (!styles) {
      throw new XlsxError("Workbook styles.xml not found");
    }

    for (const [numFmtId, existingCode] of styles.numberFormats) {
      if (existingCode === formatCode) {
        return numFmtId;
      }
    }

    const nextNumFmtId = getNextCustomNumberFormatId(styles.numberFormats);
    this.writeEntryText(styles.path, upsertNumberFormatInStylesXml(styles.xml, nextNumFmtId, formatCode));
    return nextNumFmtId;
  }

  updateFont(fontId: number, patch: CellFontPatch): void {
    assertStyleId(fontId);
    assertCellFontPatch(patch);

    const styles = this.getStylesCache();
    if (!styles) {
      throw new XlsxError("Workbook styles.xml not found");
    }

    const sourceFont = styles.fonts[fontId];
    if (!sourceFont) {
      throw new XlsxError(`Font not found: ${fontId}`);
    }

    this.writeEntryText(
      styles.path,
      replaceFontInStylesXml(styles.xml, fontId, buildPatchedFontXml(sourceFont, patch)),
    );
  }

  cloneFont(fontId: number, patch: CellFontPatch = {}): number {
    assertStyleId(fontId);
    assertCellFontPatch(patch);

    const styles = this.getStylesCache();
    if (!styles) {
      throw new XlsxError("Workbook styles.xml not found");
    }

    const sourceFont = styles.fonts[fontId];
    if (!sourceFont) {
      throw new XlsxError(`Font not found: ${fontId}`);
    }

    const nextFontId = styles.fonts.length;
    this.writeEntryText(
      styles.path,
      appendFontToStylesXml(styles.xml, buildPatchedFontXml(sourceFont, patch)),
    );
    return nextFontId;
  }

  updateFill(fillId: number, patch: CellFillPatch): void {
    assertStyleId(fillId);
    assertCellFillPatch(patch);

    const styles = this.getStylesCache();
    if (!styles) {
      throw new XlsxError("Workbook styles.xml not found");
    }

    const sourceFill = styles.fills[fillId];
    if (!sourceFill) {
      throw new XlsxError(`Fill not found: ${fillId}`);
    }

    this.writeEntryText(
      styles.path,
      replaceFillInStylesXml(styles.xml, fillId, buildPatchedFillXml(sourceFill, patch)),
    );
  }

  cloneFill(fillId: number, patch: CellFillPatch = {}): number {
    assertStyleId(fillId);
    assertCellFillPatch(patch);

    const styles = this.getStylesCache();
    if (!styles) {
      throw new XlsxError("Workbook styles.xml not found");
    }

    const sourceFill = styles.fills[fillId];
    if (!sourceFill) {
      throw new XlsxError(`Fill not found: ${fillId}`);
    }

    const nextFillId = styles.fills.length;
    this.writeEntryText(
      styles.path,
      appendFillToStylesXml(styles.xml, buildPatchedFillXml(sourceFill, patch)),
    );
    return nextFillId;
  }

  updateBorder(borderId: number, patch: CellBorderPatch): void {
    assertStyleId(borderId);
    assertCellBorderPatch(patch);

    const styles = this.getStylesCache();
    if (!styles) {
      throw new XlsxError("Workbook styles.xml not found");
    }

    const sourceBorder = styles.borders[borderId];
    if (!sourceBorder) {
      throw new XlsxError(`Border not found: ${borderId}`);
    }

    this.writeEntryText(
      styles.path,
      replaceBorderInStylesXml(styles.xml, borderId, buildPatchedBorderXml(sourceBorder, patch)),
    );
  }

  cloneBorder(borderId: number, patch: CellBorderPatch = {}): number {
    assertStyleId(borderId);
    assertCellBorderPatch(patch);

    const styles = this.getStylesCache();
    if (!styles) {
      throw new XlsxError("Workbook styles.xml not found");
    }

    const sourceBorder = styles.borders[borderId];
    if (!sourceBorder) {
      throw new XlsxError(`Border not found: ${borderId}`);
    }

    const nextBorderId = styles.borders.length;
    this.writeEntryText(
      styles.path,
      appendBorderToStylesXml(styles.xml, buildPatchedBorderXml(sourceBorder, patch)),
    );
    return nextBorderId;
  }

  updateStyle(styleId: number, patch: CellStylePatch): void {
    assertStyleId(styleId);
    assertCellStylePatch(patch);

    const styles = this.getStylesCache();
    if (!styles) {
      throw new XlsxError("Workbook styles.xml not found");
    }

    const sourceStyle = styles.cellXfs[styleId];
    if (!sourceStyle) {
      throw new XlsxError(`Style not found: ${styleId}`);
    }

    this.writeEntryText(
      styles.path,
      replaceCellXfInStylesXml(styles.xml, styleId, buildPatchedCellXfXml(sourceStyle, patch)),
    );
  }

  cloneStyle(styleId: number, patch: CellStylePatch = {}): number {
    assertStyleId(styleId);
    assertCellStylePatch(patch);

    const styles = this.getStylesCache();
    if (!styles) {
      throw new XlsxError("Workbook styles.xml not found");
    }

    const sourceStyle = styles.cellXfs[styleId];
    if (!sourceStyle) {
      throw new XlsxError(`Style not found: ${styleId}`);
    }

    const nextStyleId = styles.cellXfs.length;
    this.writeEntryText(
      styles.path,
      appendCellXfToStylesXml(styles.xml, buildPatchedCellXfXml(sourceStyle, patch)),
    );
    return nextStyleId;
  }

  setActiveSheet(sheetName: string): Sheet {
    const context = this.getWorkbookContext();
    const targetIndex = context.sheets.findIndex((sheet) => sheet.name === sheetName);
    if (targetIndex === -1) {
      throw new XlsxError(`Sheet not found: ${sheetName}`);
    }

    if (this.getSheetVisibility(sheetName) !== "visible") {
      throw new XlsxError(`Cannot activate hidden sheet: ${sheetName}`);
    }

    const workbookPath = context.workbookPath;
    this.writeEntryText(
      workbookPath,
      updateActiveSheetInWorkbookXml(this.readEntryText(workbookPath), targetIndex),
    );
    return this.getSheet(sheetName);
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

  moveSheet(sheetName: string, targetIndex: number): Sheet {
    const context = this.getWorkbookContext();
    const sourceIndex = context.sheets.findIndex((sheet) => sheet.name === sheetName);
    if (sourceIndex === -1) {
      throw new XlsxError(`Sheet not found: ${sheetName}`);
    }

    assertSheetIndex(targetIndex, context.sheets.length);
    if (sourceIndex === targetIndex) {
      return context.sheets[sourceIndex]!;
    }

    const nextSheets = [...context.sheets];
    const [movedSheet] = nextSheets.splice(sourceIndex, 1);
    nextSheets.splice(targetIndex, 0, movedSheet!);

    const workbookPath = context.workbookPath;
    const workbookXml = this.readEntryText(workbookPath);
    this.writeEntryText(
      workbookPath,
      reorderWorkbookXmlSheets(workbookXml, context.sheets, nextSheets),
    );
    this.rewriteAppSheetNames(nextSheets.map((sheet) => sheet.name));
    return this.getSheet(sheetName);
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

    this.workbookContext = resolveWorkbookContext(this, (path) => this.readEntryText(path));

    return this.workbookContext;
  }

  readSharedStrings(): string[] {
    return [...this.getSharedStringsCache()];
  }

  getSharedString(index: number): string | null {
    return this.getSharedStringsCache()[index] ?? null;
  }

  private getSharedStringsCache(): string[] {
    if (this.sharedStringsCache) {
      return this.sharedStringsCache;
    }

    const sharedStringsPath = this.getWorkbookContext().sharedStringsPath;
    if (!sharedStringsPath || !this.entries.has(sharedStringsPath)) {
      this.sharedStringsCache = [];
      return this.sharedStringsCache;
    }

    this.sharedStringsCache = parseSharedStrings(this.readEntryText(sharedStringsPath));
    return this.sharedStringsCache;
  }

  private getStylesCache(): StylesCache | null {
    if (this.stylesCache !== undefined) {
      return this.stylesCache;
    }

    const stylesPath = this.getWorkbookContext().stylesPath;
    if (!stylesPath || !this.entries.has(stylesPath)) {
      this.stylesCache = null;
      return this.stylesCache;
    }

    this.stylesCache = parseStylesXml(stylesPath, this.readEntryText(stylesPath));
    return this.stylesCache;
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

    if (this.workbookContext?.stylesPath === path) {
      this.stylesCache = undefined;
    }

    if (
      this.workbookContext &&
      (this.workbookContext.workbookPath === path || this.workbookContext.workbookRelsPath === path)
    ) {
      this.stylesCache = undefined;
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
        this.workbookContext.stylesPath === path ||
        this.workbookContext.workbookPath === path ||
        this.workbookContext.workbookRelsPath === path)
    ) {
      this.sharedStringsCache = undefined;
      this.stylesCache = undefined;
      this.workbookContext = undefined;
    }
  }
}

function parseStylesXml(path: string, xml: string): StylesCache {
  const bordersMatch = xml.match(/<borders\b([^>]*)>([\s\S]*?)<\/borders>/);
  const fontsMatch = xml.match(/<fonts\b([^>]*)>([\s\S]*?)<\/fonts>/);
  const fillsMatch = xml.match(/<fills\b([^>]*)>([\s\S]*?)<\/fills>/);
  const cellXfsMatch = xml.match(/<cellXfs\b([^>]*)>([\s\S]*?)<\/cellXfs>/);
  if (!bordersMatch) {
    throw new XlsxError("styles.xml is missing <borders>");
  }
  if (!fontsMatch) {
    throw new XlsxError("styles.xml is missing <fonts>");
  }
  if (!fillsMatch) {
    throw new XlsxError("styles.xml is missing <fills>");
  }
  if (!cellXfsMatch) {
    throw new XlsxError("styles.xml is missing <cellXfs>");
  }

  return {
    path,
    xml,
    borders: Array.from(
      bordersMatch[2].matchAll(/<border\b[^>]*?>([\s\S]*?)<\/border>|<border\b[^>]*?\/>/g),
      (match) => parseBorder(match[0]),
    ),
    fills: Array.from(fillsMatch[2].matchAll(/<fill\b[^>]*?>([\s\S]*?)<\/fill>|<fill\b[^>]*?\/>/g), (match) =>
      parseFill(match[0]),
    ),
    fonts: Array.from(fontsMatch[2].matchAll(/<font\b[^>]*?>([\s\S]*?)<\/font>|<font\b[^>]*?\/>/g), (match) =>
      parseFont(match[0]),
    ),
    numberFormats: parseNumberFormats(xml),
    cellXfs: Array.from(cellXfsMatch[2].matchAll(/<xf\b([^>]*?)(?:\/>|>([\s\S]*?)<\/xf>)/g), (match) =>
      parseCellStyle(match[1], match[2] ?? ""),
    ),
  };
}

function parseNumberFormats(stylesXml: string): Map<number, string> {
  const numberFormats = new Map<number, string>();

  for (const match of stylesXml.matchAll(/<numFmt\b([^>]*)\/>/g)) {
    const numFmtIdText = getXmlAttr(match[1], "numFmtId");
    const formatCode = getXmlAttr(match[1], "formatCode");
    if (numFmtIdText === undefined || formatCode === undefined) {
      continue;
    }

    numberFormats.set(Number(numFmtIdText), decodeXmlText(formatCode));
  }

  return numberFormats;
}

function parseBorder(borderXml: string): ParsedBorder {
  if (/^<border\b[^>]*\/>$/.test(borderXml)) {
    return {
      definition: buildEmptyBorderDefinition(),
      extraChildrenXml: "",
    };
  }

  const borderMatch = borderXml.match(/<border\b([^>]*?)>([\s\S]*?)<\/border>/);
  const borderAttributes = parseAttributes(borderMatch?.[1] ?? "");
  let remainingXml = borderMatch?.[2] ?? "";

  const [leftXml, remainingAfterLeft] = takeFirstTag(remainingXml, /<left\b([^>]*?)(?:\/>|>[\s\S]*?<\/left>)/);
  remainingXml = remainingAfterLeft;
  const [rightXml, remainingAfterRight] = takeFirstTag(remainingXml, /<right\b([^>]*?)(?:\/>|>[\s\S]*?<\/right>)/);
  remainingXml = remainingAfterRight;
  const [topXml, remainingAfterTop] = takeFirstTag(remainingXml, /<top\b([^>]*?)(?:\/>|>[\s\S]*?<\/top>)/);
  remainingXml = remainingAfterTop;
  const [bottomXml, remainingAfterBottom] = takeFirstTag(remainingXml, /<bottom\b([^>]*?)(?:\/>|>[\s\S]*?<\/bottom>)/);
  remainingXml = remainingAfterBottom;
  const [diagonalXml, remainingAfterDiagonal] = takeFirstTag(
    remainingXml,
    /<diagonal\b([^>]*?)(?:\/>|>[\s\S]*?<\/diagonal>)/,
  );
  remainingXml = remainingAfterDiagonal;
  const [verticalXml, remainingAfterVertical] = takeFirstTag(
    remainingXml,
    /<vertical\b([^>]*?)(?:\/>|>[\s\S]*?<\/vertical>)/,
  );
  remainingXml = remainingAfterVertical;
  const [horizontalXml, remainingAfterHorizontal] = takeFirstTag(
    remainingXml,
    /<horizontal\b([^>]*?)(?:\/>|>[\s\S]*?<\/horizontal>)/,
  );
  remainingXml = remainingAfterHorizontal;

  return {
    definition: {
      left: parseBorderSideDefinition(leftXml),
      right: parseBorderSideDefinition(rightXml),
      top: parseBorderSideDefinition(topXml),
      bottom: parseBorderSideDefinition(bottomXml),
      diagonal: parseBorderSideDefinition(diagonalXml),
      vertical: parseBorderSideDefinition(verticalXml),
      horizontal: parseBorderSideDefinition(horizontalXml),
      diagonalUp: parseOptionalBooleanAttribute(borderAttributes, "diagonalUp"),
      diagonalDown: parseOptionalBooleanAttribute(borderAttributes, "diagonalDown"),
      outline: parseOptionalBooleanAttribute(borderAttributes, "outline"),
    },
    extraChildrenXml: /\S/.test(remainingXml) ? remainingXml : "",
  };
}

function parseFill(fillXml: string): ParsedFill {
  if (/^<fill\b[^>]*\/>$/.test(fillXml)) {
    return {
      definition: buildEmptyFillDefinition(),
      extraChildrenXml: "",
    };
  }

  const innerXml = fillXml.match(/<fill\b[^>]*>([\s\S]*?)<\/fill>/)?.[1] ?? "";
  const patternFillMatch = innerXml.match(/<patternFill\b([^>]*?)(?:\/>|>([\s\S]*?)<\/patternFill>)/);
  if (!patternFillMatch || patternFillMatch.index === undefined) {
    return {
      definition: buildEmptyFillDefinition(),
      extraChildrenXml: /\S/.test(innerXml) ? innerXml : "",
    };
  }

  const patternAttributes = parseAttributes(patternFillMatch[1]);
  const patternInnerXml = patternFillMatch[2] ?? "";
  const remainingXml =
    innerXml.slice(0, patternFillMatch.index) + innerXml.slice(patternFillMatch.index + patternFillMatch[0].length);

  return {
    definition: {
      patternType: findAttributeValue(patternAttributes, "patternType") ?? null,
      fgColor: parseFillColorDefinition(patternInnerXml.match(/<fgColor\b([^>]*?)(?:\/>|>[\s\S]*?<\/fgColor>)/)?.[0] ?? null),
      bgColor: parseFillColorDefinition(patternInnerXml.match(/<bgColor\b([^>]*?)(?:\/>|>[\s\S]*?<\/bgColor>)/)?.[0] ?? null),
    },
    extraChildrenXml: /\S/.test(remainingXml) ? remainingXml : "",
  };
}

function parseFont(fontXml: string): ParsedFont {
  if (/^<font\b[^>]*\/>$/.test(fontXml)) {
    return {
      definition: buildEmptyFontDefinition(),
      extraChildrenXml: "",
    };
  }

  const innerXml = fontXml.match(/<font\b[^>]*>([\s\S]*?)<\/font>/)?.[1] ?? "";
  let remainingXml = innerXml;

  const [boldXml, remainingAfterBold] = takeFirstTag(remainingXml, /<b\b[^>]*?(?:\/>|>[\s\S]*?<\/b>)/);
  remainingXml = remainingAfterBold;
  const [italicXml, remainingAfterItalic] = takeFirstTag(remainingXml, /<i\b[^>]*?(?:\/>|>[\s\S]*?<\/i>)/);
  remainingXml = remainingAfterItalic;
  const [underlineXml, remainingAfterUnderline] = takeFirstTag(remainingXml, /<u\b([^>]*?)(?:\/>|>[\s\S]*?<\/u>)/);
  remainingXml = remainingAfterUnderline;
  const [strikeXml, remainingAfterStrike] = takeFirstTag(remainingXml, /<strike\b[^>]*?(?:\/>|>[\s\S]*?<\/strike>)/);
  remainingXml = remainingAfterStrike;
  const [outlineXml, remainingAfterOutline] = takeFirstTag(remainingXml, /<outline\b[^>]*?(?:\/>|>[\s\S]*?<\/outline>)/);
  remainingXml = remainingAfterOutline;
  const [shadowXml, remainingAfterShadow] = takeFirstTag(remainingXml, /<shadow\b[^>]*?(?:\/>|>[\s\S]*?<\/shadow>)/);
  remainingXml = remainingAfterShadow;
  const [condenseXml, remainingAfterCondense] = takeFirstTag(remainingXml, /<condense\b[^>]*?(?:\/>|>[\s\S]*?<\/condense>)/);
  remainingXml = remainingAfterCondense;
  const [extendXml, remainingAfterExtend] = takeFirstTag(remainingXml, /<extend\b[^>]*?(?:\/>|>[\s\S]*?<\/extend>)/);
  remainingXml = remainingAfterExtend;
  const [colorXml, remainingAfterColor] = takeFirstTag(remainingXml, /<color\b([^>]*?)(?:\/>|>[\s\S]*?<\/color>)/);
  remainingXml = remainingAfterColor;
  const [sizeXml, remainingAfterSize] = takeFirstTag(remainingXml, /<sz\b([^>]*?)(?:\/>|>[\s\S]*?<\/sz>)/);
  remainingXml = remainingAfterSize;
  const [nameXml, remainingAfterName] = takeFirstTag(remainingXml, /<name\b([^>]*?)(?:\/>|>[\s\S]*?<\/name>)/);
  remainingXml = remainingAfterName;
  const [familyXml, remainingAfterFamily] = takeFirstTag(remainingXml, /<family\b([^>]*?)(?:\/>|>[\s\S]*?<\/family>)/);
  remainingXml = remainingAfterFamily;
  const [charsetXml, remainingAfterCharset] = takeFirstTag(remainingXml, /<charset\b([^>]*?)(?:\/>|>[\s\S]*?<\/charset>)/);
  remainingXml = remainingAfterCharset;
  const [schemeXml, remainingAfterScheme] = takeFirstTag(remainingXml, /<scheme\b([^>]*?)(?:\/>|>[\s\S]*?<\/scheme>)/);
  remainingXml = remainingAfterScheme;
  const [vertAlignXml, remainingAfterVertAlign] = takeFirstTag(
    remainingXml,
    /<vertAlign\b([^>]*?)(?:\/>|>[\s\S]*?<\/vertAlign>)/,
  );
  remainingXml = remainingAfterVertAlign;

  return {
    definition: {
      bold: boldXml ? true : null,
      italic: italicXml ? true : null,
      underline: parseUnderlineValue(underlineXml),
      strike: strikeXml ? true : null,
      outline: outlineXml ? true : null,
      shadow: shadowXml ? true : null,
      condense: condenseXml ? true : null,
      extend: extendXml ? true : null,
      size: parseTagValNumber(sizeXml),
      name: parseTagValString(nameXml),
      family: parseTagValNumber(familyXml),
      charset: parseTagValNumber(charsetXml),
      scheme: parseTagValString(schemeXml),
      vertAlign: parseTagValString(vertAlignXml),
      color: parseFontColorDefinition(colorXml),
    },
    extraChildrenXml: /\S/.test(remainingXml) ? remainingXml : "",
  };
}

function parseCellStyle(attributesSource: string, innerXml: string): ParsedCellStyle {
  const attributes = parseAttributes(attributesSource);
  const alignmentMatch = innerXml.match(/<alignment\b([^>]*?)(?:\/>|>[\s\S]*?<\/alignment>)/);
  const alignmentAttributes = alignmentMatch ? parseAttributes(alignmentMatch[1]) : null;
  const extraChildrenXml =
    alignmentMatch && alignmentMatch.index !== undefined
      ? innerXml.slice(0, alignmentMatch.index) + innerXml.slice(alignmentMatch.index + alignmentMatch[0].length)
      : innerXml;

  return {
    alignmentAttributes,
    attributes,
    definition: {
      numFmtId: parseRequiredIntegerAttribute(attributes, "numFmtId", 0),
      fontId: parseRequiredIntegerAttribute(attributes, "fontId", 0),
      fillId: parseRequiredIntegerAttribute(attributes, "fillId", 0),
      borderId: parseRequiredIntegerAttribute(attributes, "borderId", 0),
      xfId: parseOptionalIntegerAttribute(attributes, "xfId"),
      quotePrefix: parseOptionalBooleanAttribute(attributes, "quotePrefix"),
      pivotButton: parseOptionalBooleanAttribute(attributes, "pivotButton"),
      applyNumberFormat: parseOptionalBooleanAttribute(attributes, "applyNumberFormat"),
      applyFont: parseOptionalBooleanAttribute(attributes, "applyFont"),
      applyFill: parseOptionalBooleanAttribute(attributes, "applyFill"),
      applyBorder: parseOptionalBooleanAttribute(attributes, "applyBorder"),
      applyAlignment: parseOptionalBooleanAttribute(attributes, "applyAlignment"),
      applyProtection: parseOptionalBooleanAttribute(attributes, "applyProtection"),
      alignment: alignmentAttributes ? parseAlignmentDefinition(alignmentAttributes) : null,
    },
    extraChildrenXml,
  };
}

function parseAlignmentDefinition(attributes: Array<[string, string]>): CellStyleAlignment {
  const alignment: CellStyleAlignment = {};

  assignStringAttribute(alignment, "horizontal", findAttributeValue(attributes, "horizontal"));
  assignStringAttribute(alignment, "vertical", findAttributeValue(attributes, "vertical"));
  assignNumberAttribute(alignment, "textRotation", findAttributeValue(attributes, "textRotation"));
  assignBooleanAttribute(alignment, "wrapText", findAttributeValue(attributes, "wrapText"));
  assignBooleanAttribute(alignment, "shrinkToFit", findAttributeValue(attributes, "shrinkToFit"));
  assignNumberAttribute(alignment, "indent", findAttributeValue(attributes, "indent"));
  assignNumberAttribute(alignment, "relativeIndent", findAttributeValue(attributes, "relativeIndent"));
  assignBooleanAttribute(alignment, "justifyLastLine", findAttributeValue(attributes, "justifyLastLine"));
  assignNumberAttribute(alignment, "readingOrder", findAttributeValue(attributes, "readingOrder"));

  return alignment;
}

function appendFontToStylesXml(stylesXml: string, fontXml: string): string {
  const fontsMatch = stylesXml.match(/<fonts\b([^>]*)>([\s\S]*?)<\/fonts>/);
  if (!fontsMatch || fontsMatch.index === undefined) {
    throw new XlsxError("styles.xml is missing <fonts>");
  }

  const attributes = parseAttributes(fontsMatch[1]);
  const nextCount = Array.from(fontsMatch[2].matchAll(/<font\b/g)).length + 1;
  const nextAttributes = upsertAttribute(attributes, "count", String(nextCount));
  const serializedAttributes = serializeAttributes(nextAttributes);
  const trailingWhitespace = fontsMatch[2].match(/\s*$/)?.[0] ?? "";
  const innerXmlWithoutTrailing = fontsMatch[2].slice(0, fontsMatch[2].length - trailingWhitespace.length);
  const closingIndentMatch = trailingWhitespace.match(/\n([ \t]*)$/);
  const entryPrefix = closingIndentMatch ? `\n${closingIndentMatch[1]}  ` : "";
  const nextInnerXml = `${innerXmlWithoutTrailing}${entryPrefix}${fontXml}${trailingWhitespace}`;
  const nextFontsXml = `<fonts${serializedAttributes ? ` ${serializedAttributes}` : ""}>${nextInnerXml}</fonts>`;

  return stylesXml.slice(0, fontsMatch.index) + nextFontsXml + stylesXml.slice(fontsMatch.index + fontsMatch[0].length);
}

function appendFillToStylesXml(stylesXml: string, fillXml: string): string {
  const fillsMatch = stylesXml.match(/<fills\b([^>]*)>([\s\S]*?)<\/fills>/);
  if (!fillsMatch || fillsMatch.index === undefined) {
    throw new XlsxError("styles.xml is missing <fills>");
  }

  const attributes = parseAttributes(fillsMatch[1]);
  const nextCount = Array.from(fillsMatch[2].matchAll(/<fill\b/g)).length + 1;
  const nextAttributes = upsertAttribute(attributes, "count", String(nextCount));
  const serializedAttributes = serializeAttributes(nextAttributes);
  const trailingWhitespace = fillsMatch[2].match(/\s*$/)?.[0] ?? "";
  const innerXmlWithoutTrailing = fillsMatch[2].slice(0, fillsMatch[2].length - trailingWhitespace.length);
  const closingIndentMatch = trailingWhitespace.match(/\n([ \t]*)$/);
  const entryPrefix = closingIndentMatch ? `\n${closingIndentMatch[1]}  ` : "";
  const nextInnerXml = `${innerXmlWithoutTrailing}${entryPrefix}${fillXml}${trailingWhitespace}`;
  const nextFillsXml = `<fills${serializedAttributes ? ` ${serializedAttributes}` : ""}>${nextInnerXml}</fills>`;

  return stylesXml.slice(0, fillsMatch.index) + nextFillsXml + stylesXml.slice(fillsMatch.index + fillsMatch[0].length);
}

function appendBorderToStylesXml(stylesXml: string, borderXml: string): string {
  const bordersMatch = stylesXml.match(/<borders\b([^>]*)>([\s\S]*?)<\/borders>/);
  if (!bordersMatch || bordersMatch.index === undefined) {
    throw new XlsxError("styles.xml is missing <borders>");
  }

  const attributes = parseAttributes(bordersMatch[1]);
  const nextCount = Array.from(bordersMatch[2].matchAll(/<border\b/g)).length + 1;
  const nextAttributes = upsertAttribute(attributes, "count", String(nextCount));
  const serializedAttributes = serializeAttributes(nextAttributes);
  const trailingWhitespace = bordersMatch[2].match(/\s*$/)?.[0] ?? "";
  const innerXmlWithoutTrailing = bordersMatch[2].slice(0, bordersMatch[2].length - trailingWhitespace.length);
  const closingIndentMatch = trailingWhitespace.match(/\n([ \t]*)$/);
  const entryPrefix = closingIndentMatch ? `\n${closingIndentMatch[1]}  ` : "";
  const nextInnerXml = `${innerXmlWithoutTrailing}${entryPrefix}${borderXml}${trailingWhitespace}`;
  const nextBordersXml = `<borders${serializedAttributes ? ` ${serializedAttributes}` : ""}>${nextInnerXml}</borders>`;

  return (
    stylesXml.slice(0, bordersMatch.index) +
    nextBordersXml +
    stylesXml.slice(bordersMatch.index + bordersMatch[0].length)
  );
}

function replaceFontInStylesXml(stylesXml: string, fontId: number, fontXml: string): string {
  const fontsMatch = stylesXml.match(/<fonts\b([^>]*)>([\s\S]*?)<\/fonts>/);
  if (!fontsMatch || fontsMatch.index === undefined) {
    throw new XlsxError("styles.xml is missing <fonts>");
  }

  const innerXml = fontsMatch[2];
  let currentFontIndex = 0;

  for (const match of innerXml.matchAll(/<font\b[^>]*?>([\s\S]*?)<\/font>|<font\b[^>]*?\/>/g)) {
    if (currentFontIndex !== fontId) {
      currentFontIndex += 1;
      continue;
    }

    if (match.index === undefined) {
      break;
    }

    const nextInnerXml = innerXml.slice(0, match.index) + fontXml + innerXml.slice(match.index + match[0].length);
    const nextFontsXml = `<fonts${fontsMatch[1]}>${nextInnerXml}</fonts>`;
    return stylesXml.slice(0, fontsMatch.index) + nextFontsXml + stylesXml.slice(fontsMatch.index + fontsMatch[0].length);
  }

  throw new XlsxError(`Font not found: ${fontId}`);
}

function replaceFillInStylesXml(stylesXml: string, fillId: number, fillXml: string): string {
  const fillsMatch = stylesXml.match(/<fills\b([^>]*)>([\s\S]*?)<\/fills>/);
  if (!fillsMatch || fillsMatch.index === undefined) {
    throw new XlsxError("styles.xml is missing <fills>");
  }

  const innerXml = fillsMatch[2];
  let currentFillIndex = 0;

  for (const match of innerXml.matchAll(/<fill\b[^>]*?>([\s\S]*?)<\/fill>|<fill\b[^>]*?\/>/g)) {
    if (currentFillIndex !== fillId) {
      currentFillIndex += 1;
      continue;
    }

    if (match.index === undefined) {
      break;
    }

    const nextInnerXml = innerXml.slice(0, match.index) + fillXml + innerXml.slice(match.index + match[0].length);
    const nextFillsXml = `<fills${fillsMatch[1]}>${nextInnerXml}</fills>`;
    return stylesXml.slice(0, fillsMatch.index) + nextFillsXml + stylesXml.slice(fillsMatch.index + fillsMatch[0].length);
  }

  throw new XlsxError(`Fill not found: ${fillId}`);
}

function replaceBorderInStylesXml(stylesXml: string, borderId: number, borderXml: string): string {
  const bordersMatch = stylesXml.match(/<borders\b([^>]*)>([\s\S]*?)<\/borders>/);
  if (!bordersMatch || bordersMatch.index === undefined) {
    throw new XlsxError("styles.xml is missing <borders>");
  }

  const innerXml = bordersMatch[2];
  let currentBorderIndex = 0;

  for (const match of innerXml.matchAll(/<border\b[^>]*?>([\s\S]*?)<\/border>|<border\b[^>]*?\/>/g)) {
    if (currentBorderIndex !== borderId) {
      currentBorderIndex += 1;
      continue;
    }

    if (match.index === undefined) {
      break;
    }

    const nextInnerXml = innerXml.slice(0, match.index) + borderXml + innerXml.slice(match.index + match[0].length);
    const nextBordersXml = `<borders${bordersMatch[1]}>${nextInnerXml}</borders>`;
    return (
      stylesXml.slice(0, bordersMatch.index) +
      nextBordersXml +
      stylesXml.slice(bordersMatch.index + bordersMatch[0].length)
    );
  }

  throw new XlsxError(`Border not found: ${borderId}`);
}

function upsertNumberFormatInStylesXml(stylesXml: string, numFmtId: number, formatCode: string): string {
  const numFmtXml = `<numFmt numFmtId="${numFmtId}" formatCode="${escapeXmlText(formatCode)}"/>`;
  const numberFormatsMatch = stylesXml.match(/<numFmts\b([^>]*)>([\s\S]*?)<\/numFmts>/);

  if (!numberFormatsMatch || numberFormatsMatch.index === undefined) {
    const fontsMatch = stylesXml.match(/<fonts\b/);
    if (!fontsMatch || fontsMatch.index === undefined) {
      throw new XlsxError("styles.xml is missing <fonts>");
    }

    return (
      stylesXml.slice(0, fontsMatch.index) +
      `<numFmts count="1">${numFmtXml}</numFmts>` +
      stylesXml.slice(fontsMatch.index)
    );
  }

  const innerXml = numberFormatsMatch[2];
  let found = false;
  const nextInnerXml = innerXml.replace(/<numFmt\b([^>]*)\/>/g, (match, attributesSource) => {
    const currentNumFmtId = getXmlAttr(attributesSource, "numFmtId");
    if (currentNumFmtId === undefined || Number(currentNumFmtId) !== numFmtId) {
      return match;
    }

    found = true;
    return numFmtXml;
  });

  const finalInnerXml = found ? nextInnerXml : `${nextInnerXml}${numFmtXml}`;
  const nextCount = Array.from(finalInnerXml.matchAll(/<numFmt\b/g)).length;
  const nextAttributes = upsertAttribute(parseAttributes(numberFormatsMatch[1]), "count", String(nextCount));
  const serializedAttributes = serializeAttributes(nextAttributes);
  const nextNumFmtsXml = `<numFmts${serializedAttributes ? ` ${serializedAttributes}` : ""}>${finalInnerXml}</numFmts>`;

  return (
    stylesXml.slice(0, numberFormatsMatch.index) +
    nextNumFmtsXml +
    stylesXml.slice(numberFormatsMatch.index + numberFormatsMatch[0].length)
  );
}

function appendCellXfToStylesXml(stylesXml: string, xfXml: string): string {
  const cellXfsMatch = stylesXml.match(/<cellXfs\b([^>]*)>([\s\S]*?)<\/cellXfs>/);
  if (!cellXfsMatch || cellXfsMatch.index === undefined) {
    throw new XlsxError("styles.xml is missing <cellXfs>");
  }

  const attributes = parseAttributes(cellXfsMatch[1]);
  const nextCount = Array.from(cellXfsMatch[2].matchAll(/<xf\b/g)).length + 1;
  const nextAttributes = upsertAttribute(attributes, "count", String(nextCount));
  const serializedAttributes = serializeAttributes(nextAttributes);
  const trailingWhitespace = cellXfsMatch[2].match(/\s*$/)?.[0] ?? "";
  const innerXmlWithoutTrailing = cellXfsMatch[2].slice(0, cellXfsMatch[2].length - trailingWhitespace.length);
  const closingIndentMatch = trailingWhitespace.match(/\n([ \t]*)$/);
  const entryPrefix = closingIndentMatch ? `\n${closingIndentMatch[1]}  ` : "";
  const nextInnerXml = `${innerXmlWithoutTrailing}${entryPrefix}${xfXml}${trailingWhitespace}`;
  const nextCellXfsXml = `<cellXfs${serializedAttributes ? ` ${serializedAttributes}` : ""}>${nextInnerXml}</cellXfs>`;

  return (
    stylesXml.slice(0, cellXfsMatch.index) +
    nextCellXfsXml +
    stylesXml.slice(cellXfsMatch.index + cellXfsMatch[0].length)
  );
}

function replaceCellXfInStylesXml(stylesXml: string, styleId: number, xfXml: string): string {
  const cellXfsMatch = stylesXml.match(/<cellXfs\b([^>]*)>([\s\S]*?)<\/cellXfs>/);
  if (!cellXfsMatch || cellXfsMatch.index === undefined) {
    throw new XlsxError("styles.xml is missing <cellXfs>");
  }

  const innerXml = cellXfsMatch[2];
  let xfIndex = 0;

  for (const match of innerXml.matchAll(/<xf\b([^>]*?)(?:\/>|>([\s\S]*?)<\/xf>)/g)) {
    if (xfIndex !== styleId) {
      xfIndex += 1;
      continue;
    }

    if (match.index === undefined) {
      break;
    }

    const nextInnerXml =
      innerXml.slice(0, match.index) + xfXml + innerXml.slice(match.index + match[0].length);
    const nextCellXfsXml = `<cellXfs${cellXfsMatch[1]}>${nextInnerXml}</cellXfs>`;
    return (
      stylesXml.slice(0, cellXfsMatch.index) +
      nextCellXfsXml +
      stylesXml.slice(cellXfsMatch.index + cellXfsMatch[0].length)
    );
  }

  throw new XlsxError(`Style not found: ${styleId}`);
}

function buildPatchedCellXfXml(sourceStyle: ParsedCellStyle, patch: CellStylePatch): string {
  const attributes = applyCellStylePatch(sourceStyle.attributes, patch);
  const alignmentAttributes = applyAlignmentPatch(sourceStyle.alignmentAttributes, patch.alignment);
  const alignmentXml = alignmentAttributes ? buildSelfClosingTag("alignment", alignmentAttributes) : "";
  const innerXml = alignmentXml + sourceStyle.extraChildrenXml;
  const serializedAttributes = serializeAttributes(attributes);

  if (innerXml.length === 0) {
    return `<xf${serializedAttributes ? ` ${serializedAttributes}` : ""}/>`;
  }

  return `<xf${serializedAttributes ? ` ${serializedAttributes}` : ""}>${innerXml}</xf>`;
}

function buildPatchedFontXml(sourceFont: ParsedFont, patch: CellFontPatch): string {
  const font = applyFontPatch(sourceFont.definition, patch);
  const childXml = buildFontChildXml(font) + sourceFont.extraChildrenXml;
  return childXml.length === 0 ? "<font/>" : `<font>${childXml}</font>`;
}

function buildPatchedFillXml(sourceFill: ParsedFill, patch: CellFillPatch): string {
  const fill = applyFillPatch(sourceFill.definition, patch);
  const childXml = buildFillChildXml(fill) + sourceFill.extraChildrenXml;
  return childXml.length === 0 ? "<fill/>" : `<fill>${childXml}</fill>`;
}

function buildPatchedBorderXml(sourceBorder: ParsedBorder, patch: CellBorderPatch): string {
  const border = applyBorderPatch(sourceBorder.definition, patch);
  const attributes = buildBorderAttributes(border);
  const serializedAttributes = serializeAttributes(attributes);
  const childXml = buildBorderChildXml(border) + sourceBorder.extraChildrenXml;
  return childXml.length === 0
    ? `<border${serializedAttributes ? ` ${serializedAttributes}` : ""}/>`
    : `<border${serializedAttributes ? ` ${serializedAttributes}` : ""}>${childXml}</border>`;
}

function applyCellStylePatch(attributes: Array<[string, string]>, patch: CellStylePatch): Array<[string, string]> {
  let nextAttributes = [...attributes];

  nextAttributes = applyRequiredIntegerPatch(nextAttributes, "numFmtId", patch.numFmtId);
  nextAttributes = applyRequiredIntegerPatch(nextAttributes, "fontId", patch.fontId);
  nextAttributes = applyRequiredIntegerPatch(nextAttributes, "fillId", patch.fillId);
  nextAttributes = applyRequiredIntegerPatch(nextAttributes, "borderId", patch.borderId);
  nextAttributes = applyOptionalIntegerPatch(nextAttributes, "xfId", patch.xfId);
  nextAttributes = applyOptionalBooleanPatch(nextAttributes, "quotePrefix", patch.quotePrefix);
  nextAttributes = applyOptionalBooleanPatch(nextAttributes, "pivotButton", patch.pivotButton);
  nextAttributes = applyOptionalBooleanPatch(nextAttributes, "applyNumberFormat", patch.applyNumberFormat);
  nextAttributes = applyOptionalBooleanPatch(nextAttributes, "applyFont", patch.applyFont);
  nextAttributes = applyOptionalBooleanPatch(nextAttributes, "applyFill", patch.applyFill);
  nextAttributes = applyOptionalBooleanPatch(nextAttributes, "applyBorder", patch.applyBorder);
  nextAttributes = applyOptionalBooleanPatch(nextAttributes, "applyAlignment", patch.applyAlignment);
  nextAttributes = applyOptionalBooleanPatch(nextAttributes, "applyProtection", patch.applyProtection);

  return nextAttributes;
}

function applyAlignmentPatch(
  attributes: Array<[string, string]> | null,
  patch: CellStyleAlignmentPatch | null | undefined,
): Array<[string, string]> | null {
  if (patch === undefined) {
    return attributes ? [...attributes] : null;
  }

  if (patch === null) {
    return null;
  }

  let nextAttributes = attributes ? [...attributes] : [];
  nextAttributes = applyOptionalStringPatch(nextAttributes, "horizontal", patch.horizontal);
  nextAttributes = applyOptionalStringPatch(nextAttributes, "vertical", patch.vertical);
  nextAttributes = applyOptionalIntegerPatch(nextAttributes, "textRotation", patch.textRotation);
  nextAttributes = applyOptionalBooleanPatch(nextAttributes, "wrapText", patch.wrapText);
  nextAttributes = applyOptionalBooleanPatch(nextAttributes, "shrinkToFit", patch.shrinkToFit);
  nextAttributes = applyOptionalIntegerPatch(nextAttributes, "indent", patch.indent);
  nextAttributes = applyOptionalIntegerPatch(nextAttributes, "relativeIndent", patch.relativeIndent);
  nextAttributes = applyOptionalBooleanPatch(nextAttributes, "justifyLastLine", patch.justifyLastLine);
  nextAttributes = applyOptionalIntegerPatch(nextAttributes, "readingOrder", patch.readingOrder);

  return nextAttributes.length === 0 ? null : nextAttributes;
}

function buildSelfClosingTag(tagName: string, attributes: Array<[string, string]>): string {
  const serializedAttributes = serializeAttributes(attributes);
  return `<${tagName}${serializedAttributes ? ` ${serializedAttributes}` : ""}/>`;
}

function buildFontChildXml(font: CellFontDefinition): string {
  const parts: string[] = [];

  if (font.bold) {
    parts.push("<b/>");
  }
  if (font.italic) {
    parts.push("<i/>");
  }
  if (font.underline !== null) {
    parts.push(font.underline === "single" ? "<u/>" : buildSelfClosingTag("u", [["val", font.underline]]));
  }
  if (font.strike) {
    parts.push("<strike/>");
  }
  if (font.outline) {
    parts.push("<outline/>");
  }
  if (font.shadow) {
    parts.push("<shadow/>");
  }
  if (font.condense) {
    parts.push("<condense/>");
  }
  if (font.extend) {
    parts.push("<extend/>");
  }
  if (font.color) {
    parts.push(buildSelfClosingTag("color", buildFontColorAttributes(font.color)));
  }
  if (font.size !== null) {
    parts.push(buildSelfClosingTag("sz", [["val", String(font.size)]]));
  }
  if (font.name !== null) {
    parts.push(buildSelfClosingTag("name", [["val", font.name]]));
  }
  if (font.family !== null) {
    parts.push(buildSelfClosingTag("family", [["val", String(font.family)]]));
  }
  if (font.charset !== null) {
    parts.push(buildSelfClosingTag("charset", [["val", String(font.charset)]]));
  }
  if (font.scheme !== null) {
    parts.push(buildSelfClosingTag("scheme", [["val", font.scheme]]));
  }
  if (font.vertAlign !== null) {
    parts.push(buildSelfClosingTag("vertAlign", [["val", font.vertAlign]]));
  }

  return parts.join("");
}

function buildFillChildXml(fill: CellFillDefinition): string {
  if (fill.patternType === null && fill.fgColor === null && fill.bgColor === null) {
    return "";
  }

  const attributes = fill.patternType === null ? [] : ([["patternType", fill.patternType]] as Array<[string, string]>);
  const colorXml =
    (fill.fgColor ? buildSelfClosingTag("fgColor", buildFillColorAttributes(fill.fgColor)) : "") +
    (fill.bgColor ? buildSelfClosingTag("bgColor", buildFillColorAttributes(fill.bgColor)) : "");

  if (colorXml.length === 0) {
    return buildSelfClosingTag("patternFill", attributes);
  }

  const serializedAttributes = serializeAttributes(attributes);
  return `<patternFill${serializedAttributes ? ` ${serializedAttributes}` : ""}>${colorXml}</patternFill>`;
}

function buildBorderChildXml(border: CellBorderDefinition): string {
  return [
    buildBorderSideXml("left", border.left),
    buildBorderSideXml("right", border.right),
    buildBorderSideXml("top", border.top),
    buildBorderSideXml("bottom", border.bottom),
    buildBorderSideXml("diagonal", border.diagonal),
    buildBorderSideXml("vertical", border.vertical),
    buildBorderSideXml("horizontal", border.horizontal),
  ].join("");
}

function buildBorderAttributes(border: CellBorderDefinition): Array<[string, string]> {
  const attributes: Array<[string, string]> = [];
  if (border.diagonalUp !== null) {
    attributes.push(["diagonalUp", border.diagonalUp ? "1" : "0"]);
  }
  if (border.diagonalDown !== null) {
    attributes.push(["diagonalDown", border.diagonalDown ? "1" : "0"]);
  }
  if (border.outline !== null) {
    attributes.push(["outline", border.outline ? "1" : "0"]);
  }
  return attributes;
}

function buildBorderSideXml(tagName: string, side: CellBorderSideDefinition | null): string {
  if (side === null) {
    return "";
  }

  const attributes: Array<[string, string]> = [];
  if (side.style !== null) {
    attributes.push(["style", side.style]);
  }

  const colorXml = side.color ? buildSelfClosingTag("color", buildBorderColorAttributes(side.color)) : "";
  if (colorXml.length === 0) {
    return buildSelfClosingTag(tagName, attributes);
  }

  const serializedAttributes = serializeAttributes(attributes);
  return `<${tagName}${serializedAttributes ? ` ${serializedAttributes}` : ""}>${colorXml}</${tagName}>`;
}

function cloneCellStyleDefinition(style: CellStyleDefinition | null): CellStyleDefinition | null {
  if (!style) {
    return null;
  }

  return {
    ...style,
    alignment: style.alignment ? { ...style.alignment } : null,
  };
}

function cloneCellFontDefinition(font: CellFontDefinition | null): CellFontDefinition | null {
  if (!font) {
    return null;
  }

  return {
    ...font,
    color: font.color ? { ...font.color } : null,
  };
}

function cloneCellFillDefinition(fill: CellFillDefinition | null): CellFillDefinition | null {
  if (!fill) {
    return null;
  }

  return {
    ...fill,
    fgColor: fill.fgColor ? { ...fill.fgColor } : null,
    bgColor: fill.bgColor ? { ...fill.bgColor } : null,
  };
}

function cloneCellBorderDefinition(border: CellBorderDefinition | null): CellBorderDefinition | null {
  if (!border) {
    return null;
  }

  return {
    left: cloneCellBorderSideDefinition(border.left),
    right: cloneCellBorderSideDefinition(border.right),
    top: cloneCellBorderSideDefinition(border.top),
    bottom: cloneCellBorderSideDefinition(border.bottom),
    diagonal: cloneCellBorderSideDefinition(border.diagonal),
    vertical: cloneCellBorderSideDefinition(border.vertical),
    horizontal: cloneCellBorderSideDefinition(border.horizontal),
    diagonalUp: border.diagonalUp,
    diagonalDown: border.diagonalDown,
    outline: border.outline,
  };
}

function buildEmptyFontDefinition(): CellFontDefinition {
  return {
    bold: null,
    italic: null,
    underline: null,
    strike: null,
    outline: null,
    shadow: null,
    condense: null,
    extend: null,
    size: null,
    name: null,
    family: null,
    charset: null,
    scheme: null,
    vertAlign: null,
    color: null,
  };
}

function buildEmptyFillDefinition(): CellFillDefinition {
  return {
    patternType: null,
    fgColor: null,
    bgColor: null,
  };
}

function buildEmptyBorderDefinition(): CellBorderDefinition {
  return {
    left: null,
    right: null,
    top: null,
    bottom: null,
    diagonal: null,
    vertical: null,
    horizontal: null,
    diagonalUp: null,
    diagonalDown: null,
    outline: null,
  };
}

function applyFontPatch(sourceFont: CellFontDefinition, patch: CellFontPatch): CellFontDefinition {
  return {
    bold: patch.bold === undefined ? sourceFont.bold : patch.bold,
    italic: patch.italic === undefined ? sourceFont.italic : patch.italic,
    underline: patch.underline === undefined ? sourceFont.underline : patch.underline,
    strike: patch.strike === undefined ? sourceFont.strike : patch.strike,
    outline: patch.outline === undefined ? sourceFont.outline : patch.outline,
    shadow: patch.shadow === undefined ? sourceFont.shadow : patch.shadow,
    condense: patch.condense === undefined ? sourceFont.condense : patch.condense,
    extend: patch.extend === undefined ? sourceFont.extend : patch.extend,
    size: patch.size === undefined ? sourceFont.size : patch.size,
    name: patch.name === undefined ? sourceFont.name : patch.name,
    family: patch.family === undefined ? sourceFont.family : patch.family,
    charset: patch.charset === undefined ? sourceFont.charset : patch.charset,
    scheme: patch.scheme === undefined ? sourceFont.scheme : patch.scheme,
    vertAlign: patch.vertAlign === undefined ? sourceFont.vertAlign : patch.vertAlign,
    color: applyFontColorPatch(sourceFont.color, patch.color),
  };
}

function applyFillPatch(sourceFill: CellFillDefinition, patch: CellFillPatch): CellFillDefinition {
  return {
    patternType: patch.patternType === undefined ? sourceFill.patternType : patch.patternType,
    fgColor: applyFillColorPatch(sourceFill.fgColor, patch.fgColor),
    bgColor: applyFillColorPatch(sourceFill.bgColor, patch.bgColor),
  };
}

function applyBorderPatch(sourceBorder: CellBorderDefinition, patch: CellBorderPatch): CellBorderDefinition {
  return {
    left: applyBorderSidePatch(sourceBorder.left, patch.left),
    right: applyBorderSidePatch(sourceBorder.right, patch.right),
    top: applyBorderSidePatch(sourceBorder.top, patch.top),
    bottom: applyBorderSidePatch(sourceBorder.bottom, patch.bottom),
    diagonal: applyBorderSidePatch(sourceBorder.diagonal, patch.diagonal),
    vertical: applyBorderSidePatch(sourceBorder.vertical, patch.vertical),
    horizontal: applyBorderSidePatch(sourceBorder.horizontal, patch.horizontal),
    diagonalUp: patch.diagonalUp === undefined ? sourceBorder.diagonalUp : patch.diagonalUp,
    diagonalDown: patch.diagonalDown === undefined ? sourceBorder.diagonalDown : patch.diagonalDown,
    outline: patch.outline === undefined ? sourceBorder.outline : patch.outline,
  };
}

function applyFontColorPatch(
  sourceColor: CellFontColor | null,
  patch: CellFontColorPatch | null | undefined,
): CellFontColor | null {
  if (patch === undefined) {
    return sourceColor ? { ...sourceColor } : null;
  }
  if (patch === null) {
    return null;
  }

  const nextColor: CellFontColor = sourceColor ? { ...sourceColor } : {};
  updateOptionalObjectProperty(nextColor, "rgb", patch.rgb);
  updateOptionalObjectProperty(nextColor, "theme", patch.theme);
  updateOptionalObjectProperty(nextColor, "indexed", patch.indexed);
  updateOptionalObjectProperty(nextColor, "auto", patch.auto);
  updateOptionalObjectProperty(nextColor, "tint", patch.tint);

  return Object.keys(nextColor).length === 0 ? null : nextColor;
}

function buildFontColorAttributes(color: CellFontColor): Array<[string, string]> {
  const attributes: Array<[string, string]> = [];
  if (color.rgb !== undefined) {
    attributes.push(["rgb", color.rgb]);
  }
  if (color.theme !== undefined) {
    attributes.push(["theme", String(color.theme)]);
  }
  if (color.indexed !== undefined) {
    attributes.push(["indexed", String(color.indexed)]);
  }
  if (color.auto !== undefined) {
    attributes.push(["auto", color.auto ? "1" : "0"]);
  }
  if (color.tint !== undefined) {
    attributes.push(["tint", String(color.tint)]);
  }
  return attributes;
}

function applyFillColorPatch(
  sourceColor: CellFillColor | null,
  patch: CellFillColorPatch | null | undefined,
): CellFillColor | null {
  if (patch === undefined) {
    return sourceColor ? { ...sourceColor } : null;
  }
  if (patch === null) {
    return null;
  }

  const nextColor: CellFillColor = sourceColor ? { ...sourceColor } : {};
  updateOptionalObjectProperty(nextColor, "rgb", patch.rgb);
  updateOptionalObjectProperty(nextColor, "theme", patch.theme);
  updateOptionalObjectProperty(nextColor, "indexed", patch.indexed);
  updateOptionalObjectProperty(nextColor, "auto", patch.auto);
  updateOptionalObjectProperty(nextColor, "tint", patch.tint);

  return Object.keys(nextColor).length === 0 ? null : nextColor;
}

function buildFillColorAttributes(color: CellFillColor): Array<[string, string]> {
  const attributes: Array<[string, string]> = [];
  if (color.rgb !== undefined) {
    attributes.push(["rgb", color.rgb]);
  }
  if (color.theme !== undefined) {
    attributes.push(["theme", String(color.theme)]);
  }
  if (color.indexed !== undefined) {
    attributes.push(["indexed", String(color.indexed)]);
  }
  if (color.auto !== undefined) {
    attributes.push(["auto", color.auto ? "1" : "0"]);
  }
  if (color.tint !== undefined) {
    attributes.push(["tint", String(color.tint)]);
  }
  return attributes;
}

function applyBorderSidePatch(
  sourceSide: CellBorderSideDefinition | null,
  patch: CellBorderSidePatch | null | undefined,
): CellBorderSideDefinition | null {
  if (patch === undefined) {
    return cloneCellBorderSideDefinition(sourceSide);
  }
  if (patch === null) {
    return null;
  }

  return {
    style: patch.style === undefined ? (sourceSide?.style ?? null) : patch.style,
    color: applyBorderColorPatch(sourceSide?.color ?? null, patch.color),
  };
}

function applyBorderColorPatch(
  sourceColor: CellBorderColor | null,
  patch: CellBorderColorPatch | null | undefined,
): CellBorderColor | null {
  if (patch === undefined) {
    return sourceColor ? { ...sourceColor } : null;
  }
  if (patch === null) {
    return null;
  }

  const nextColor: CellBorderColor = sourceColor ? { ...sourceColor } : {};
  updateOptionalObjectProperty(nextColor, "rgb", patch.rgb);
  updateOptionalObjectProperty(nextColor, "theme", patch.theme);
  updateOptionalObjectProperty(nextColor, "indexed", patch.indexed);
  updateOptionalObjectProperty(nextColor, "auto", patch.auto);
  updateOptionalObjectProperty(nextColor, "tint", patch.tint);

  return Object.keys(nextColor).length === 0 ? null : nextColor;
}

function buildBorderColorAttributes(color: CellBorderColor): Array<[string, string]> {
  const attributes: Array<[string, string]> = [];
  if (color.rgb !== undefined) {
    attributes.push(["rgb", color.rgb]);
  }
  if (color.theme !== undefined) {
    attributes.push(["theme", String(color.theme)]);
  }
  if (color.indexed !== undefined) {
    attributes.push(["indexed", String(color.indexed)]);
  }
  if (color.auto !== undefined) {
    attributes.push(["auto", color.auto ? "1" : "0"]);
  }
  if (color.tint !== undefined) {
    attributes.push(["tint", String(color.tint)]);
  }
  return attributes;
}

function updateOptionalObjectProperty<T extends object, K extends keyof T>(
  target: T,
  key: K,
  value: T[K] | null | undefined,
): void {
  if (value === undefined) {
    return;
  }

  if (value === null) {
    delete target[key];
    return;
  }

  target[key] = value;
}

function cloneCellBorderSideDefinition(side: CellBorderSideDefinition | null): CellBorderSideDefinition | null {
  if (!side) {
    return null;
  }

  return {
    style: side.style,
    color: side.color ? { ...side.color } : null,
  };
}

function findAttributeValue(attributes: Array<[string, string]>, name: string): string | undefined {
  return attributes.find(([attributeName]) => attributeName === name)?.[1];
}

function takeFirstTag(xml: string, pattern: RegExp): [string | null, string] {
  const match = xml.match(pattern);
  if (!match || match.index === undefined) {
    return [null, xml];
  }

  return [match[0], xml.slice(0, match.index) + xml.slice(match.index + match[0].length)];
}

function parseTagValNumber(tagXml: string | null): number | null {
  if (!tagXml) {
    return null;
  }
  const value = getXmlAttr(tagXml, "val");
  return value === undefined ? null : Number(value);
}

function parseTagValString(tagXml: string | null): string | null {
  if (!tagXml) {
    return null;
  }
  return getXmlAttr(tagXml, "val") ?? null;
}

function parseUnderlineValue(tagXml: string | null): string | null {
  if (!tagXml) {
    return null;
  }
  return getXmlAttr(tagXml, "val") ?? "single";
}

function parseFontColorDefinition(tagXml: string | null): CellFontColor | null {
  if (!tagXml) {
    return null;
  }

  const color: CellFontColor = {};
  const rgb = getXmlAttr(tagXml, "rgb");
  const theme = getXmlAttr(tagXml, "theme");
  const indexed = getXmlAttr(tagXml, "indexed");
  const auto = getXmlAttr(tagXml, "auto");
  const tint = getXmlAttr(tagXml, "tint");

  if (rgb !== undefined) {
    color.rgb = rgb;
  }
  if (theme !== undefined) {
    color.theme = Number(theme);
  }
  if (indexed !== undefined) {
    color.indexed = Number(indexed);
  }
  if (auto !== undefined) {
    color.auto = auto === "1" || auto === "true";
  }
  if (tint !== undefined) {
    color.tint = Number(tint);
  }

  return Object.keys(color).length === 0 ? null : color;
}

function parseFillColorDefinition(tagXml: string | null): CellFillColor | null {
  if (!tagXml) {
    return null;
  }

  const color: CellFillColor = {};
  const rgb = getXmlAttr(tagXml, "rgb");
  const theme = getXmlAttr(tagXml, "theme");
  const indexed = getXmlAttr(tagXml, "indexed");
  const auto = getXmlAttr(tagXml, "auto");
  const tint = getXmlAttr(tagXml, "tint");

  if (rgb !== undefined) {
    color.rgb = rgb;
  }
  if (theme !== undefined) {
    color.theme = Number(theme);
  }
  if (indexed !== undefined) {
    color.indexed = Number(indexed);
  }
  if (auto !== undefined) {
    color.auto = auto === "1" || auto === "true";
  }
  if (tint !== undefined) {
    color.tint = Number(tint);
  }

  return Object.keys(color).length === 0 ? null : color;
}

function parseBorderSideDefinition(tagXml: string | null): CellBorderSideDefinition | null {
  if (!tagXml) {
    return null;
  }

  const style = getXmlAttr(tagXml, "style") ?? null;
  const colorMatch = tagXml.match(/<color\b([^>]*?)(?:\/>|>[\s\S]*?<\/color>)/);
  const color = parseBorderColorDefinition(colorMatch?.[0] ?? null);
  return {
    style,
    color,
  };
}

function parseBorderColorDefinition(tagXml: string | null): CellBorderColor | null {
  if (!tagXml) {
    return null;
  }

  const color: CellBorderColor = {};
  const rgb = getXmlAttr(tagXml, "rgb");
  const theme = getXmlAttr(tagXml, "theme");
  const indexed = getXmlAttr(tagXml, "indexed");
  const auto = getXmlAttr(tagXml, "auto");
  const tint = getXmlAttr(tagXml, "tint");

  if (rgb !== undefined) {
    color.rgb = rgb;
  }
  if (theme !== undefined) {
    color.theme = Number(theme);
  }
  if (indexed !== undefined) {
    color.indexed = Number(indexed);
  }
  if (auto !== undefined) {
    color.auto = auto === "1" || auto === "true";
  }
  if (tint !== undefined) {
    color.tint = Number(tint);
  }

  return Object.keys(color).length === 0 ? null : color;
}

function getNextCustomNumberFormatId(numberFormats: Map<number, string>): number {
  let nextNumFmtId = 164;

  for (const numFmtId of numberFormats.keys()) {
    nextNumFmtId = Math.max(nextNumFmtId, numFmtId + 1);
  }

  return nextNumFmtId;
}

function parseRequiredIntegerAttribute(
  attributes: Array<[string, string]>,
  name: string,
  fallback: number,
): number {
  const value = findAttributeValue(attributes, name);
  return value === undefined ? fallback : Number(value);
}

function parseOptionalIntegerAttribute(attributes: Array<[string, string]>, name: string): number | null {
  const value = findAttributeValue(attributes, name);
  return value === undefined ? null : Number(value);
}

function parseOptionalBooleanAttribute(attributes: Array<[string, string]>, name: string): boolean | null {
  const value = findAttributeValue(attributes, name);
  if (value === undefined) {
    return null;
  }

  return value === "1" || value === "true";
}

function assignStringAttribute(target: CellStyleAlignment, name: keyof CellStyleAlignment, value?: string): void {
  if (value !== undefined) {
    target[name] = value as never;
  }
}

function assignNumberAttribute(target: CellStyleAlignment, name: keyof CellStyleAlignment, value?: string): void {
  if (value !== undefined) {
    target[name] = Number(value) as never;
  }
}

function assignBooleanAttribute(target: CellStyleAlignment, name: keyof CellStyleAlignment, value?: string): void {
  if (value !== undefined) {
    target[name] = (value === "1" || value === "true") as never;
  }
}

function upsertAttribute(
  attributes: Array<[string, string]>,
  name: string,
  value: string | null,
): Array<[string, string]> {
  const nextAttributes: Array<[string, string]> = [];
  let found = false;

  for (const [attributeName, attributeValue] of attributes) {
    if (attributeName !== name) {
      nextAttributes.push([attributeName, attributeValue]);
      continue;
    }

    found = true;
    if (value !== null) {
      nextAttributes.push([attributeName, value]);
    }
  }

  if (!found && value !== null) {
    nextAttributes.push([name, value]);
  }

  return nextAttributes;
}

function applyRequiredIntegerPatch(
  attributes: Array<[string, string]>,
  name: string,
  value: number | undefined,
): Array<[string, string]> {
  return value === undefined ? attributes : upsertAttribute(attributes, name, String(value));
}

function applyOptionalIntegerPatch(
  attributes: Array<[string, string]>,
  name: string,
  value: number | null | undefined,
): Array<[string, string]> {
  return value === undefined ? attributes : upsertAttribute(attributes, name, value === null ? null : String(value));
}

function applyOptionalBooleanPatch(
  attributes: Array<[string, string]>,
  name: string,
  value: boolean | null | undefined,
): Array<[string, string]> {
  return value === undefined ? attributes : upsertAttribute(attributes, name, value === null ? null : value ? "1" : "0");
}

function applyOptionalStringPatch(
  attributes: Array<[string, string]>,
  name: string,
  value: string | null | undefined,
): Array<[string, string]> {
  return value === undefined ? attributes : upsertAttribute(attributes, name, value);
}

function assertSheetName(sheetName: string): void {
  if (sheetName.length === 0 || sheetName.length > 31 || /[\\/*?:[\]]/.test(sheetName)) {
    throw new XlsxError(`Invalid sheet name: ${sheetName}`);
  }
}

function assertStyleId(styleId: number): void {
  if (!Number.isInteger(styleId) || styleId < 0) {
    throw new XlsxError(`Invalid style id: ${styleId}`);
  }
}

function assertCellStylePatch(patch: CellStylePatch): void {
  assertOptionalNonNegativeInteger(patch.numFmtId, "numFmtId");
  assertOptionalNonNegativeInteger(patch.fontId, "fontId");
  assertOptionalNonNegativeInteger(patch.fillId, "fillId");
  assertOptionalNonNegativeInteger(patch.borderId, "borderId");
  assertOptionalNullableNonNegativeInteger(patch.xfId, "xfId");
  assertOptionalNullableBoolean(patch.quotePrefix, "quotePrefix");
  assertOptionalNullableBoolean(patch.pivotButton, "pivotButton");
  assertOptionalNullableBoolean(patch.applyNumberFormat, "applyNumberFormat");
  assertOptionalNullableBoolean(patch.applyFont, "applyFont");
  assertOptionalNullableBoolean(patch.applyFill, "applyFill");
  assertOptionalNullableBoolean(patch.applyBorder, "applyBorder");
  assertOptionalNullableBoolean(patch.applyAlignment, "applyAlignment");
  assertOptionalNullableBoolean(patch.applyProtection, "applyProtection");

  if (patch.alignment !== undefined && patch.alignment !== null) {
    assertCellStyleAlignmentPatch(patch.alignment);
  }
}

function assertCellFontPatch(patch: CellFontPatch): void {
  assertOptionalNullableBoolean(patch.bold, "bold");
  assertOptionalNullableBoolean(patch.italic, "italic");
  assertOptionalNullableString(patch.underline, "underline");
  assertOptionalNullableBoolean(patch.strike, "strike");
  assertOptionalNullableBoolean(patch.outline, "outline");
  assertOptionalNullableBoolean(patch.shadow, "shadow");
  assertOptionalNullableBoolean(patch.condense, "condense");
  assertOptionalNullableBoolean(patch.extend, "extend");
  assertOptionalNullableFiniteNumber(patch.size, "size");
  assertOptionalNullableString(patch.name, "name");
  assertOptionalNullableNonNegativeInteger(patch.family, "family");
  assertOptionalNullableNonNegativeInteger(patch.charset, "charset");
  assertOptionalNullableString(patch.scheme, "scheme");
  assertOptionalNullableString(patch.vertAlign, "vertAlign");

  if (patch.color !== undefined && patch.color !== null) {
    assertCellFontColorPatch(patch.color);
  }
}

function assertCellFillPatch(patch: CellFillPatch): void {
  assertOptionalNullableString(patch.patternType, "patternType");

  if (patch.fgColor !== undefined && patch.fgColor !== null) {
    assertCellFillColorPatch(patch.fgColor, "fgColor");
  }

  if (patch.bgColor !== undefined && patch.bgColor !== null) {
    assertCellFillColorPatch(patch.bgColor, "bgColor");
  }
}

function assertCellBorderPatch(patch: CellBorderPatch): void {
  assertOptionalNullableBoolean(patch.diagonalUp, "diagonalUp");
  assertOptionalNullableBoolean(patch.diagonalDown, "diagonalDown");
  assertOptionalNullableBoolean(patch.outline, "outline");

  assertCellBorderSidePatch(patch.left, "left");
  assertCellBorderSidePatch(patch.right, "right");
  assertCellBorderSidePatch(patch.top, "top");
  assertCellBorderSidePatch(patch.bottom, "bottom");
  assertCellBorderSidePatch(patch.diagonal, "diagonal");
  assertCellBorderSidePatch(patch.vertical, "vertical");
  assertCellBorderSidePatch(patch.horizontal, "horizontal");
}

function assertFormatCode(formatCode: string): void {
  if (formatCode.length === 0) {
    throw new XlsxError("Invalid format code: empty");
  }
}

function assertCellFontColorPatch(patch: CellFontColorPatch): void {
  assertOptionalNullableString(patch.rgb, "color.rgb");
  assertOptionalNullableNonNegativeInteger(patch.theme, "color.theme");
  assertOptionalNullableNonNegativeInteger(patch.indexed, "color.indexed");
  assertOptionalNullableBoolean(patch.auto, "color.auto");
  assertOptionalNullableFiniteNumber(patch.tint, "color.tint");
}

function assertCellFillColorPatch(patch: CellFillColorPatch, name: string): void {
  assertOptionalNullableString(patch.rgb, `${name}.rgb`);
  assertOptionalNullableNonNegativeInteger(patch.theme, `${name}.theme`);
  assertOptionalNullableNonNegativeInteger(patch.indexed, `${name}.indexed`);
  assertOptionalNullableBoolean(patch.auto, `${name}.auto`);
  assertOptionalNullableFiniteNumber(patch.tint, `${name}.tint`);
}

function assertCellBorderSidePatch(patch: CellBorderSidePatch | null | undefined, name: string): void {
  if (patch === undefined || patch === null) {
    return;
  }

  assertOptionalNullableString(patch.style, `${name}.style`);
  if (patch.color !== undefined && patch.color !== null) {
    assertCellBorderColorPatch(patch.color, `${name}.color`);
  }
}

function assertCellBorderColorPatch(patch: CellBorderColorPatch, name: string): void {
  assertOptionalNullableString(patch.rgb, `${name}.rgb`);
  assertOptionalNullableNonNegativeInteger(patch.theme, `${name}.theme`);
  assertOptionalNullableNonNegativeInteger(patch.indexed, `${name}.indexed`);
  assertOptionalNullableBoolean(patch.auto, `${name}.auto`);
  assertOptionalNullableFiniteNumber(patch.tint, `${name}.tint`);
}

function assertCellStyleAlignmentPatch(patch: CellStyleAlignmentPatch): void {
  assertOptionalNullableString(patch.horizontal, "alignment.horizontal");
  assertOptionalNullableString(patch.vertical, "alignment.vertical");
  assertOptionalNullableNonNegativeInteger(patch.textRotation, "alignment.textRotation");
  assertOptionalNullableBoolean(patch.wrapText, "alignment.wrapText");
  assertOptionalNullableBoolean(patch.shrinkToFit, "alignment.shrinkToFit");
  assertOptionalNullableNonNegativeInteger(patch.indent, "alignment.indent");
  assertOptionalNullableNonNegativeInteger(patch.relativeIndent, "alignment.relativeIndent");
  assertOptionalNullableBoolean(patch.justifyLastLine, "alignment.justifyLastLine");
  assertOptionalNullableNonNegativeInteger(patch.readingOrder, "alignment.readingOrder");
}

function assertOptionalNonNegativeInteger(value: number | undefined, name: string): void {
  if (value !== undefined && (!Number.isInteger(value) || value < 0)) {
    throw new XlsxError(`Invalid ${name}: ${value}`);
  }
}

function assertOptionalNullableNonNegativeInteger(value: number | null | undefined, name: string): void {
  if (value !== undefined && value !== null && (!Number.isInteger(value) || value < 0)) {
    throw new XlsxError(`Invalid ${name}: ${value}`);
  }
}

function assertOptionalNullableFiniteNumber(value: number | null | undefined, name: string): void {
  if (value !== undefined && value !== null && !Number.isFinite(value)) {
    throw new XlsxError(`Invalid ${name}: ${value}`);
  }
}

function assertOptionalNullableBoolean(value: boolean | null | undefined, name: string): void {
  if (value !== undefined && value !== null && typeof value !== "boolean") {
    throw new XlsxError(`Invalid ${name}: ${String(value)}`);
  }
}

function assertOptionalNullableString(value: string | null | undefined, name: string): void {
  if (value !== undefined && value !== null && typeof value !== "string") {
    throw new XlsxError(`Invalid ${name}: ${String(value)}`);
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

function assertSheetIndex(sheetIndex: number, sheetCount: number): void {
  if (!Number.isInteger(sheetIndex) || sheetIndex < 0 || sheetIndex >= sheetCount) {
    throw new XlsxError(`Invalid sheet index: ${sheetIndex}`);
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

  for (const match of workbookXml.matchAll(/<sheet\b[^>]*\bsheetId\s*=\s*["'](\d+)["']/g)) {
    nextSheetId = Math.max(nextSheetId, Number(match[1]) + 1);
  }

  return nextSheetId;
}

function getNextRelationshipId(relationshipsXml: string): string {
  let nextId = 1;

  for (const match of relationshipsXml.matchAll(/\bId\s*=\s*["']rId(\d+)["']/g)) {
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

function reorderWorkbookXmlSheets(
  workbookXml: string,
  currentSheets: Sheet[],
  nextSheets: Sheet[],
): string {
  const sheetsTag = findFirstXmlTag(workbookXml, "sheets");
  if (!sheetsTag || sheetsTag.innerXml === null) {
    throw new XlsxError("Workbook is missing <sheets>");
  }

  const sheetNodes = new Map<string, string>();
  for (const sheetTag of findXmlTags(sheetsTag.innerXml, "sheet")) {
    const relationshipId = getTagAttr(sheetTag, "r:id");
    if (relationshipId) {
      sheetNodes.set(relationshipId, sheetTag.source);
    }
  }

  const reorderedSheetsXml = nextSheets
    .map((sheet) => {
      const sheetXml = sheetNodes.get(sheet.relationshipId);
      if (!sheetXml) {
        throw new XlsxError(`Sheet relationship not found: ${sheet.relationshipId}`);
      }

      return sheetXml;
    })
    .join("");
  const localSheetIdMap = buildLocalSheetIdMap(currentSheets, nextSheets);

  return workbookXml
    .replace(/<sheets>[\s\S]*?<\/sheets>/, `<sheets>${reorderedSheetsXml}</sheets>`)
    .replace(
      /<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g,
      (match, attributesSource, nameSource) => {
        const attributes = parseAttributes(attributesSource);
        const localSheetIdIndex = attributes.findIndex(([name]) => name === "localSheetId");
        if (localSheetIdIndex === -1) {
          return match;
        }

        const localSheetIdText = attributes[localSheetIdIndex]?.[1];
        if (localSheetIdText === undefined) {
          return match;
        }

        const nextLocalSheetId = localSheetIdMap.get(Number(localSheetIdText));
        if (nextLocalSheetId === undefined) {
          return match;
        }

        attributes[localSheetIdIndex] = ["localSheetId", String(nextLocalSheetId)];
        const serializedAttributes = serializeAttributes(attributes);
        return `<definedName${serializedAttributes ? ` ${serializedAttributes}` : ""}>${nameSource}</definedName>`;
      },
    )
    .replace(
      /<workbookView\b([^>]*?)\/>/g,
      (match, attributesSource) => {
        const attributes = parseAttributes(attributesSource);
        const activeTabIndex = attributes.findIndex(([name]) => name === "activeTab");
        if (activeTabIndex === -1) {
          return match;
        }

        const activeTabText = attributes[activeTabIndex]?.[1];
        if (activeTabText === undefined) {
          return match;
        }

        const nextActiveTab = localSheetIdMap.get(Number(activeTabText));
        if (nextActiveTab === undefined) {
          return match;
        }

        attributes[activeTabIndex] = ["activeTab", String(nextActiveTab)];
        const serializedAttributes = serializeAttributes(attributes);
        return `<workbookView${serializedAttributes ? ` ${serializedAttributes}` : ""}/>`;
      },
    );
}

function buildLocalSheetIdMap(currentSheets: Sheet[], nextSheets: Sheet[]): Map<number, number> {
  const nextIndexesByRelationshipId = new Map<string, number>();
  nextSheets.forEach((sheet, index) => {
    nextIndexesByRelationshipId.set(sheet.relationshipId, index);
  });

  const localSheetIdMap = new Map<number, number>();
  currentSheets.forEach((sheet, index) => {
    const nextIndex = nextIndexesByRelationshipId.get(sheet.relationshipId);
    if (nextIndex !== undefined) {
      localSheetIdMap.set(index, nextIndex);
    }
  });

  return localSheetIdMap;
}

function parseSheetVisibility(workbookXml: string, relationshipId: string): SheetVisibility {
  for (const sheetTag of findXmlTags(workbookXml, "sheet")) {
    if (getTagAttr(sheetTag, "r:id") !== relationshipId) {
      continue;
    }

    const state = getTagAttr(sheetTag, "state");
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

function parseActiveSheetIndex(workbookXml: string, sheetCount: number): number {
  const workbookViewTag = findFirstXmlTag(workbookXml, "workbookView");
  const activeTabText = workbookViewTag ? getTagAttr(workbookViewTag, "activeTab") : undefined;
  const activeTab = activeTabText === undefined ? 0 : Number(activeTabText);

  if (!Number.isInteger(activeTab) || activeTab < 0 || activeTab >= sheetCount) {
    return 0;
  }

  return activeTab;
}

function updateActiveSheetInWorkbookXml(workbookXml: string, activeSheetIndex: number): string {
  const workbookViewsTag = findFirstXmlTag(workbookXml, "bookViews");
  const workbookViewXml = `<workbookView activeTab="${activeSheetIndex}"/>`;

  if (!workbookViewsTag) {
    return workbookXml.replace(/<sheets>/, `<bookViews>${workbookViewXml}</bookViews><sheets>`);
  }

  if (!findFirstXmlTag(workbookViewsTag.innerXml ?? "", "workbookView")) {
    return workbookXml.replace(/<bookViews>[\s\S]*?<\/bookViews>/, `<bookViews>${workbookViewXml}</bookViews>`);
  }

  return workbookXml.replace(
    /<workbookView\b([^>]*?)\/>/,
    (match, attributesSource) => {
      const attributes = parseAttributes(attributesSource);
      const activeTabIndex = attributes.findIndex(([name]) => name === "activeTab");
      if (activeTabIndex === -1) {
        attributes.push(["activeTab", String(activeSheetIndex)]);
      } else {
        attributes[activeTabIndex] = ["activeTab", String(activeSheetIndex)];
      }

      const serializedAttributes = serializeAttributes(attributes);
      return `<workbookView${serializedAttributes ? ` ${serializedAttributes}` : ""}/>`;
    },
  );
}

function parseDefinedNames(workbookXml: string, sheets: Sheet[]): DefinedName[] {
  return findXmlTags(workbookXml, "definedName")
    .filter((tag) => tag.innerXml !== null)
    .map((tag) => {
      const localSheetIdText = getTagAttr(tag, "localSheetId");
      const localSheetId = localSheetIdText === undefined ? null : Number(localSheetIdText);
      return {
        hidden: getTagAttr(tag, "hidden") === "1",
        name: getTagAttr(tag, "name") ?? "",
        scope: localSheetId === null ? null : (sheets[localSheetId]?.name ?? null),
        value: decodeXmlText(tag.innerXml ?? ""),
      };
    })
    .filter((definedName) => definedName.name.length > 0);
}

function buildDefinedNameXml(name: string, value: string, localSheetId: number | null): string {
  const attributes: Array<[string, string]> = [["name", name]];
  if (localSheetId !== null) {
    attributes.push(["localSheetId", String(localSheetId)]);
  }

  return `<definedName ${serializeAttributes(attributes)}>${escapeXmlText(value)}</definedName>`;
}

function insertDefinedNameIntoWorkbookXml(workbookXml: string, definedNameXml: string): string {
  const definedNamesTag = findFirstXmlTag(workbookXml, "definedNames");
  if (definedNamesTag) {
    const insertionIndex = definedNamesTag.end - "</definedNames>".length;
    return workbookXml.slice(0, insertionIndex) + definedNameXml + workbookXml.slice(insertionIndex);
  }

  return insertBeforeClosingTag(workbookXml, "workbook", `<definedNames>${definedNameXml}</definedNames>`);
}

function removeDefinedNameFromWorkbookXml(
  workbookXml: string,
  name: string,
  localSheetId: number | null,
): string {
  const definedNamesTag = findFirstXmlTag(workbookXml, "definedNames");
  if (!definedNamesTag || definedNamesTag.innerXml === null) {
    return workbookXml;
  }

  const nextInnerXml = definedNamesTag.innerXml.replace(
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
    workbookXml.slice(0, definedNamesTag.start) +
    nextDefinedNamesXml +
    workbookXml.slice(definedNamesTag.end);

  return /<definedName\b/.test(nextInnerXml)
    ? nextWorkbookXml
    : workbookXml.slice(0, definedNamesTag.start) +
        workbookXml.slice(definedNamesTag.end);
}

function removeSheetFromWorkbookXml(
  workbookXml: string,
  relationshipId: string,
  deletedSheetName: string,
  deletedSheetIndex: number,
): string {
  const withoutSheet = workbookXml.replace(
    new RegExp(`<sheet\\b[^>]*\\br:id\\s*=\\s*["']${escapeRegex(relationshipId)}["'][^>]*/>`),
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
    new RegExp(`<Relationship\\b[^>]*\\bId\\s*=\\s*["']${escapeRegex(relationshipId)}["'][^>]*/>`),
    "",
  );
}

function removeContentTypeOverride(contentTypesXml: string, partPath: string): string {
  return contentTypesXml.replace(
    new RegExp(`<Override\\b[^>]*\\bPartName\\s*=\\s*["']/${escapeRegex(partPath)}["'][^>]*/>`),
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
