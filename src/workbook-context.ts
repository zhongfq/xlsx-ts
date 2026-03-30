import { Sheet } from "./sheet.js";
import type { Workbook } from "./workbook.js";
import { basenamePosix, dirnamePosix, resolveRelationshipTarget } from "./utils/path.js";
import { findXmlTags, getTagAttr } from "./utils/xml-read.js";

export interface WorkbookContext {
  workbookDir: string;
  workbookPath: string;
  workbookRelsPath: string;
  sharedStringsPath?: string;
  stylesPath?: string;
  sheets: Sheet[];
}

export function resolveWorkbookContext(
  workbook: Workbook,
  readEntryText: (path: string) => string,
): WorkbookContext {
  const rootRels = readEntryText("_rels/.rels");
  const workbookTarget = findRelationshipTarget(rootRels, /\/officeDocument$/) ?? "xl/workbook.xml";
  const workbookPath = workbookTarget.replace(/^\/+/, "");
  const workbookDir = dirnamePosix(workbookPath);
  const workbookRelsPath = `${workbookDir}/_rels/${basenamePosix(workbookPath)}.rels`;
  const workbookXml = readEntryText(workbookPath);
  const workbookRelsXml = readEntryText(workbookRelsPath);
  const relationships = parseRelationships(workbookRelsXml, workbookDir);
  const sheets = parseSheets(workbook, workbookXml, relationships);
  const sharedStringsPath = findRelationshipTarget(workbookRelsXml, /\/sharedStrings$/, workbookDir);
  const stylesPath = findRelationshipTarget(workbookRelsXml, /\/styles$/, workbookDir);

  return {
    workbookDir,
    workbookPath,
    workbookRelsPath,
    sharedStringsPath,
    stylesPath,
    sheets,
  };
}

function parseRelationships(xml: string, baseDir: string): Map<string, string> {
  const relationships = new Map<string, string>();

  for (const relationshipTag of findXmlTags(xml, "Relationship")) {
    if (!relationshipTag.selfClosing) {
      continue;
    }

    const id = getTagAttr(relationshipTag, "Id");
    const target = getTagAttr(relationshipTag, "Target");

    if (!id || !target) {
      continue;
    }

    relationships.set(id, resolveRelationshipTarget(baseDir, target));
  }

  return relationships;
}

function parseSheets(
  workbook: Workbook,
  workbookXml: string,
  relationships: Map<string, string>,
): Sheet[] {
  const sheets: Sheet[] = [];

  for (const sheetTag of findXmlTags(workbookXml, "sheet")) {
    const name = getTagAttr(sheetTag, "name");
    const relationshipId = getTagAttr(sheetTag, "r:id");

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
  for (const relationshipTag of findXmlTags(xml, "Relationship")) {
    if (!relationshipTag.selfClosing) {
      continue;
    }

    const type = getTagAttr(relationshipTag, "Type");
    const target = getTagAttr(relationshipTag, "Target");

    if (!type || !target || !typePattern.test(type)) {
      continue;
    }

    return baseDir ? resolveRelationshipTarget(baseDir, target) : target.replace(/^\/+/, "");
  }

  return undefined;
}
