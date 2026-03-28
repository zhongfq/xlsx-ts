import { Sheet } from "./sheet.js";
import type { Workbook } from "./workbook.js";
import { basenamePosix, dirnamePosix, resolvePosix } from "./utils/path.js";
import { getXmlAttr } from "./utils/xml.js";

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
