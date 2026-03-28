import { decodeXmlText, extractAllTagTexts } from "./utils/xml.js";

export function parseSharedStrings(xml: string): string[] {
  return Array.from(xml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g), (match) =>
    extractAllTagTexts(match[1], "t").map(decodeXmlText).join(""),
  );
}
