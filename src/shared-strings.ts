import { decodeXmlText, extractAllTagTexts } from "./utils/xml.js";

export function parseSharedStrings(xml: string): string[] {
  return Array.from(xml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g), (match) => parseStringItemText(match[1]));
}

export function parseStringItemText(xml: string): string {
  const visibleXml = xml.replace(/<rPh\b[^>]*>[\s\S]*?<\/rPh>/g, "");
  return extractAllTagTexts(visibleXml, "t").map(decodeXmlText).join("");
}
