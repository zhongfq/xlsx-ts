import { findXmlTags } from "./utils/xml-read.js";
import { decodeXmlText } from "./utils/xml.js";

export function parseSharedStrings(xml: string): string[] {
  return findXmlTags(xml, "si")
    .filter((tag) => tag.innerXml !== null)
    .map((tag) => parseStringItemText(tag.innerXml ?? ""));
}

export function parseStringItemText(xml: string): string {
  const hiddenTags = findXmlTags(xml, "rPh");
  const visibleXml = [...hiddenTags]
    .sort((left, right) => right.start - left.start)
    .reduce((currentXml, tag) => currentXml.slice(0, tag.start) + currentXml.slice(tag.end), xml);

  return findXmlTags(visibleXml, "t")
    .map((tag) => decodeXmlText(tag.innerXml ?? ""))
    .join("");
}
