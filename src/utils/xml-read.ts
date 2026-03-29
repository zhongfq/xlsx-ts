import { getXmlAttr } from "./xml.js";

export interface XmlTag {
  attributesSource: string;
  end: number;
  innerXml: string | null;
  selfClosing: boolean;
  source: string;
  start: number;
  tagName: string;
}

export function findXmlTags(xml: string, tagName: string): XmlTag[] {
  const tags: XmlTag[] = [];
  const openingPattern = `<${tagName}`;
  let searchStart = 0;

  while (searchStart < xml.length) {
    const tagStart = xml.indexOf(openingPattern, searchStart);
    if (tagStart === -1) {
      break;
    }

    if (!isTagBoundary(xml, tagStart + openingPattern.length)) {
      searchStart = tagStart + openingPattern.length;
      continue;
    }

    const tagOpenEnd = findTagOpenEnd(xml, tagStart + 1);
    if (tagOpenEnd === -1) {
      break;
    }

    const tagOpenSource = xml.slice(tagStart + 1 + tagName.length, tagOpenEnd);
    const selfClosing = isSelfClosingTagSource(tagOpenSource);

    if (selfClosing) {
      tags.push({
        attributesSource: trimTagAttributesSource(tagOpenSource),
        end: tagOpenEnd + 1,
        innerXml: null,
        selfClosing: true,
        source: xml.slice(tagStart, tagOpenEnd + 1),
        start: tagStart,
        tagName,
      });
      searchStart = tagOpenEnd + 1;
      continue;
    }

    const closingPattern = `</${tagName}>`;
    const closeStart = xml.indexOf(closingPattern, tagOpenEnd + 1);
    if (closeStart === -1) {
      break;
    }

    tags.push({
      attributesSource: trimTagAttributesSource(tagOpenSource),
      end: closeStart + closingPattern.length,
      innerXml: xml.slice(tagOpenEnd + 1, closeStart),
      selfClosing: false,
      source: xml.slice(tagStart, closeStart + closingPattern.length),
      start: tagStart,
      tagName,
    });
    searchStart = closeStart + closingPattern.length;
  }

  return tags;
}

export function findFirstXmlTag(xml: string, tagName: string): XmlTag | null {
  return findXmlTags(xml, tagName)[0] ?? null;
}

export function getTagAttr(tag: XmlTag, attributeName: string): string | undefined {
  return getXmlAttr(tag.attributesSource, attributeName);
}

function findTagOpenEnd(xml: string, start: number): number {
  let quote: number | null = null;

  for (let index = start; index < xml.length; index += 1) {
    const code = xml.charCodeAt(index);

    if (quote !== null) {
      if (code === quote) {
        quote = null;
      }
      continue;
    }

    if (code === 34 || code === 39) {
      quote = code;
      continue;
    }

    if (code === 62) {
      return index;
    }
  }

  return -1;
}

function isTagBoundary(xml: string, index: number): boolean {
  if (index >= xml.length) {
    return true;
  }

  const code = xml.charCodeAt(index);
  return code === 47 || code === 62 || isXmlWhitespaceCode(code);
}

function trimTagAttributesSource(source: string): string {
  let end = source.length;
  while (end > 0 && isXmlWhitespaceCode(source.charCodeAt(end - 1))) {
    end -= 1;
  }

  if (end > 0 && source.charCodeAt(end - 1) === 47) {
    end -= 1;
    while (end > 0 && isXmlWhitespaceCode(source.charCodeAt(end - 1))) {
      end -= 1;
    }
  }

  let start = 0;
  while (start < end && isXmlWhitespaceCode(source.charCodeAt(start))) {
    start += 1;
  }

  return source.slice(start, end);
}

function isSelfClosingTagSource(source: string): boolean {
  let index = source.length - 1;

  while (index >= 0 && isXmlWhitespaceCode(source.charCodeAt(index))) {
    index -= 1;
  }

  return index >= 0 && source.charCodeAt(index) === 47;
}

function isXmlWhitespaceCode(code: number): boolean {
  return code === 9 || code === 10 || code === 13 || code === 32;
}
