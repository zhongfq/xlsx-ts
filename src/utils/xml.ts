const ATTRIBUTE_REGEX = /([A-Za-z_][\w:.-]*)="([^"]*)"/g;

export function escapeXmlText(value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll("\"", "&quot;");
}

export function decodeXmlText(value: string): string {
  return value
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", "\"")
    .replaceAll("&apos;", "'")
    .replaceAll("&amp;", "&");
}

export function getXmlAttr(source: string, attributeName: string): string | undefined {
  const regex = new RegExp(`${escapeRegex(attributeName)}="([^"]*)"`);
  const match = source.match(regex);
  return match?.[1];
}

export function parseAttributes(source: string): Array<[string, string]> {
  const attributes: Array<[string, string]> = [];

  for (const match of source.matchAll(ATTRIBUTE_REGEX)) {
    attributes.push([match[1], match[2]]);
  }

  return attributes;
}

export function serializeAttributes(attributes: Array<[string, string]>): string {
  return attributes.map(([key, value]) => `${key}="${escapeXmlText(value)}"`).join(" ");
}

export function escapeRegex(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

export function extractTagText(xml: string, tagName: string): string | undefined {
  const match = xml.match(new RegExp(`<${tagName}\\b[^>]*>([\\s\\S]*?)</${tagName}>`));
  return match?.[1];
}

export function extractAllTagTexts(xml: string, tagName: string): string[] {
  return Array.from(
    xml.matchAll(new RegExp(`<${tagName}\\b[^>]*>([\\s\\S]*?)</${tagName}>`, "g")),
    (match) => match[1],
  );
}
