const ATTRIBUTE_REGEX = /([A-Za-z_][\w:.-]*)\s*=\s*(["'])([\s\S]*?)\2/g;
const XML_ENTITY_REGEX = /&(#x[0-9a-fA-F]+|#\d+|lt|gt|quot|apos|amp);/g;

export function escapeXmlText(value: string): string {
  if (
    !value.includes("&") &&
    !value.includes("<") &&
    !value.includes(">") &&
    !value.includes("\"")
  ) {
    return value;
  }

  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll("\"", "&quot;");
}

export function decodeXmlText(value: string): string {
  if (!value.includes("&")) {
    return value;
  }

  return value.replace(XML_ENTITY_REGEX, (entity, token: string) => {
    switch (token) {
      case "lt":
        return "<";
      case "gt":
        return ">";
      case "quot":
        return "\"";
      case "apos":
        return "'";
      case "amp":
        return "&";
      default:
        return decodeNumericXmlEntity(entity, token);
    }
  });
}

export function getXmlAttr(source: string, attributeName: string): string | undefined {
  const regex = new RegExp(`${escapeRegex(attributeName)}\\s*=\\s*(["'])([\\s\\S]*?)\\1`);
  const match = source.match(regex);
  return match ? decodeXmlText(match[2]) : undefined;
}

export function parseAttributes(source: string): Array<[string, string]> {
  const attributes: Array<[string, string]> = [];

  for (const match of source.matchAll(ATTRIBUTE_REGEX)) {
    attributes.push([match[1], decodeXmlText(match[3])]);
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

function decodeNumericXmlEntity(entity: string, token: string): string {
  const isHex = token.startsWith("#x");
  const digits = token.slice(isHex ? 2 : 1);
  const codePoint = Number.parseInt(digits, isHex ? 16 : 10);

  if (!Number.isInteger(codePoint) || codePoint < 0 || codePoint > 0x10ffff) {
    return entity;
  }

  try {
    return String.fromCodePoint(codePoint);
  } catch {
    return entity;
  }
}
