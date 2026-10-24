import { XmlParseError } from "./errors";

// XML 1.0 allows TAB, LF, CR, U+0020-D7FF, E000-FFFD, and paired U+10000-10FFFF.
// Native regex first rejects the common valid string in one scan.
const INVALID_XML_CHARACTERS =
  /[\p{Control}--[\t\n\r]]|[\uFFFE\uFFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/v;

export function isValidXmlCharacter(code: number): boolean {
  return (
    code === 0x09 ||
    code === 0x0a ||
    code === 0x0d ||
    (code >= 0x20 && code <= 0xd7ff) ||
    (code >= 0xe000 && code <= 0xfffd) ||
    (code >= 0x10000 && code <= 0x10ffff)
  );
}

export function assertValidXmlCharacters(source: string): void {
  if (!INVALID_XML_CHARACTERS.test(source)) return;
  for (let index = 0; index < source.length; index += 1) {
    const code = source.charCodeAt(index);
    if (code >= 0xd800 && code <= 0xdbff) {
      const low = source.charCodeAt(index + 1);
      const combined =
        low >= 0xdc00 && low <= 0xdfff ? 0x10000 + ((code - 0xd800) << 10) + (low - 0xdc00) : 0;
      if (combined >= 0x10000 && combined <= 0x10ffff) {
        index += 1;
        continue;
      }
      throw invalidCharacter(code);
    }
    if (!isValidXmlCharacter(code)) throw invalidCharacter(code);
  }
}

function invalidCharacter(code: number): XmlParseError {
  return new XmlParseError(
    `Malformed XML: invalid character U+${code.toString(16).padStart(4, "0")}.`,
  );
}
