import { escapeXml } from "@office-open/xml";

/** Escapes XML 1.0 control characters using the SpreadsheetML string form. */
export function escapeSpreadsheetXml(value: string): string {
  const escapedLiterals = value.replace(/_x[0-9a-f]{4}_/giu, (match) => `_x005F_${match.slice(1)}`);
  let escapedControls = "";
  for (let index = 0; index < escapedLiterals.length; index++) {
    const character = escapedLiterals[index]!;
    const code = character.charCodeAt(0);
    escapedControls +=
      code < 0x20 && code !== 0x09 && code !== 0x0a && code !== 0x0d
        ? `_x${code.toString(16).padStart(4, "0").toUpperCase()}_`
        : character;
  }
  return escapeXml(escapedControls);
}

/** Decodes SpreadsheetML character escapes after XML parsing. */
export function decodeSpreadsheetXml(value: string): string {
  return value.replace(
    /_x([0-9a-f]{4})_(x[0-9a-f]{4}_)?/giu,
    (_match, hex: string, escaped: string | undefined) =>
      Number.parseInt(hex, 16) === 0x005f && escaped !== undefined
        ? `_${escaped}`
        : String.fromCharCode(Number.parseInt(hex, 16)) + (escaped ?? ""),
  );
}
