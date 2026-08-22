/**
 * Font scheme (a:fontScheme / CT_FontScheme) stringify + parse.
 *
 * CT_FontCollection order: latin → ea → cs → font[] (supplemental).
 *
 * @module
 */
import { escapeXml } from "@office-open/xml";
import { findChild } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

import type {
  FontCollectionOptions,
  FontSchemeOptions,
  SupplementalFontOptions,
  TextFontOptions,
} from "./theme-options";

// ── TextFont (a:latin / a:ea / a:cs) ──

function stringifyTextFont(tag: string, opts: TextFontOptions | undefined): string {
  if (!opts) return "";
  const attrs = [`typeface="${escapeXml(opts.typeface)}"`];
  if (opts.panose !== undefined) attrs.push(`panose="${escapeXml(opts.panose)}"`);
  if (opts.pitchFamily !== undefined) attrs.push(`pitchFamily="${opts.pitchFamily}"`);
  if (opts.charset !== undefined) attrs.push(`charset="${opts.charset}"`);
  return `<${tag} ${attrs.join(" ")}/>`;
}

function parseTextFont(el: XmlElement | undefined): TextFontOptions | undefined {
  if (!el?.attributes) return undefined;
  const typeface = el.attributes["typeface"];
  if (typeface === undefined) return undefined;
  const result: TextFontOptions = { typeface: String(typeface) };
  if (el.attributes["panose"] !== undefined) result.panose = String(el.attributes["panose"]);
  if (el.attributes["pitchFamily"] !== undefined)
    result.pitchFamily = Number(el.attributes["pitchFamily"]);
  if (el.attributes["charset"] !== undefined) result.charset = Number(el.attributes["charset"]);
  return result;
}

// ── Font collection (a:majorFont / a:minorFont) ──

function stringifyFontCollection(tag: string, opts: FontCollectionOptions | undefined): string {
  const latin = stringifyTextFont("a:latin", opts?.latin);
  const eastAsian = stringifyTextFont("a:ea", opts?.eastAsian);
  const complexScript = stringifyTextFont("a:cs", opts?.complexScript);
  const supplemental = (opts?.supplementalFonts ?? [])
    .map((f) => `<a:font script="${escapeXml(f.script)}" typeface="${escapeXml(f.typeface)}"/>`)
    .join("");
  return `<${tag}>${latin}${eastAsian}${complexScript}${supplemental}</${tag}>`;
}

function parseFontCollection(el: XmlElement | undefined): FontCollectionOptions | undefined {
  if (!el) return undefined;
  const result: FontCollectionOptions = {};
  const latin = parseTextFont(findChild(el, "a:latin"));
  if (latin) result.latin = latin;
  const eastAsian = parseTextFont(findChild(el, "a:ea"));
  if (eastAsian) result.eastAsian = eastAsian;
  const complexScript = parseTextFont(findChild(el, "a:cs"));
  if (complexScript) result.complexScript = complexScript;
  const supplementalFonts = (el.elements ?? [])
    .filter((c) => c.name === "a:font")
    .map((c): SupplementalFontOptions | undefined => {
      const script = c.attributes?.["script"];
      const typeface = c.attributes?.["typeface"];
      if (script === undefined || typeface === undefined) return undefined;
      return { script: String(script), typeface: String(typeface) };
    })
    .filter((v): v is SupplementalFontOptions => v !== undefined);
  if (supplementalFonts.length > 0) result.supplementalFonts = supplementalFonts;
  return Object.keys(result).length > 0 ? result : undefined;
}

// ── Fresh defaults (Office theme major/minor fonts) ──

export const DEFAULT_MAJOR_FONT: FontCollectionOptions = {
  latin: { typeface: "Calibri Light", panose: "020F0302020204030204" },
  eastAsian: { typeface: "" },
  complexScript: { typeface: "" },
};

export const DEFAULT_MINOR_FONT: FontCollectionOptions = {
  latin: { typeface: "Calibri", panose: "020F0502020204030204" },
  eastAsian: { typeface: "" },
  complexScript: { typeface: "" },
};

// ── Font scheme (a:fontScheme) ──

/** Serialize a:fontScheme. Undefined options emit the Office default scheme. */
export function stringifyFontScheme(
  opts: FontSchemeOptions | undefined,
  fallbackName: string,
): string {
  // Merge defaults so the XSD-required latin/ea/cs slots are always present,
  // even when the caller only customizes one font (e.g. latin typeface).
  const majorFont = stringifyFontCollection(
    "a:majorFont",
    opts?.majorFont ? { ...DEFAULT_MAJOR_FONT, ...opts.majorFont } : DEFAULT_MAJOR_FONT,
  );
  const minorFont = stringifyFontCollection(
    "a:minorFont",
    opts?.minorFont ? { ...DEFAULT_MINOR_FONT, ...opts.minorFont } : DEFAULT_MINOR_FONT,
  );
  const name = opts?.name ?? fallbackName;
  return `<a:fontScheme name="${name}">${majorFont}${minorFont}</a:fontScheme>`;
}

/** Parse a:fontScheme. Returns undefined only when neither major nor minor font is present. */
export function parseFontScheme(el: XmlElement | undefined): FontSchemeOptions | undefined {
  if (!el) return undefined;
  const result: FontSchemeOptions = {};
  const name = el.attributes?.["name"];
  if (name) result.name = String(name);
  const majorFont = parseFontCollection(findChild(el, "a:majorFont"));
  if (majorFont) result.majorFont = majorFont;
  const minorFont = parseFontCollection(findChild(el, "a:minorFont"));
  if (minorFont) result.minorFont = minorFont;
  return Object.keys(result).length > 0 ? result : undefined;
}
