/**
 * Font table descriptor — produces word/fontTable.xml.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-w_fonts.html, CT_FontsList
 *
 * @module
 */

import { uCharHexNumber } from "@office-open/core";
import type { CustomDescriptor, WriteContext } from "@office-open/core/descriptor";
import { attr, attrBool, escapeXml, findChild } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import {
  DocumentAttributeNamespaces,
  documentNamespaceAttributes,
} from "@parts/document/document-attributes";
import type { EmbeddedFontPartName, FontSignature } from "@parts/fonts/font-table";
import {
  embeddedFontElementName,
  type FontTableFont,
  type EmbeddedFontPartWithKey,
} from "@parts/fonts/font-wrapper";

// ── Input ──

export interface FontTableInput {
  fonts: FontTableFont[];
}

// ── XML helpers ──

const NS = documentNamespaceAttributes([
  "mc",
  "r",
  "w",
  "w14",
  "w15",
  "w16",
  "w16cex",
  "w16cid",
  "w16sdtdh",
  "w16se",
]);

/** Default font signature for regular embedded fonts. */
const DEFAULT_SIG: FontSignature = {
  usb0: "E0002AFF",
  usb1: "C000247B",
  usb2: "00000009",
  usb3: "00000000",
  csb0: "000001FF",
  csb1: "00000000",
};

const EMBED_FONT_ELEMENTS: readonly EmbeddedFontPartName[] = [
  "embedRegular",
  "embedBold",
  "embedItalic",
  "embedBoldItalic",
];

function embeddedFontXml(partName: EmbeddedFontPartName, part: EmbeddedFontPartWithKey): string {
  if (part.embedRid === undefined) return "";
  // ST_Guid is uppercase in canonical form — the SDK validator rejects a
  // lowercase fontKey even though ISO's lexical space allows both cases.
  const attrs = [`r:id="${part.embedRid}"`, `w:fontKey="{${(part.fontKey ?? "").toUpperCase()}}"`];
  if (part.subsetted !== undefined) attrs.push(`w:subsetted="${part.subsetted ? 1 : 0}"`);
  return `<${embeddedFontElementName(partName)} ${attrs.join(" ")}/>`;
}

function fontXml(font: FontTableFont): string {
  const parts: string[] = [`<w:font w:name="${escapeXml(font.nameRaw ?? font.name ?? "")}">`];

  // CT_Font child order (wml.xsd): altName, panose1, charset, family,
  // notTrueType, pitch, sig, embed*. Embedded fonts (carrying data) fall back
  // to defaults when the parsed value is absent; metadata-only fonts emit only
  // what was parsed.
  if (font.altName) parts.push(`<w:altName w:val="${escapeXml(font.altName)}"/>`);
  if (font.panose1) parts.push(`<w:panose1 w:val="${escapeXml(font.panose1)}"/>`);
  if (font.characterSet || font.characterSetRaw || font.characterSetName) {
    const valAttr = font.characterSetRaw
      ? ` w:val="${escapeXml(font.characterSetRaw)}"`
      : font.characterSet
        ? ` w:val="${escapeXml(font.characterSet)}"`
        : "";
    const csAttr = font.characterSetName
      ? ` w:characterSet="${escapeXml(font.characterSetName)}"`
      : "";
    parts.push(`<w:charset${valAttr}${csAttr}/>`);
  }
  const hasEmbeddedFace = EMBED_FONT_ELEMENTS.some((name) => font[name]?.embedRid !== undefined);
  const family = font.family ?? (hasEmbeddedFace ? "auto" : undefined);
  if (family) parts.push(`<w:family w:val="${escapeXml(family)}"/>`);
  if (font.notTrueType !== undefined) {
    parts.push(font.notTrueType ? "<w:notTrueType/>" : '<w:notTrueType w:val="0"/>');
  }
  const pitch = font.pitch ?? (hasEmbeddedFace ? "variable" : undefined);
  if (pitch) parts.push(`<w:pitch w:val="${escapeXml(pitch)}"/>`);
  const sig =
    font.sig === null ? undefined : (font.sig ?? (hasEmbeddedFace ? DEFAULT_SIG : undefined));
  if (sig) {
    parts.push(
      `<w:sig w:usb0="${sig.usb0}" w:usb1="${sig.usb1}" w:usb2="${sig.usb2}" ` +
        `w:usb3="${sig.usb3}" w:csb0="${sig.csb0}" w:csb1="${sig.csb1}"/>`,
    );
  }
  for (const partName of EMBED_FONT_ELEMENTS) {
    const part = font[partName];
    if (part) parts.push(embeddedFontXml(partName, part));
  }

  parts.push("</w:font>");
  return parts.join("");
}

// ── Descriptor ──

export const fontTableDesc: CustomDescriptor<
  FontTableInput,
  WriteContext,
  { fonts: FontTableFont[] }
> = {
  kind: "custom",

  stringify(opts, _ctx) {
    // Declare each Requires= prefix referenced by a wrapped font; the base NS
    // list stays fixed so ordinary tables keep their byte layout.
    const nsByPrefix = DocumentAttributeNamespaces as Record<string, string>;
    const requiredNs = [...new Set(opts.fonts.map((font) => font.requires).filter(Boolean))]
      .flatMap((prefix): string[] => {
        const uri = prefix === undefined ? undefined : nsByPrefix[prefix];
        return uri === undefined ? [] : [` xmlns:${prefix}="${uri}"`];
      })
      .join("");
    const parts: string[] = [
      `<w:fonts ${NS}${requiredNs} mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh">`,
    ];
    for (const font of opts.fonts) {
      const xml = fontXml(font);
      parts.push(
        font.requires
          ? `<mc:AlternateContent><mc:Choice Requires="${escapeXml(font.requires)}">${xml}</mc:Choice><mc:Fallback/></mc:AlternateContent>`
          : xml,
      );
    }
    parts.push("</w:fonts>");
    return parts.join("");
  },

  parse(el, _ctx) {
    const fonts: FontTableFont[] = [];
    const pushFont = (child: Element, requires?: string) => {
      const rawName = child.attributes?.["w:name"];
      const font: FontTableFont =
        rawName === undefined ? { name: "", nameRaw: "" } : { name: String(rawName) };
      if (requires) font.requires = requires;

      const altNameEl = findChild(child, "w:altName");
      if (altNameEl) {
        const val = attr(altNameEl, "w:val");
        if (val) font.altName = val;
      }

      const panose1El = findChild(child, "w:panose1");
      if (panose1El) {
        const val = attr(panose1El, "w:val");
        if (val) font.panose1 = val;
      }

      const charsetEl = findChild(child, "w:charset");
      if (charsetEl) {
        const val = attr(charsetEl, "w:val");
        if (val !== undefined) {
          try {
            font.characterSet = uCharHexNumber(val);
            if (font.characterSet !== val) font.characterSetRaw = val;
          } catch {
            font.characterSetRaw = val;
          }
        }
        const csName = attr(charsetEl, "w:characterSet");
        if (csName) font.characterSetName = csName;
      }

      const familyEl = findChild(child, "w:family");
      if (familyEl) {
        const val = attr(familyEl, "w:val");
        if (val) font.family = val;
      }

      // Bare element means true (CT_OnOff default); w:val="0" is the explicit false.
      const notTrueTypeEl = findChild(child, "w:notTrueType");
      if (notTrueTypeEl) font.notTrueType = attrBool(notTrueTypeEl, "w:val") ?? true;

      const pitchEl = findChild(child, "w:pitch");
      if (pitchEl) {
        const val = attr(pitchEl, "w:val");
        if (val) font.pitch = val;
      }

      const sigEl = findChild(child, "w:sig");
      if (sigEl) {
        const sig: FontSignature = {
          usb0: "",
          usb1: "",
          usb2: "",
          usb3: "",
          csb0: "",
          csb1: "",
        };
        for (const key of ["usb0", "usb1", "usb2", "usb3", "csb0", "csb1"] as const) {
          const val = attr(sigEl, `w:${key}`);
          if (val) sig[key] = val;
        }
        if (Object.values(sig).some(Boolean)) font.sig = sig;
      }

      for (const partName of EMBED_FONT_ELEMENTS) {
        const embedEl = findChild(child, embeddedFontElementName(partName));
        if (!embedEl) continue;
        const part: EmbeddedFontPartWithKey = {};
        const rawKey = attr(embedEl, "w:fontKey");
        if (rawKey) part.fontKey = rawKey.replace(/^\{\{|\}\}$/g, "").replace(/^\{|\}$/g, "");
        const rid = attr(embedEl, "r:id");
        if (rid) part.embedRid = rid;
        const subsetted = attrBool(embedEl, "w:subsetted");
        if (subsetted !== undefined) part.subsetted = subsetted;
        font[partName] = part;
      }
      // Mirror the stringify-side defaults so a round-trip of a minimal
      // embedded face (family/pitch/sig filled in by the writer) parses back
      // to the same model instead of gaining explicit fields.
      if (EMBED_FONT_ELEMENTS.some((name) => font[name]?.embedRid !== undefined)) {
        font.family ??= "auto";
        font.pitch ??= "variable";
        font.sig ??= null;
      }

      fonts.push(font);
    };
    for (const child of el.elements ?? []) {
      if (child.name === "w:font") {
        pushFont(child);
        continue;
      }
      // Word 2010+ can gate a font declaration behind a markup-compatibility
      // Choice (e.g. Requires="wpc"); the Fallback branch is empty in practice.
      if (child.name === "mc:AlternateContent") {
        const choice = findChild(child, "mc:Choice");
        const choiceFont = choice ? findChild(choice, "w:font") : undefined;
        if (choiceFont) pushFont(choiceFont, attr(choice, "Requires") ?? undefined);
      }
    }
    return { fonts };
  },
};
