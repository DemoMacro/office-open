/**
 * Font table descriptor — produces word/fontTable.xml.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-w_fonts.html, CT_FontsList
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, attrBool, escapeXml, findChild } from "@office-open/xml";
import type { FontSignature } from "@parts/fonts/font-table";
import type { EmbeddedFontOptionsWithKey } from "@parts/fonts/font-wrapper";

// ── Input ──

export interface FontTableInput {
  fonts: EmbeddedFontOptionsWithKey[];
}

// ── XML helpers ──

const NS =
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ' +
  'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" ' +
  'xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml" ' +
  'xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" ' +
  'xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" ' +
  'xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash" ' +
  'xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex"';

/** Default font signature for regular embedded fonts. */
const DEFAULT_SIG: FontSignature = {
  usb0: "E0002AFF",
  usb1: "C000247B",
  usb2: "00000009",
  usb3: "00000000",
  csb0: "000001FF",
  csb1: "00000000",
};

function fontXml(font: EmbeddedFontOptionsWithKey): string {
  const parts: string[] = [`<w:font w:name="${escapeXml(font.name)}">`];

  // CT_Font child order (wml.xsd): altName, panose1, charset, family,
  // notTrueType, pitch, sig, embed*. Embedded fonts (carrying data) fall back
  // to defaults when the parsed value is absent; metadata-only fonts emit only
  // what was parsed.
  if (font.altName) parts.push(`<w:altName w:val="${escapeXml(font.altName)}"/>`);
  if (font.panose1) parts.push(`<w:panose1 w:val="${escapeXml(font.panose1)}"/>`);
  if (font.characterSet) {
    parts.push(`<w:charset w:val="${escapeXml(font.characterSet)}"/>`);
  }
  const family = font.family ?? (font.embedRid ? "auto" : undefined);
  if (family) parts.push(`<w:family w:val="${escapeXml(family)}"/>`);
  const pitch = font.pitch ?? (font.embedRid ? "variable" : undefined);
  if (pitch) parts.push(`<w:pitch w:val="${escapeXml(pitch)}"/>`);
  const sig = font.sig ?? (font.embedRid ? DEFAULT_SIG : undefined);
  if (sig) {
    parts.push(
      `<w:sig w:usb0="${sig.usb0}" w:usb1="${sig.usb1}" w:usb2="${sig.usb2}" ` +
        `w:usb3="${sig.usb3}" w:csb0="${sig.csb0}" w:csb1="${sig.csb1}"/>`,
    );
  }
  if (font.embedRid) {
    const embedAttrs = [`r:id="${font.embedRid}"`, `w:fontKey="{${font.fontKey}}"`];
    if (font.subsetted !== undefined) embedAttrs.push(`w:subsetted="${font.subsetted ? 1 : 0}"`);
    parts.push(`<w:embedRegular ${embedAttrs.join(" ")}/>`);
  }

  parts.push("</w:font>");
  return parts.join("");
}

// ── Descriptor ──

export const fontTableDesc: CustomDescriptor<FontTableInput> = {
  kind: "custom",

  stringify(opts, _ctx) {
    const parts: string[] = [
      `<w:fonts ${NS} mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh">`,
    ];
    for (const font of opts.fonts) {
      parts.push(fontXml(font));
    }
    parts.push("</w:fonts>");
    return parts.join("");
  },

  parse(el, _ctx) {
    const fonts: EmbeddedFontOptionsWithKey[] = [];
    for (const child of el.elements ?? []) {
      if (child.name !== "w:font") continue;
      const font: Partial<EmbeddedFontOptionsWithKey> = { fontKey: "" };
      const name = child.attributes?.["w:name"];
      if (name) font.name = String(name);

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
        if (val) font.characterSet = val as EmbeddedFontOptionsWithKey["characterSet"];
      }

      const familyEl = findChild(child, "w:family");
      if (familyEl) {
        const val = attr(familyEl, "w:val");
        if (val) font.family = val;
      }

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

      // Obfuscation key + relationship id from embedRegular (the .odttf bytes
      // are resolved separately by parse.ts via fontTable.xml.rels).
      const embedEl = findChild(child, "w:embedRegular");
      if (embedEl) {
        const rawKey = attr(embedEl, "w:fontKey");
        if (rawKey) {
          font.fontKey = rawKey.replace(/^\{\{|\}\}$/g, "").replace(/^\{|\}$/g, "");
        }
        const rid = attr(embedEl, "r:id");
        if (rid) font.embedRid = rid;
        const subsetted = attrBool(embedEl, "w:subsetted");
        if (subsetted !== undefined) font.subsetted = subsetted;
      }

      fonts.push(font as EmbeddedFontOptionsWithKey);
    }
    return { fonts };
  },
};
