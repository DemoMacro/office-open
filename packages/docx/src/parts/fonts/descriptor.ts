/**
 * Font table descriptor — produces word/fontTable.xml.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-w_fonts.html, CT_FontsList
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, escapeXml, findChild } from "@office-open/xml";
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
const DEFAULT_SIG = {
  usb0: "E0002AFF",
  usb1: "C000247B",
  usb2: "00000009",
  usb3: "00000000",
  csb0: "000001FF",
  csb1: "00000000",
};

function fontXml(font: EmbeddedFontOptionsWithKey, index: number): string {
  const parts: string[] = [`<w:font w:name="${escapeXml(font.name)}">`];

  // Default charset, family, pitch, sig for regular embedded font
  if (font.characterSet) {
    parts.push(`<w:charset w:val="${escapeXml(font.characterSet)}"/>`);
  }
  parts.push(`<w:family w:val="auto"/>`);
  parts.push(`<w:pitch w:val="variable"/>`);
  parts.push(
    `<w:sig w:usb0="${DEFAULT_SIG.usb0}" w:usb1="${DEFAULT_SIG.usb1}" ` +
      `w:usb2="${DEFAULT_SIG.usb2}" w:usb3="${DEFAULT_SIG.usb3}" ` +
      `w:csb0="${DEFAULT_SIG.csb0}" w:csb1="${DEFAULT_SIG.csb1}"/>`,
  );
  // Embedded regular font
  parts.push(`<w:embedRegular r:id="rId${index + 1}" w:fontKey="{${font.fontKey}}"/>`);

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
    for (let i = 0; i < opts.fonts.length; i++) {
      parts.push(fontXml(opts.fonts[i], i));
    }
    parts.push("</w:fonts>");
    return parts.join("");
  },

  parse(el, _ctx) {
    const fonts: Record<string, unknown>[] = [];
    for (const child of el.elements ?? []) {
      if (child.name !== "w:font") continue;
      const font: Record<string, unknown> = {};
      const name = child.attributes?.["w:name"];
      if (name) font.name = String(name);

      // charset → w:charset/@w:val
      const charsetEl = findChild(child, "w:charset");
      if (charsetEl) {
        const val = attr(charsetEl, "w:val");
        if (val) font.characterSet = val;
      }

      // family → w:family/@w:val
      const familyEl = findChild(child, "w:family");
      if (familyEl) {
        const val = attr(familyEl, "w:val");
        if (val) font.family = val;
      }

      // pitch → w:pitch/@w:val
      const pitchEl = findChild(child, "w:pitch");
      if (pitchEl) {
        const val = attr(pitchEl, "w:val");
        if (val) font.pitch = val;
      }

      // sig → w:sig attributes
      const sigEl = findChild(child, "w:sig");
      if (sigEl) {
        const sig: Record<string, string> = {};
        for (const key of ["usb0", "usb1", "usb2", "usb3", "csb0", "csb1"] as const) {
          const val = attr(sigEl, `w:${key}`);
          if (val) sig[key] = val;
        }
        if (Object.keys(sig).length > 0) font.sig = sig;
      }

      // fontKey → w:embedRegular/@w:fontKey (strip curly braces)
      for (const embedTag of [
        "w:embedRegular",
        "w:embedBold",
        "w:embedItalic",
        "w:embedBoldItalic",
      ] as const) {
        const embedEl = findChild(child, embedTag);
        if (embedEl) {
          const rawKey = attr(embedEl, "w:fontKey");
          if (rawKey) {
            // Strip wrapping curly braces: {{...}} → ...
            font.fontKey = rawKey.replace(/^\{\{|\}\}$/g, "").replace(/^\{|\}$/g, "");
          }
          break; // Use first embed element found
        }
      }

      fonts.push(font);
    }
    return { fonts } as unknown as FontTableInput;
  },
};
