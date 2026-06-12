/**
 * Notes Master (p:notesMaster) descriptor for PPTX.
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, attrNum, findChild } from "@office-open/xml";
import { buildNotesMasterXml } from "@parts/notes-master";
import type {
  NotesMasterOptions,
  NotesLevelProperties,
  ColorMapOptions,
} from "@parts/notes-master";

// ── Types ──

export interface NotesMasterDescriptorOptions {
  options?: NotesMasterOptions;
}

const COLOR_MAP_KEYS: (keyof ColorMapOptions)[] = [
  "bg1",
  "tx1",
  "bg2",
  "tx2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "hlink",
  "folHlink",
];

const LEVEL_TAGS = [
  "a:lvl1pPr",
  "a:lvl2pPr",
  "a:lvl3pPr",
  "a:lvl4pPr",
  "a:lvl5pPr",
  "a:lvl6pPr",
  "a:lvl7pPr",
  "a:lvl8pPr",
  "a:lvl9pPr",
];

// ── Descriptor ──

export const notesMasterDesc: CustomDescriptor<NotesMasterDescriptorOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    return buildNotesMasterXml(opts.options);
  },

  parse(el, _ctx) {
    const result: Record<string, unknown> = {};
    const options: Record<string, unknown> = {};

    // colorMap
    const clrMap = findChild(el, "p:clrMap");
    if (clrMap) {
      const colorMap: Record<string, string> = {};
      for (const key of COLOR_MAP_KEYS) {
        const v = attr(clrMap, key);
        if (v !== undefined) colorMap[key] = v;
      }
      if (Object.keys(colorMap).length > 0) options.colorMap = colorMap;
    }

    // headerFooter
    const hf = findChild(el, "p:hf");
    if (hf) {
      const headerFooter: Record<string, boolean> = {};
      const dt = attr(hf, "dt");
      if (dt !== undefined) headerFooter.date = dt === "1";
      const hdr = attr(hf, "hdr");
      if (hdr !== undefined) headerFooter.header = hdr === "1";
      const ftr = attr(hf, "ftr");
      if (ftr !== undefined) headerFooter.footer = ftr === "1";
      const sldNum = attr(hf, "sldNum");
      if (sldNum !== undefined) headerFooter.slideNumber = sldNum === "1";
      if (Object.keys(headerFooter).length > 0) options.headerFooter = headerFooter;
    }

    // notesStyle
    const notesStyle = findChild(el, "p:notesStyle");
    if (notesStyle) {
      const levels: NotesLevelProperties[] = [];
      for (const tag of LEVEL_TAGS) {
        const lvlEl = findChild(notesStyle, tag);
        if (lvlEl) {
          const lvl: Record<string, unknown> = {};
          const defRPr = findChild(lvlEl, "a:defRPr");
          if (defRPr) {
            const sz = attrNum(defRPr, "sz");
            if (sz !== undefined) lvl.fontSize = sz;
          }
          const marL = attrNum(lvlEl, "marL");
          if (marL !== undefined) lvl.marginLeft = marL;
          const algn = attr(lvlEl, "algn");
          if (algn !== undefined) lvl.alignment = algn;
          levels.push(lvl as NotesLevelProperties);
        }
      }
      if (levels.length > 0) options.notesStyle = levels;
    }

    if (Object.keys(options).length > 0) result.options = options;
    return result as unknown as NotesMasterDescriptorOptions;
  },
};
