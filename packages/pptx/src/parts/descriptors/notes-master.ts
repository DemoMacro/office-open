/**
 * Notes Master (p:notesMaster) descriptor for PPTX.
 *
 * @module
 */

import { convertToPt, parseColorMapping } from "@office-open/core";
import type { UniversalMeasure } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, attrMeasure, findChild } from "@office-open/xml";
import { parseHeaderFooter } from "@parts/handout-master";
import { buildNotesMasterXml } from "@parts/notes-master";
import type { NotesMasterOptions, NotesLevelProperties } from "@parts/notes-master";

// ── Types ──

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

export const notesMasterDesc: CustomDescriptor<NotesMasterOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    return buildNotesMasterXml(opts);
  },

  parse(el, _ctx) {
    const options: Partial<NotesMasterOptions> = {};

    const colorMapping = parseColorMapping(findChild(el, "p:clrMap"));
    if (colorMapping) options.colorMapping = colorMapping;

    const headerFooter = parseHeaderFooter(findChild(el, "p:hf"));
    if (headerFooter) options.headerFooter = headerFooter;

    // notesStyle
    const notesStyle = findChild(el, "p:notesStyle");
    if (notesStyle) {
      const levels: NotesLevelProperties[] = [];
      for (const tag of LEVEL_TAGS) {
        const lvlEl = findChild(notesStyle, tag);
        if (lvlEl) {
          const lvl: Partial<NotesLevelProperties> = {};
          const defRPr = findChild(lvlEl, "a:defRPr");
          if (defRPr) {
            // ST_TextPoint: unqualified int is 1/100 pt; a UniversalMeasure
            // string converts to points.
            const sz = attrMeasure(defRPr, "sz");
            if (sz !== undefined) {
              lvl.fontSize =
                typeof sz === "number" ? sz / 100 : convertToPt(sz as UniversalMeasure);
            }
          }
          const marL = attrMeasure(lvlEl, "marL");
          if (marL !== undefined) lvl.marginLeft = marL as number | UniversalMeasure;
          const algn = attr(lvlEl, "algn");
          if (algn !== undefined) lvl.alignment = algn;
          levels.push(lvl as NotesLevelProperties);
        }
      }
      if (levels.length > 0) options.notesStyle = levels;
    }

    return options as NotesMasterOptions;
  },
};
