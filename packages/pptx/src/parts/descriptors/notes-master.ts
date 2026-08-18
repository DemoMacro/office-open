/**
 * Notes Master (p:notesMaster) descriptor for PPTX.
 *
 * @module
 */

import { parseColorMapping, stringifyColorMapping } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { stringifyTextListStyleTag, textListStyleDesc } from "@office-open/core/drawing";
import { findChild, stringify } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";
import { buildHfAttrs, parseHeaderFooter } from "@parts/handout-master";
import { DEFAULT_NOTES_STYLE } from "@parts/notes-master";
import type { NotesMasterOptions } from "@parts/notes-master";
import { NS } from "@parts/slide-layout";
import { buildBackgroundXml } from "@parts/slide-master";
import type { SlideChild } from "@parts/slide/slide-child";
import { SP_TREE_HEADER } from "@shared/constants";

import type { PptxWriteContext } from "../../context";
import { backgroundDesc } from "./background";
import { parseChild, stringifyChild } from "./bridge";
import { withChildId } from "./slide-master";

// ── Descriptor ──

export const notesMasterDesc: CustomDescriptor<NotesMasterOptions, PptxWriteContext> = {
  kind: "custom",

  stringify(opts, ctx) {
    const parts: string[] = [];

    parts.push(`<p:notesMaster ${NS}>`);

    // p:cSld — bg + spTree.
    parts.push("<p:cSld>");
    parts.push(buildBackgroundXml(opts.background));
    parts.push("<p:spTree>");
    parts.push(SP_TREE_HEADER);
    // Children carry explicit cNvPr ids so they never collide with the group
    // header (id 1) or ids kept from a parsed source.
    let childId = 100;
    for (const child of opts.children ?? []) {
      const xml = stringifyChild(withChildId(child, childId++), ctx);
      if (xml) parts.push(xml);
    }
    parts.push("</p:spTree>");
    if (opts.ext !== undefined) parts.push(`<p:extLst>${opts.ext}</p:extLst>`);
    parts.push("</p:cSld>");

    parts.push(stringifyColorMapping(opts.colorMapping, "p:clrMap"));
    // p:hf — optional; Office notes masters carry none by default.
    if (opts.headerFooter !== undefined) {
      parts.push(`<p:hf ${buildHfAttrs(opts.headerFooter)}/>`);
    }
    parts.push(
      stringifyTextListStyleTag("p:notesStyle", opts.notesStyle ?? DEFAULT_NOTES_STYLE, ctx),
    );

    parts.push("</p:notesMaster>");
    return parts.join("");
  },

  parse(el, ctx) {
    const result: Partial<NotesMasterOptions> = {};

    const cSld = findChild(el, "p:cSld");
    if (cSld) {
      const bg = findChild(cSld, "p:bg");
      if (bg) {
        const bgOpts = backgroundDesc.parse(bg, ctx);
        if (bgOpts && Object.keys(bgOpts).length > 0) result.background = bgOpts;
      }

      const spTree = findChild(cSld, "p:spTree");
      if (spTree) {
        const children: SlideChild[] = [];
        for (const child of spTree.elements ?? []) {
          if (child.name === "p:nvGrpSpPr" || child.name === "p:grpSpPr") continue;
          const parsed = parseChild(child, ctx);
          if (parsed !== undefined) children.push(parsed);
        }
        if (children.length > 0) result.children = children;
      }
    }

    const colorMapping = parseColorMapping(findChild(el, "p:clrMap"));
    if (colorMapping) result.colorMapping = colorMapping;

    const headerFooter = parseHeaderFooter(findChild(el, "p:hf"));
    if (headerFooter) result.headerFooter = headerFooter;

    const extLst = cSld && findChild(cSld, "p:extLst");
    if (extLst) result.ext = stringify(extLst);

    const notesStyleEl: XmlElement | undefined = findChild(el, "p:notesStyle");
    if (notesStyleEl) {
      const notesStyle = textListStyleDesc.parse(notesStyleEl, ctx);
      // An empty p:notesStyle parses to an empty list; keep the Office default.
      if (
        notesStyle.defaultParagraph ||
        notesStyle.ext !== undefined ||
        (notesStyle.levels?.length ?? 0) > 0
      ) {
        result.notesStyle = notesStyle;
      }
    }

    return result as NotesMasterOptions;
  },
};
